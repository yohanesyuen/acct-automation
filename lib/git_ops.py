"""
Git operations module using pygit2.

Provides functions to stage, commit, and push changes from the project
repository via the main.py CLI.
"""

import pygit2
from pathlib import Path
from typing import List, Optional


PROJECT_ROOT = Path(__file__).resolve().parent.parent


def get_repo() -> pygit2.Repository:
    """Open the project git repository."""
    return pygit2.Repository(str(PROJECT_ROOT))


def get_changed_files(repo: pygit2.Repository) -> List[str]:
    """
    Return a list of changed file paths (staged + unstaged + untracked).

    Returns:
        List of relative file paths that have changes.
    """
    status = repo.status()
    return sorted(path for path, flags in status.items() if flags != pygit2.GIT_STATUS_CURRENT)


def get_staged_files(repo: pygit2.Repository) -> List[str]:
    """
    Return a list of currently staged file paths.

    Returns:
        List of relative file paths that are staged for commit.
    """
    status = repo.status()
    staged_flags = (
        pygit2.GIT_STATUS_INDEX_NEW
        | pygit2.GIT_STATUS_INDEX_MODIFIED
        | pygit2.GIT_STATUS_INDEX_DELETED
        | pygit2.GIT_STATUS_INDEX_RENAMED
        | pygit2.GIT_STATUS_INDEX_TYPECHANGE
    )
    return sorted(
        path for path, flags in status.items()
        if flags & staged_flags
    )


def stage_all(repo: pygit2.Repository) -> List[str]:
    """
    Stage all changed and untracked files.

    Returns:
        List of file paths that were staged.
    """
    changed = get_changed_files(repo)
    if not changed:
        return []

    repo.index.read()
    repo.index.add_all()
    repo.index.write()
    return changed


def commit(repo: pygit2.Repository, summary: str) -> str:
    """
    Create a commit with the given summary and auto-generated file list body.

    The summary is truncated to 80 characters. The body lists all staged files.

    Args:
        repo: The pygit2 Repository object.
        summary: Short description of the change (max 80 chars).

    Returns:
        The hex string of the new commit SHA.

    Raises:
        ValueError: If no files are staged or summary is empty.
    """
    if not summary.strip():
        raise ValueError("Commit summary cannot be empty.")

    # Truncate summary to 80 chars
    summary = summary.strip()[:80]

    staged = get_staged_files(repo)
    if not staged:
        raise ValueError("No staged changes to commit.")

    # Build commit message: summary + blank line + file list
    body_lines = ["", "Changed files:"] + [f"  - {f}" for f in staged]
    message = summary + "\n" + "\n".join(body_lines)

    # Get signature from git config
    config = repo.config
    try:
        name = config["user.name"]
        email = config["user.email"]
    except KeyError:
        raise ValueError(
            "Git user.name and user.email must be configured. "
            "Run: git config user.name 'Your Name' && git config user.email 'you@example.com'"
        )

    sig = pygit2.Signature(name, email)

    # Write the index tree and create commit
    repo.index.read()
    tree = repo.index.write_tree()

    # Determine parent(s)
    try:
        parents = [repo.head.target]
    except pygit2.GitError:
        parents = []

    oid = repo.create_commit("HEAD", sig, sig, message, tree, parents)
    return str(oid)


def _load_github_pat() -> Optional[str]:
    """
    Load GITHUB_PAT from environment or .env file at project root.

    Returns:
        The token string, or None if not found.
    """
    import os

    # Check environment first
    pat = os.environ.get("GITHUB_PAT")
    if pat:
        return pat

    # Try loading from .env file
    env_path = PROJECT_ROOT / ".env"
    if env_path.exists():
        try:
            for line in env_path.read_text(encoding="utf-8").splitlines():
                line = line.strip()
                if line.startswith("#") or "=" not in line:
                    continue
                key, _, value = line.partition("=")
                key = key.strip()
                value = value.strip().strip('"').strip("'")
                if key == "GITHUB_PAT":
                    return value
        except Exception:
            pass

    return None


def _get_credentials_callbacks():
    """Build pygit2 RemoteCallbacks with PAT if available."""
    pat = _load_github_pat()
    if pat:
        return pygit2.RemoteCallbacks(
            credentials=pygit2.UserPass("x-access-token", pat)
        )
    return pygit2.RemoteCallbacks()


def fetch(repo: pygit2.Repository, remote_name: str = "origin") -> None:
    """
    Fetch latest refs from a remote.

    Args:
        repo: The pygit2 Repository object.
        remote_name: Name of the remote (default: "origin").
    """
    try:
        remote = repo.remotes[remote_name]
    except KeyError:
        return  # No remote configured — skip silently

    callbacks = _get_credentials_callbacks()
    remote.fetch(callbacks=callbacks)


def get_commits_behind(repo: pygit2.Repository, remote_name: str = "origin") -> int:
    """
    Count how many commits the local branch is behind the remote tracking branch.

    Args:
        repo: The pygit2 Repository object.
        remote_name: Name of the remote (default: "origin").

    Returns:
        Number of commits behind, or 0 if up-to-date or no tracking branch.
    """
    try:
        branch = repo.head.shorthand
        local_oid = repo.head.target
        remote_ref = f"refs/remotes/{remote_name}/{branch}"
        remote_oid = repo.references[remote_ref].target
    except (KeyError, pygit2.GitError):
        return 0

    behind, _ahead = repo.ahead_behind(local_oid, remote_oid)
    # ahead_behind returns (ahead, behind) from local perspective
    # We want "behind" = how many commits remote has that we don't
    _, behind_count = repo.ahead_behind(local_oid, remote_oid)
    return behind_count


def has_uncommitted_changes(repo: pygit2.Repository) -> bool:
    """Check if there are any uncommitted changes (staged, modified, or untracked)."""
    status = repo.status()
    return any(flags != pygit2.GIT_STATUS_CURRENT for flags in status.values())


def stash(repo: pygit2.Repository) -> Optional[str]:
    """
    Stash all uncommitted changes.

    Returns:
        The stash OID as hex string, or None if nothing to stash.
    """
    if not has_uncommitted_changes(repo):
        return None

    config = repo.config
    try:
        name = config["user.name"]
        email = config["user.email"]
    except KeyError:
        name = "acct-automation"
        email = "automation@local"

    sig = pygit2.Signature(name, email)
    oid = repo.stash(sig, message="Auto-stash before pull")
    return str(oid)


def stash_pop(repo: pygit2.Repository) -> bool:
    """
    Pop the most recent stash entry.

    If the pop fails (e.g. conflicts), drops the stash to discard the
    local changes and resets the working directory to a clean state.

    Returns:
        True if pop succeeded, False if stash was dropped due to failure.
    """
    try:
        repo.stash_pop()
        return True
    except (KeyError, pygit2.GitError):
        # Pop failed — drop the stash (discard local changes)
        try:
            repo.stash_drop()
        except (KeyError, pygit2.GitError):
            pass
        # Reset working directory to HEAD to ensure clean state
        try:
            repo.checkout_head(strategy=pygit2.GIT_CHECKOUT_FORCE)
        except pygit2.GitError:
            pass
        return False


def pull(repo: pygit2.Repository, remote_name: str = "origin") -> dict:
    """
    Pull (fetch + fast-forward merge) from the remote tracking branch.

    Automatically stashes uncommitted changes before pulling and pops
    them after.

    Args:
        repo: The pygit2 Repository object.
        remote_name: Name of the remote (default: "origin").

    Returns:
        Dict with keys:
          - "updated": True if branch was updated
          - "changed_files": list of file paths changed in the pull
    """
    result = {"updated": False, "changed_files": []}
    branch = repo.head.shorthand

    # Stash if needed
    had_changes = has_uncommitted_changes(repo)
    if had_changes:
        stash(repo)

    try:
        # Fetch
        fetch(repo, remote_name)

        # Find remote tracking ref
        remote_ref = f"refs/remotes/{remote_name}/{branch}"
        try:
            remote_oid = repo.references[remote_ref].target
        except KeyError:
            return result

        local_oid = repo.head.target
        if local_oid == remote_oid:
            return result

        # Determine which files changed between local and remote
        try:
            diff = repo.diff(repo.get(local_oid), repo.get(remote_oid))
            result["changed_files"] = [patch.delta.new_file.path for patch in diff]
        except Exception:
            pass

        # Fast-forward merge
        repo.checkout_tree(repo.get(remote_oid))
        ref = repo.references.get(f"refs/heads/{branch}")
        if ref:
            ref.set_target(remote_oid)
        repo.head.set_target(remote_oid)

        result["updated"] = True
        return result
    finally:
        # Pop stash if we stashed — if it fails, discard local changes
        if had_changes:
            success = stash_pop(repo)
            if not success:
                print("  [git] Stash pop failed — local changes discarded.")


def push(repo: pygit2.Repository, remote_name: str = "origin", branch: Optional[str] = None) -> None:
    """
    Push the current branch to a remote.

    Uses GITHUB_PAT from the environment or .env file for HTTPS
    authentication if available.

    Args:
        repo: The pygit2 Repository object.
        remote_name: Name of the remote (default: "origin").
        branch: Branch name to push. If None, uses current HEAD branch.

    Raises:
        ValueError: If the remote doesn't exist or push fails.
    """
    try:
        remote = repo.remotes[remote_name]
    except KeyError:
        raise ValueError(f"Remote '{remote_name}' not found.")

    if branch is None:
        branch = repo.head.shorthand

    refspec = f"refs/heads/{branch}:refs/heads/{branch}"

    callbacks = _get_credentials_callbacks()
    remote.push([refspec], callbacks=callbacks)
