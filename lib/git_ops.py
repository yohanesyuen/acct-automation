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

    # Build credentials callback using PAT if available
    pat = _load_github_pat()
    if pat:
        callbacks = pygit2.RemoteCallbacks(
            credentials=pygit2.UserPass("x-access-token", pat)
        )
    else:
        callbacks = pygit2.RemoteCallbacks()

    remote.push([refspec], callbacks=callbacks)
