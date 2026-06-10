"""
CLI subcommand handlers and parser for acct-automation.

Provides:
  - build_parser(): constructs the argparse CLI
  - cmd_generate, cmd_list, cmd_run, cmd_commit, cmd_push: subcommand handlers
  - _run_script: dynamic script loader and invoker
"""

import argparse
import importlib.util
import inspect
import re
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent.parent
SCRIPTS_DIR = PROJECT_ROOT / "scripts"


# ---------------------------------------------------------------------------
# Script discovery and dynamic loading
# ---------------------------------------------------------------------------

def get_available_scripts() -> list:
    """Return script names (without .py) from the scripts/ directory."""
    return sorted(
        p.stem for p in SCRIPTS_DIR.glob("*.py")
        if not p.name.startswith("_")
    )


def _get_script_info(script_path: Path) -> dict:
    """Extract description and argparse arguments from a script file."""
    info = {"name": script_path.stem, "description": "", "args": [], "default_task": ""}

    try:
        source = script_path.read_text(encoding="utf-8")
    except Exception:
        return info

    # Extract module docstring (first line)
    if source.startswith('"""') or source.startswith("'''"):
        quote = source[:3]
        end = source.find(quote, 3)
        if end != -1:
            doc = source[3:end].strip()
            info["description"] = doc.split("\n")[0]

    # Extract argparse add_argument calls (explicit flags)
    for match in re.finditer(
        r'add_argument\(\s*["\'](-{1,2}[\w-]+)["\'].*?help\s*=\s*["\']([^"\']+)["\']',
        source,
        re.DOTALL,
    ):
        flag, help_text = match.group(1), match.group(2)
        info["args"].append(f"{flag}: {help_text}")

    # Detect parse_task_args usage and extract default_task
    task_match = re.search(
        r'parse_task_args\(.*?default_task\s*=\s*["\'](\w+)["\']',
        source,
        re.DOTALL,
    )
    if task_match:
        info["default_task"] = task_match.group(1)

    return info


def _load_script_module(script_name: str):
    """Dynamically load a script module from scripts/ directory."""
    script_path = SCRIPTS_DIR / f"{script_name}.py"
    if not script_path.exists():
        return None

    spec = importlib.util.spec_from_file_location(
        f"scripts.{script_name}", str(script_path)
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _get_script_entry_function(module, script_name: str):
    """Get the main entry function from a script module (matches script name)."""
    func = getattr(module, script_name, None)
    if func and callable(func):
        return func

    # Fallback: find any public function that takes a single 'config' arg
    for name, obj in inspect.getmembers(module, inspect.isfunction):
        if name.startswith("_"):
            continue
        sig = inspect.signature(obj)
        params = list(sig.parameters.keys())
        if params == ["config"]:
            return obj

    return None


def run_script(script_name: str, argv: list = None):
    """Load and run a script by name, passing argv for parse_task_args."""
    if argv:
        sys.argv = [f"scripts/{script_name}.py"] + argv
    else:
        sys.argv = [f"scripts/{script_name}.py"]

    module = _load_script_module(script_name)
    if module is None:
        available = ", ".join(get_available_scripts())
        print(f"Error: script '{script_name}' not found in scripts/")
        print(f"Available: {available}")
        sys.exit(1)

    entry_fn = _get_script_entry_function(module, script_name)
    if entry_fn is None:
        print(f"Error: no entry function found in scripts/{script_name}.py")
        sys.exit(1)

    from lib.task_config import parse_task_args

    description = ""
    if module.__doc__:
        description = module.__doc__.strip().split("\n")[0]

    config = parse_task_args(
        description=description or script_name.replace("_", " ").title(),
        default_task=script_name,
        argv=argv,
    )

    entry_fn(config)


# ---------------------------------------------------------------------------
# Subcommand handlers
# ---------------------------------------------------------------------------

def cmd_generate(args):
    """Generate pre_prompt.md and write it to disk."""
    # Import from main to access introspection functions
    from main import generate_pre_prompt, OUTPUT_PATH
    content = generate_pre_prompt()
    OUTPUT_PATH.write_text(content, encoding="utf-8")
    print(f"Generated: {OUTPUT_PATH}")


def cmd_list(args):
    """Print available task scripts with descriptions and arguments."""
    print("Available task scripts:\n")

    for script_path in sorted(SCRIPTS_DIR.glob("*.py")):
        if script_path.name.startswith("_"):
            continue

        info = _get_script_info(script_path)
        print(f"  {info['name']}")
        if info["description"]:
            print(f"    {info['description']}")
        if info["default_task"]:
            print(f"    Task config: tasks/{info['default_task']}.yml (override with --task)")
            print(f"    Config values are passable as CLI args (use --help for details)")
        if info["args"]:
            print(f"    Extra args: {', '.join(info['args'])}")
        print()


def cmd_run(args):
    """Run a task script by name, or list available scripts if none specified."""
    if args.script is None:
        cmd_list(args)
        return

    script_name = args.script
    if not (SCRIPTS_DIR / f"{script_name}.py").exists():
        available = ", ".join(get_available_scripts())
        print(f"Error: script '{script_name}' not found in scripts/")
        print(f"Available: {available}")
        sys.exit(1)

    run_script(script_name, args.script_args or [])


def cmd_commit(args):
    """Stage all changes and commit with a summary + file list."""
    from lib.git_ops import get_repo, stage_all, commit, get_changed_files

    repo = get_repo()

    changed = get_changed_files(repo)
    if not changed:
        print("Nothing to commit — working tree is clean.")
        return

    stage_all(repo)

    summary = args.message
    if not summary:
        print("Changed files:")
        for f in changed:
            print(f"  - {f}")
        print()
        summary = input("Commit summary (max 80 chars): ").strip()
        if not summary:
            print("Aborted — no summary provided.")
            sys.exit(1)

    if len(summary) > 80:
        print("Warning: summary truncated to 80 chars.")
        summary = summary[:80]

    sha = commit(repo, summary)
    print(f"Committed: {sha[:8]} {summary}")


def cmd_push(args):
    """Push the current branch to remote."""
    from lib.git_ops import get_repo, push

    repo = get_repo()
    remote = args.remote or "origin"
    branch = args.branch or repo.head.shorthand

    print(f"Pushing {branch} to {remote}...")
    try:
        push(repo, remote_name=remote, branch=branch)
        print("Done.")
    except Exception as e:
        print(f"Push failed: {e}")
        sys.exit(1)


# ---------------------------------------------------------------------------
# Parser builder
# ---------------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    """Build the main argparse parser with all subcommands."""
    parser = argparse.ArgumentParser(
        prog="main.py",
        description="acct-automation CLI — generate docs or run task scripts.",
    )
    subparsers = parser.add_subparsers(dest="command")

    # generate
    sp_generate = subparsers.add_parser(
        "generate",
        help="Regenerate pre_prompt.md from template and lib/ introspection.",
    )
    sp_generate.set_defaults(func=cmd_generate)

    # list
    sp_list = subparsers.add_parser(
        "list",
        help="Show available task scripts in scripts/.",
    )
    sp_list.set_defaults(func=cmd_list)

    # run
    sp_run = subparsers.add_parser(
        "run",
        help="Execute a task script by name. If no script specified, lists available scripts.",
    )
    sp_run.add_argument(
        "script",
        nargs="?",
        default=None,
        help="Name of the script to run (without .py extension).",
    )
    sp_run.add_argument(
        "script_args",
        nargs=argparse.REMAINDER,
        help="Additional arguments passed through to the script.",
    )
    sp_run.set_defaults(func=cmd_run)

    # commit
    sp_commit = subparsers.add_parser(
        "commit",
        help="Stage all changes and commit with a summary + auto-generated file list.",
    )
    sp_commit.add_argument(
        "-m", "--message",
        default=None,
        help="Commit summary (max 80 chars). If omitted, will prompt interactively.",
    )
    sp_commit.set_defaults(func=cmd_commit)

    # push
    sp_push = subparsers.add_parser(
        "push",
        help="Push the current branch to a remote.",
    )
    sp_push.add_argument(
        "--remote",
        default=None,
        help="Remote name (default: origin).",
    )
    sp_push.add_argument(
        "--branch",
        default=None,
        help="Branch to push (default: current branch).",
    )
    sp_push.set_defaults(func=cmd_push)

    return parser
