"""
Main entry point for acct-automation.

Subcommands:
  generate   Regenerate pre_prompt.md from the template and lib/ introspection.
  list       Show available task scripts.
  run        Execute a task script by name.
  commit     Git commit with auto-generated file list.
  push       Git push using GITHUB_PAT.

Usage:
    python main.py                       # GUI task selector
    python main.py generate              # Regenerate pre_prompt.md
    python main.py list                  # List scripts
    python main.py run <script>          # Run a script
    python main.py run <script> --help   # Show CLI flags
"""

import importlib
import inspect
import os
import pkgutil
import sys
from pathlib import Path


def _check_venv() -> None:
    """Warn if not running inside a dedicated virtual environment."""
    in_venv = bool(os.environ.get("VIRTUAL_ENV"))
    conda_env = os.environ.get("CONDA_DEFAULT_ENV", "")
    in_named_conda = bool(conda_env) and conda_env.lower() != "base"
    if in_venv or in_named_conda:
        return
    env_label = f"conda '{conda_env}'" if conda_env else "system Python"
    print(
        f"WARNING: running in {env_label} — activate the project venv first.\n"
        "  PowerShell : .venv\\Scripts\\Activate.ps1\n"
        "  bash/sh    : source .venv/Scripts/activate\n"
        "  Create it  : python -m virtualenv .venv",
        file=sys.stderr,
    )


_check_venv()


PROJECT_ROOT = Path(__file__).resolve().parent
LIB_ROOT = PROJECT_ROOT / "lib"
SCRIPTS_DIR = PROJECT_ROOT / "scripts"
TEMPLATE_PATH = PROJECT_ROOT / "pre_prompt_template.md"
OUTPUT_PATH = PROJECT_ROOT / "pre_prompt.md"

# Ensure project root is on sys.path for script imports
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


# ---------------------------------------------------------------------------
# Lib introspection (used by 'generate' command)
# ---------------------------------------------------------------------------

def get_module_summary(module) -> list:
    """Return one-line summaries for each public callable in a module."""
    lines = []
    all_names = getattr(module, "__all__", None)
    members = inspect.getmembers(module)

    for name, obj in members:
        if name.startswith("_"):
            continue
        if all_names and name not in all_names:
            continue
        if not (inspect.isfunction(obj) or inspect.isclass(obj)):
            continue

        sig = ""
        try:
            sig = str(inspect.signature(obj))
        except (ValueError, TypeError):
            pass

        doc = inspect.getdoc(obj)
        short_doc = doc.split("\n")[0] if doc else ""

        kind = "class" if inspect.isclass(obj) else "fn"
        lines.append(f"  {kind} {name}{sig}  # {short_doc}" if short_doc else f"  {kind} {name}{sig}")

    return lines


def generate_lib_listing() -> str:
    """Walk lib/ and produce a token-efficient listing of available modules."""
    output_lines = ["## Available lib/ modules", ""]

    for _importer, modname, _ispkg in pkgutil.walk_packages(
        path=[str(LIB_ROOT)], prefix="lib."
    ):
        if "__pycache__" in modname or modname.split(".")[-1].startswith("_"):
            continue

        try:
            module = importlib.import_module(modname)
        except Exception:
            continue

        mod_doc = inspect.getdoc(module)
        short_mod_doc = mod_doc.split("\n")[0] if mod_doc else ""

        header = f"### {modname}"
        if short_mod_doc:
            header += f"  — {short_mod_doc}"
        output_lines.append(header)

        members = get_module_summary(module)
        if members:
            output_lines.append("```")
            output_lines.extend(members)
            output_lines.append("```")

        output_lines.append("")

    return "\n".join(output_lines)


def generate_scripts_listing() -> str:
    """List available task scripts with their one-line descriptions."""
    output_lines = ["## Available task scripts (scripts/)", ""]
    output_lines.append("| Script | Description |")
    output_lines.append("|--------|-------------|")

    for script_path in sorted(SCRIPTS_DIR.glob("*.py")):
        if script_path.name.startswith("_"):
            continue

        description = ""
        try:
            source = script_path.read_text(encoding="utf-8")
            if source.startswith('"""') or source.startswith("'''"):
                quote = source[:3]
                end = source.find(quote, 3)
                if end != -1:
                    doc = source[3:end].strip()
                    description = doc.split("\n")[0]
        except Exception:
            pass

        name = script_path.stem
        output_lines.append(f"| `{name}` | {description} |")

    output_lines.append("")
    output_lines.append("Run with: `python main.py run <script_name>`")
    output_lines.append("")
    return "\n".join(output_lines)


def generate_pre_prompt() -> str:
    """Build the full pre_prompt.md content from template + introspected data."""
    lib_listing = generate_lib_listing()
    scripts_listing = generate_scripts_listing()

    if TEMPLATE_PATH.exists():
        template = TEMPLATE_PATH.read_text(encoding="utf-8")
    else:
        template = "# acct-automation\n\n{{SCRIPTS_LISTING}}\n\n{{LIB_LISTING}}\n"

    content = template.replace("{{LIB_LISTING}}", lib_listing)
    content = content.replace("{{SCRIPTS_LISTING}}", scripts_listing)
    return content


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    from lib.cli import build_parser, run_script

    parser = build_parser()
    args = parser.parse_args()

    if args.command is None:
        # Default: open GUI task selector
        from lib.gui_config import gui_select_task
        while True:
            selected = gui_select_task()
            if selected is None:
                print("No task selected.")
                sys.exit(0)
            if selected == "__restart__":
                # After a pull, restart the selector to pick up new scripts
                continue
            break
        run_script(selected)
    else:
        args.func(args)


if __name__ == "__main__":
    main()
