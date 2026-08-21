# copy_py_to_md.py

```python
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent
OUTPUT_DIR = PROJECT_ROOT / "copy"
SOURCE_EXTENSIONS = {
    ".py": "python",
    ".json": "json",
}
SKIP_DIRS = {
    ".git",
    ".idea",
    ".mypy_cache",
    ".pytest_cache",
    ".ruff_cache",
    ".tox",
    ".venv",
    "__pycache__",
    "copy",
    "venv",
}


def should_skip(path: Path) -> bool:
    return any(part in SKIP_DIRS for part in path.relative_to(PROJECT_ROOT).parts)


def iter_source_files() -> list[Path]:
    return sorted(
        path
        for extension in SOURCE_EXTENSIONS
        for path in PROJECT_ROOT.rglob(f"*{extension}")
        if path.is_file() and not should_skip(path)
    )


def markdown_fence(code: str) -> str:
    longest_backticks = 0
    current = 0
    for char in code:
        if char == "`":
            current += 1
            longest_backticks = max(longest_backticks, current)
        else:
            current = 0
    return "`" * max(3, longest_backticks + 1)


def write_markdown(source_path: Path, md_path: Path, code: str) -> None:
    relative_source = source_path.relative_to(PROJECT_ROOT).as_posix()
    fence = markdown_fence(code)
    language = SOURCE_EXTENSIONS.get(source_path.suffix, "")
    md_path.write_text(
        f"# {relative_source}\n\n{fence}{language}\n{code.rstrip()}\n{fence}\n",
        encoding="utf-8",
    )


def clean_output_dir() -> int:
    if not OUTPUT_DIR.exists():
        return 0

    deleted_count = 0
    for path in OUTPUT_DIR.rglob("*.md"):
        if path.is_file():
            path.unlink()
            deleted_count += 1
    return deleted_count


def main() -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)
    deleted_count = clean_output_dir()

    generated_count = 0
    for source_path in iter_source_files():
        relative_path = source_path.relative_to(PROJECT_ROOT)
        md_path = (OUTPUT_DIR / relative_path).with_suffix(".md")

        md_path.parent.mkdir(parents=True, exist_ok=True)
        code = source_path.read_text(encoding="utf-8")
        write_markdown(source_path, md_path, code)
        generated_count += 1

    print(
        f"Generated {generated_count} Markdown files in {OUTPUT_DIR}; "
        f"deleted {deleted_count} old Markdown files"
    )


if __name__ == "__main__":
    main()
```
