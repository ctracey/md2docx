# Environment setup

## Prerequisites

- [pandoc](https://pandoc.org/) ≥ 3.11 — `brew install pandoc`
- [uv](https://docs.astral.sh/uv/) — `brew install uv`

No Python packages are needed at runtime. Python itself is managed by uv.

## Development setup

Python 3.13 is pinned in `.python-version`. uv reads this automatically.

```
# Install uv (once, system-wide)
brew install uv

# Install Python 3.13 and project dependencies
uv sync
```

`uv sync` creates a `.venv`, installs the pinned Python version if needed, and installs all dev dependencies from `uv.lock`. No separate pip or virtualenv step required.

## Running the tool locally

During development, use `uv run` to execute against local source without installing:

```
uv run md2docx <content.md> <style-guide.docx> <output.docx>
```

This always reflects your current working-tree changes. The globally installed tool (`uv tool install .`) is a snapshot — changes to source are not picked up until you reinstall.

## Running tests

```
uv run pytest tests/ -v
```

> **Note on test inspection:** The template embeds DM Sans fonts whose MIME
> types pandoc does not register in `[Content_Types].xml`. Tests therefore
> inspect OOXML directly via `zipfile` + `xml.etree.ElementTree` rather than
> using python-docx (which fails to open such files). This is a test-only
> concern — the generated DOCX opens correctly in Word and Google Docs.
