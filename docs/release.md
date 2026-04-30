# sheetkit release checklist

This document describes a practical release flow for `sheetkit`.

## 1. Pre-release checks (local)

From repository root:

```powershell
python -m pip install -U pip
python -m pip install -e ".[dev]"
python -m ruff check sheetkit tests scripts examples
python -m mypy sheetkit --config-file mypy.ini
python -m pytest -q
```

If all checks pass:

```powershell
python -m build
python -m twine check dist/*
```

## 2. Publish candidate to TestPyPI

Create a TestPyPI API token, then set environment variables:

```powershell
$env:TWINE_USERNAME="__token__"
$env:TWINE_PASSWORD="<testpypi-token>"
```

Upload:

```powershell
python -m twine upload --repository testpypi dist/*
```

Install and smoke-test from TestPyPI in a clean virtualenv:

```powershell
python -m venv .venv-testpypi
. .\.venv-testpypi\Scripts\Activate.ps1
python -m pip install -U pip
python -m pip install -i https://test.pypi.org/simple/ --extra-index-url https://pypi.org/simple sheetkit==0.1.0
python -c "import sheetkit; print(sheetkit.__name__)"
deactivate
```

## 3. Publish to PyPI

Create a PyPI token, then set:

```powershell
$env:TWINE_USERNAME="__token__"
$env:TWINE_PASSWORD="<pypi-token>"
```

Upload:

```powershell
python -m twine upload dist/*
```

## 4. Git release hygiene

Recommended order:

```powershell
git status
git add .
git commit -m "release: prepare v0.1.0"
git tag v0.1.0
git push origin HEAD
git push origin v0.1.0
```

## 5. Version bump policy (SemVer)

- PATCH (`x.y.Z`): bug fixes, no breaking API changes.
- MINOR (`x.Y.z`): backward-compatible new features.
- MAJOR (`X.y.z`): breaking public API changes.

Before each release, update `version` in `pyproject.toml` and mention key changes in release notes.
