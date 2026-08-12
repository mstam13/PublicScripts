# Contributing to PublicScripts

Thanks for your interest in contributing! This repository follows a simple
branch-and-pull-request workflow to keep `main` stable and reviewable.

## Getting Started Locally

1. Clone the repository:

   ```powershell
   git clone https://github.com/mstam13/PublicScripts.git
   cd PublicScripts
   ```

2. Open the folder in VS Code (recommended) — the `.vscode` and
   `.github/instructions` settings provide linting and Copilot guidance
   specific to this repo.
3. Create a topic branch off `main` for your change (see
   [Branching & Pull Requests](#branching--pull-requests) below).

## Requirements

- **PowerShell** — PowerShell 5.1 or PowerShell 7+, depending on the script
  (see each script's `.md` file for its minimum version).
  - [PSScriptAnalyzer](https://github.com/PowerShell/PSScriptAnalyzer) —
    required to lint all `.ps1` changes before opening a pull request.
  - [ImportExcel](https://github.com/dfinke/ImportExcel) — required by
    several scripts for `.xlsx` export.
  - RSAT modules (`ActiveDirectory`, `GroupPolicy`) — required by scripts
    that query Active Directory or Group Policy.
- **SQL scripts** — SQL Server 2016 or later; SSMS or `sqlcmd`.
- **Markdown** — [markdownlint](https://github.com/DavidAnson/markdownlint)
  (CLI or the VS Code extension) to lint any `.md` changes.

See the root [README.md](README.md#requirements) for the full,
per-script requirements matrix.

## Making Changes

- Follow the existing script structure: comment-based help
  (`.SYNOPSIS`, `.DESCRIPTION`, `.PARAMETER`, `.OUTPUTS`, `.EXAMPLE`,
  `.NOTES`), `#region`/`#endregion` blocks, and `Verb-Noun` naming.
- Every script change must be paired with an update to its accompanying
  `.md` file (synopsis, parameters, outputs, examples, version history).
- Run `Invoke-ScriptAnalyzer` against any `.ps1` file you change and resolve
  all findings.
- Run `markdownlint` against any `.md` file you change and resolve all
  findings.
- If your change affects the folder listing, versions, or requirements,
  update the root [README.md](README.md) accordingly.

## Branching & Pull Requests

- **Never commit directly to `main`.** All changes must go through a
  feature branch and a pull request.
- Branch names should be short and descriptive, e.g.
  `feature/add-x`, `fix/y-bug`, `docs/update-readme`.
- Open a pull request against `main` once your branch is ready. Ensure:
  - PSScriptAnalyzer and markdownlint findings are resolved.
  - Documentation (`.md` files, README, wiki if applicable) is updated.
  - The PR description explains the change and its motivation.
- Pull requests require review and approval from a code owner (see
  [CODEOWNERS](.github/CODEOWNERS)) before merging.

## Code of Conduct

This project adheres to a [Code of Conduct](CODE_OF_CONDUCT.md). By
participating, you are expected to uphold it.
