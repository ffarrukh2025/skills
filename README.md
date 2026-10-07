# hardware-design-document

Cursor Agent skill that creates and updates professional **Word architecture documents** for RTL / FPGA / hardware modules.

Copyright (c) 2026 Altera Corporation. All rights reserved.

## What it produces

A `.docx` with:

- Title page, revision history, and contents list
- Executive overview
- Interface and parameter tables
- Visio-style block / datapath / pipeline / FSM diagrams (embedded PNG)
- Native Word editable equations (OMML)
- Throughput, latency, buffering, and timing notes
- Reset, errors, and design limitations

Typical output layout:

```text
docs/<module>/
├── doc_spec.json          # document source (text, tables, equations)
├── figures/
│   ├── make_figures.ps1   # optional local figure script
│   └── fig_*.png
└── <Module>_Architecture.docx
```

## Requirements

- Microsoft Word (desktop, for COM automation)
- PowerShell
- Cursor Agent mode (recommended)

No Python is required.

## Install

### Personal skill (one user)

Copy or clone this folder to:

```text
%USERPROFILE%\.cursor\skills\hardware-design-document\
```

### Shared / team skill

Place the same folder under a project:

```text
.cursor/skills/hardware-design-document/
```

or share a copy of this directory. Do **not** share personal `config.local.json` if it contains your name.

## Set your author name (optional)

Author is **empty by default** so the skill is safe to share.

To put your name on documents you generate:

1. Copy `config.local.example.json` to `config.local.json` next to this README.
2. Edit:

```json
{
  "authors": ["Your Full Name"]
}
```

Author resolution order:

1. Non-empty `authors` in `doc_spec.json`
2. Else `authors` in personal `config.local.json`
3. Else no Author line

Altera copyright is always added on the title page and footer.

## Use in Cursor

In Agent mode, point at RTL (open files or a module path) and ask, for example:

- `Using the hardware-design-document skill, create a design document for this RTL module.`
- `Update the architecture document after these pipeline changes.`
- `Document the syndrome calculation with block diagrams and editable equations.`
- `Add a clock-by-clock data-flow diagram.`
- `Compare the implemented architecture with the specification.`

The agent should read RTL/specs, ask about unclear items, write `doc_spec.json`, draw figures, build the Word file, and validate it.

## Manual build (optional)

If `doc_spec.json` and PNGs already exist:

```powershell
$skill = "$env:USERPROFILE\.cursor\skills\hardware-design-document"

powershell -NoProfile -File "$skill\scripts\build_docx.ps1" `
  -SpecPath ".\doc_spec.json" `
  -OutPath ".\MyModule_Architecture.docx" `
  -AssetDir ".\figures"

powershell -NoProfile -File "$skill\scripts\validate_docx.ps1" `
  -DocPath ".\MyModule_Architecture.docx" `
  -SpecPath ".\doc_spec.json"
```

Close the `.docx` in Word before rebuilding, or the file may be locked.

## Figures (Visio-style)

Use `scripts/VisioStyle.ps1` from a local figure script:

```powershell
. "$env:USERPROFILE\.cursor\skills\hardware-design-document\scripts\VisioStyle.ps1"
New-VsCanvas 2200 1400
Add-VsTitle "Module - top view"
# Add-VsShape / Add-VsConnector / ...
Save-VsCanvas ".\figures\fig_top.png"
```

See [DIAGRAM_STYLE.md](DIAGRAM_STYLE.md) for colors, shape kinds, and layout rules.

## Folder contents

| Path | Purpose |
|------|---------|
| `SKILL.md` | Agent instructions (auto-loaded by Cursor) |
| `DOCUMENT_TEMPLATE.md` | Section order and `doc_spec.json` schema |
| `DIAGRAM_STYLE.md` | Visio-like diagram conventions |
| `config.local.example.json` | Template for personal author |
| `config.local.json` | Personal author override (do not share) |
| `scripts/VisioStyle.ps1` | PNG drawing library |
| `scripts/build_docx.ps1` | Build Word document from JSON + figures |
| `scripts/validate_docx.ps1` | Open in Word and check figures / equations / TOC |

## Update workflow

When RTL changes:

1. Patch only the affected parts of `doc_spec.json` and the matching PNGs.
2. Bump `revision` and add a revision-history row.
3. Rebuild the **same** `.docx` path.
4. Run validation again.

## Notes

- Prefer facts from RTL and specs; do not invent latency, Fmax, or architecture details.
- Keep equations as Word OMML (via UnicodeMath in `doc_spec.json`), not plain LaTeX text.
- Contents must list real headings; it must not show `Error! Bookmark not defined.`
