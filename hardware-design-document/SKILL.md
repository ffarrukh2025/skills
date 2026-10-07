---
name: hardware-design-document
description: Creates and updates professional Word architecture documents for RTL and hardware modules, including polished block diagrams, native editable equations, interfaces, pipelines, state machines, latency, throughput, and timing explanations. Use when designing or documenting a new RTL module, encoder, decoder, datapath, protocol block, or FPGA architecture.
Author: Fasih ud Din Farrukh
copyright: Copyright (c) 2026 Altera Corporation. All rights reserved.
---

# Hardware Design Document

**Copyright:** Copyright (c) 2026 Altera Corporation. All rights reserved.

Create and update professional Word architecture documents for RTL and hardware modules. Prefer updating an existing `.docx` over rewriting unrelated content.

Read [DOCUMENT_TEMPLATE.md](DOCUMENT_TEMPLATE.md) for section order and JSON spec.
Read [DIAGRAM_STYLE.md](DIAGRAM_STYLE.md) before drawing figures.

## When to use

Apply this skill for prompts such as:

- “Create a design document for this RTL module.”
- “Update the architecture document after these pipeline changes.”
- “Document the syndrome calculation with block diagrams and editable equations.”
- “Add a clock-by-clock data-flow diagram.”
- “Compare the implemented architecture with the specification.”

## Hard rules

- Read the RTL, specifications, interfaces, parameters, state machines, pipelines, and timing constraints.
- Confirm uncertain architectural details instead of inventing them.
- Draw polished Visio-style color-coded diagrams as embedded images (header-bar shapes, orthogonal connectors).
- Generate equations as native Word OMML equations so they remain editable.
- Keep all explanatory text and tables editable.
- Validate the `.docx` by opening it in Word and checking the diagram and equation counts.
- Update the same document as the RTL design evolves, rather than recreating unrelated content.
- Do not dump ASCII art into Word as a substitute for figures.
- Do not paste LaTeX source as plain text in place of OMML.
- Do not claim latency, throughput, buffering, or Fmax unless they follow from RTL, constraints, or a stated assumption the user confirmed.
- Do not prefix list items with Unicode bullets in JSON. `build_docx.ps1` applies Word list bullets.
- After build, the Contents page must list headings. It must not show `Error! Bookmark not defined.`
- Do **not** invent or hardcode an author name for shared use. Author resolution order:
  1. Non-empty `authors` in `doc_spec.json`
  2. Else non-empty `authors` in personal [config.local.json](config.local.json) next to this skill (not shared with the team)
  3. Else leave author empty (omit the Author line; revision-history author cell blank)
- Every generated `.docx` must show Altera copyright on the title page and in the footer: `Copyright (c) <year> Altera Corporation. All rights reserved.` Use the document year (from `date` or the current year).

## Produce

Every new document must include:

- title page
- contents list
- Executive overview
- Interface and parameter tables
- Top-level block diagram
- Detailed datapath diagrams
- Pipeline and clock-cycle diagrams
- State-machine diagrams
- Equations and derivations
- Throughput, latency, buffering, and timing explanations
- Reset, error handling, and design limitations

Omit a section only when the module genuinely has no such content (for example no FSM). State the omission in the overview; do not invent a dummy machine.

## Workflow

Copy this checklist and track progress:

```
Task Progress:
- [ ] 1. Locate RTL, spec, constraints, and any existing .docx
- [ ] 2. Extract facts; list open questions
- [ ] 3. Confirm unknowns with the user
- [ ] 4. Write or patch docs/<module>/doc_spec.json
- [ ] 5. Render PNG figures per DIAGRAM_STYLE.md
- [ ] 6. Build Word doc with scripts/build_docx.ps1
- [ ] 7. Validate with scripts/validate_docx.ps1
- [ ] 8. Fix failures and rebuild until validation passes
```

### 1. Research (do not skip)

From RTL and related files collect:

- Module name, hierarchy, and file list
- Parameters, localparams, generated widths
- Ports: direction, width, clock/reset domain, handshake
- Clock and reset domains; CDC if present
- Pipeline stages, valid/ready/enable, stall, flush
- Memories/FIFOs: depth, width, almost-full policy
- FSMs: states, transitions, outputs
- Datapath equations (GF arithmetic, packing, CRC, mapping)
- Timing: initiation interval, latency, backpressure, Fmax constraints
- Error, ECC/FEC, overflow, and recovery behavior

If a spec and RTL disagree, document both and mark **Implemented** vs **Specified**. Do not silently pick one.

### 2. Confirm unknowns

Ask the user before writing when any of these are missing or ambiguous:

- Clock period / target Fmax
- Intended latency vs incidental RTL delay
- Reset polarity and synchronous vs asynchronous
- Whether a block is instantiated or only planned
- External protocol meaning not named in RTL comments

### 3. Create vs update

**New document:** write `doc_spec.json` and figures, then build.

**Existing document:** diff RTL against the last documented revision. Change only affected headings (bookmark names in the spec). Keep title, unchanged sections, figure numbers for untouched figures, and prior revision history. Add a revision row.

Default output path: `docs/<module>/<Module>_Architecture.docx`

Keep sources next to the doc:

```
docs/<module>/
├── doc_spec.json
├── figures/
│   └── *.png
└── <Module>_Architecture.docx
```

### 4. Figures

Render PNGs at 200 DPI with [scripts/VisioStyle.ps1](scripts/VisioStyle.ps1) (PowerShell, no Python). Follow [DIAGRAM_STYLE.md](DIAGRAM_STYLE.md): header-bar shapes, orthogonal connectors, port dots, connector labels, containers.

Embed them; never leave “see figure” with a missing file.

Required figure set when the RTL has the corresponding structure:

| Figure | Typical filename |
|--------|------------------|
| Top-level block | `fig_top.png` |
| Datapath / detailed blocks | `fig_datapath_*.png` |
| Pipeline / occupancy | `fig_pipeline.png` |
| Clock-by-clock flow | `fig_cycle.png` |
| FSM | `fig_fsm_*.png` |

### 5. Equations

Put every derivation in spec `type: equation` entries using Word **UnicodeMath** (linear format). `build_docx.ps1` inserts `OMath` and calls `BuildUp()` so Word stores native OMML.

Examples:

- `S_j=\sum_(i=0)^(n-1) r_i \alpha^(j i)`
- `T_clk=1/F_max`
- `Latency=N_pipe+N_mem`
- `GF(2^10):\alpha^1023=1`

Number displayed equations. Keep variable names identical to RTL signals when they are the same quantity.

### 6. Build

Run from the module docs directory (or pass absolute paths):

```powershell
powershell -NoProfile -File "<skill>/scripts/build_docx.ps1" `
  -SpecPath ".\doc_spec.json" `
  -OutPath ".\<Module>_Architecture.docx" `
  -AssetDir ".\figures"
```

`<skill>` is this skill directory. Recreate the `.docx` from `doc_spec.json` on each full build so OMML and styles stay consistent. That is still an in-place update of the same file when the JSON was patched, not a new unrelated document.

### 7. Validate

```powershell
powershell -NoProfile -File "<skill>/scripts/validate_docx.ps1" `
  -DocPath ".\<Module>_Architecture.docx" `
  -ExpectedFigures <n> `
  -ExpectedEquations <n>
```

Must pass: Word opens the file, figure count matches embedded PNGs, equation count matches `OMath` objects, and body text is not pictures of paragraphs. If validation fails, fix spec/figures and rebuild. Do not hand the user an unvalidated file.

## Writing style

- Precise, engineering tone. Name signals, parameters, and modules as in RTL.
- Tables for ports, parameters, timing, and address maps.
- Short paragraphs after each figure saying what the figure shows and what the RTL does.
- Revision history on the title page area (JSON `revisionHistory`).
- For team-shared specs, leave `"authors": []` (or omit `authors`). Individuals put their name only in personal `config.local.json` or in their own `doc_spec.json`.
- Always set Altera copyright (builder default if omitted).

## Personal author (do not share)

Copy [config.local.example.json](config.local.example.json) to `config.local.json` and set your name. Keep `config.local.json` on your machine only when sharing the skill folder.

## Additional resources

- [DOCUMENT_TEMPLATE.md](DOCUMENT_TEMPLATE.md) — required sections and `doc_spec.json` schema
- [DIAGRAM_STYLE.md](DIAGRAM_STYLE.md) — Visio-like layout; uses VisioStyle.ps1
- [config.local.example.json](config.local.example.json) — template for personal author override
- [scripts/VisioStyle.ps1](scripts/VisioStyle.ps1) — PowerShell drawing library (dot-source)
- [scripts/build_docx.ps1](scripts/build_docx.ps1) — Word COM builder (execute)
- [scripts/validate_docx.ps1](scripts/validate_docx.ps1) — Word COM validator (execute)
