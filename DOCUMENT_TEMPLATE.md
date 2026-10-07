# Document template and JSON spec
# Author: Fasih ud Din Farrukh

**Copyright:** Copyright (c) 2026 Altera Corporation. All rights reserved.

Output is a `.docx` built from `doc_spec.json`. Body text, headings, tables, and captions stay as Word text. Figures are PNG. Equations are OMML via Word `OMath.BuildUp()`.

## Required outline

Use this heading order. Bookmark names must match `id` values so later updates can target a section.

1. **Title page** — from `title`, `subtitle`, `authors`, `copyright`, `revision`, `date`, `classification` (optional)
2. **Revision history** — table from `revisionHistory`
3. **Contents list** — Word TOC field, updated on build
4. **1 Executive overview** (`id`: `executive-overview`)
   - Purpose, where the module sits, key throughputs/latencies
   - Implemented vs specified if they differ
5. **2 Interfaces and parameters** (`id`: `interfaces`)
   - Clock/reset table
   - Port table (name, dir, width, domain, description)
   - Parameter table (name, default, range, effect)
6. **3 Top-level architecture** (`id`: `top-level`)
   - `fig_top.png` plus narrative
7. **4 Datapath** (`id`: `datapath`)
   - Detailed block diagrams and data formatting
8. **5 Pipeline and timing** (`id`: `pipeline`)
   - Stage table, clock-cycle diagram, initiation interval, stall/backpressure
9. **6 Control and state machines** (`id`: `control`)
   - FSM diagrams and transition tables
10. **7 Equations and derivations** (`id`: `equations`)
11. **8 Throughput, latency, and buffering** (`id`: `performance`)
12. **9 Reset, errors, and limitations** (`id`: `reset-errors`)
13. **10 Specification comparison** (`id`: `spec-compare`) — optional; required when a spec file exists
14. **11 References** (`id`: `references`) — RTL paths, specs, standards

## `doc_spec.json` schema

```json
{
  "title": "RS(544,514) FEC Decoder Architecture",
  "subtitle": "Link-layer receive datapath",
  "authors": [],
  "copyright": "Copyright (c) 2026 Altera Corporation. All rights reserved.",
  "revision": "A",
  "date": "2026-10-05",
  "classification": "Internal",
  "revisionHistory": [
    {
      "rev": "A",
      "date": "2026-10-05",
      "author": "",
      "description": "Initial architecture from RTL"
    }
  ],
  "sections": []
}
```

Author and copyright defaults:

| Field | Shared default | Personal override |
|-------|----------------|-------------------|
| `authors` | empty (`[]` or omit) | non-empty `authors` in `doc_spec.json`, else `config.local.json` |
| revision-history `author` | empty | same author resolution as above when the cell is blank |
| `copyright` | `Copyright (c) <year> Altera Corporation. All rights reserved.` | optional `copyright` string in `doc_spec.json` |

Do not put a team member’s name in the shared skill files. Each user who wants a default name creates `config.local.json` (see `config.local.example.json`).

### Block types in `sections`

| `type` | Fields | Notes |
|--------|--------|--------|
| `heading` | `level` (1–3), `text`, `id` (required for level 1) | Bookmark on level 1 |
| `para` | `text` | Editable body. Use `\n` for paragraph breaks |
| `bullets` | `items` (string array) | |
| `table` | `headers`, `rows` (array of string arrays), `caption` optional | Header row shaded |
| `figure` | `file` (name under AssetDir), `caption`, `widthIn` optional | Default width 6.5 in |
| `equation` | `math`, `number` optional, `caption` optional | UnicodeMath linear format |
| `note` | `text` | Callout for assumptions or open items |

Unknown `type` values must fail the build.

### Example section fragment

```json
{
  "type": "heading",
  "level": 1,
  "id": "equations",
  "text": "7 Equations and derivations"
},
{
  "type": "para",
  "text": "Syndrome j is the evaluation of the received polynomial at α^j."
},
{
  "type": "equation",
  "number": "1",
  "math": "S_j=r(α^j)=∑_(i=0)^(n-1) r_i α^(j i)",
  "caption": "Syndrome definition"
},
{
  "type": "figure",
  "file": "fig_syndrome.png",
  "caption": "Figure 4. Syndrome datapath, one lane"
}
```

### UnicodeMath (linear) conventions

Word converts these with `BuildUp()`:

| Meaning | Linear form |
|---------|-------------|
| Subscript | `S_j`, `clk_tx` |
| Superscript | `α^j`, `2^10` |
| Sub and super | `α_i^j` |
| Fraction | `(T_setup+T_co)/T_clk` |
| Sum | `∑_(i=0)^(n-1) x_i` or `\sum_(i=0)^(n-1) x_i` |
| Product | `∏_(k=1)^t (x-α^k)` |
| Square root | `√(1+ε)` |
| Greek | `α β γ δ ε λ σ ω π` (paste Unicode) |
| Not equal | `≠` or `<>` |
| Floor | `⌊n/k⌋` |

Keep one equation per `equation` block. Do not put several `=` derivations in one `OMath` unless they are a single aligned object.

### Port and parameter tables

**Clocks and resets**

| Signal | Dir | Edge / type | Domain | Description |

**Ports**

| Port | Dir | Width | Domain | Handshake | Description |

**Parameters**

| Parameter | Default | Type | Description |

**Pipeline stages**

| Stage | Name | Cycles | Registers | Function |

**Timing**

| Metric | Value | Derivation |

## Update patches

When revising:

1. Bump `revision` and append `revisionHistory`.
2. Edit only JSON blocks under changed `id` headings and replace the matching PNGs.
3. Rebuild the same `OutPath`.
4. Keep figure captions stable unless a figure was added or removed; then renumber captions in JSON.

## Title-page and body style (applied by the builder)

- Paper: Letter, 1 in margins
- Body: Calibri 11 pt
- Headings: Calibri Light, dark steel blue
- Title: 28 pt; subtitle 16 pt
- Table header fill: RGB (31, 56, 100), white bold text
- Caption: 9 pt italic, steel blue
- Header: document title
- Footer: Altera copyright | revision | page number
- Title page: Author line only when an author was resolved; always Altera copyright line
