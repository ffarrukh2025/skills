# Diagram style (Visio-like)
# Author: Fasih ud Din Farrukh

**Copyright:** Copyright (c) 2026 Altera Corporation. All rights reserved.

Polished, color-coded PNGs embedded in Word. Look like a Visio block diagram, not a slide sketch.

**Required renderer:** [scripts/VisioStyle.ps1](scripts/VisioStyle.ps1) (PowerShell + System.Drawing). Python is not required.

Dot-source it from the figure script:

```powershell
. "$env:USERPROFILE\.cursor\skills\hardware-design-document\scripts\VisioStyle.ps1"
New-VsCanvas 2200 1400
Add-VsTitle "Module name - view"
$a = Add-VsShape 80 160 420 280 "u_tx" "encode`nTable 68" "comb" "process"
$b = Add-VsShape 620 160 420 280 "PHY" "rxq 64 x 160" "mem" "store"
Add-VsConnector $a "right" $b "left" "tx_parallel_data[159:0]"
Add-VsLegend 80 500 @(@("comb","Datapath"), @("mem","FIFO"))
Save-VsCanvas ".\figures\fig_top.png"
```

## Visio look (do not skip)

- **Header bar** on process boxes: instance/role name in the dark strip, details in the body.
- **Orthogonal connectors** only (right-angle elbows). No diagonal arrows across the page except an FSM that cannot be routed otherwise.
- **Connection ports** at box mid-edges (small white dots).
- **Connector labels** in a white rounded tag sitting on the wire, never overlapping a box.
- **Drop shadow** on shapes (library default).
- **Containers** (dashed groups) for clock domain, testbench vs DUT, or a pipeline rank.
- **Shape kinds** by RTL role:

| Kind | Use |
|------|-----|
| `process` | Combinational or mixed module |
| `io` | Stimulus, AXIS, off-chip |
| `store` | FIFO, RAM, elastic queue (cylinder) |
| `register` | Flop stage, pipeline pill |
| `decision` | Compare / branch |
| `terminator` | Reset / start / stop |
| `note` | Checker, assumption, caption block |

- Reset/start FSM state: `doubleBorder = $true`.
- Back-pressure / reverse control: `Add-VsConnector ... -back` (dashed).
- Packet/Table bit maps: `Add-VsBitBar` (one bar, adjacent fields), not scattered independent boxes.
- Align on a 20 px grid. Keep ≥40 px between boxes so elbows and labels fit.
- One primary left-to-right data path per figure. Control loops go below, dashed.

## Canvas

- PNG, 200 DPI, ~2200 px wide (6.5 in on Letter)
- Background `#F7F9FC`
- Font Calibri
- Title uses the left accent bar from `Add-VsTitle`

## Color roles (same across a document)

| Role | Fill | Header / border |
|------|------|-----------------|
| `in` | `#D6EAF8` | `#2471A3` / `#1A5276` |
| `comb` | `#D5F5E3` | `#145A32` / `#196F3D` |
| `reg` | `#FCF3CF` | `#927709` / `#B7950B` |
| `mem` | `#FADBD8` | `#7B241C` / `#922B21` |
| `ctl` | `#E8DAEF` | `#5B2C6F` / `#6C3483` |
| `clk` | `#F5CBA7` | `#935116` / `#AF601A` |
| `err` | `#F5B7B1` | `#922B21` / `#C0392B` |
| `csr` | `#D4E6F1` | `#1F618D` / `#2874A6` |
| `ext` | `#F4F6F7` | `#34495E` / `#566573` |

## What not to do

- Do not use matplotlib / Graphviz defaults.
- Do not draw diagonal arrows when an orthogonal route exists.
- Do not place signal names on top of boxes.
- Do not put paragraphs inside shapes; one header + short body, rest in Word text.
- Do not skip the legend when ≥3 roles appear.

## Content checklist per figure

- [ ] Title states the module and view
- [ ] Every box maps to an RTL instance, process, or named stage
- [ ] Bus widths match parameters
- [ ] Connectors are orthogonal with labeled buses
- [ ] CDC / clock domain called out in a container if multiple clocks
- [ ] Legend present
