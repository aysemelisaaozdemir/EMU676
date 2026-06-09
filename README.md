
# ARP-SA: Assembly Routing Problem with Subcontractor Allocation

> **EMU676 — Optimization Models and Algorithms in Transportation and Distribution**  
> Hacettepe University · Industrial Engineering Department  
> **Author:** Ayşe Melisa Özdemir

---

## Problem Overview

The ARP-SA is a new combinatorial optimisation problem that integrates
three interleaved manufacturing decisions:

| Decision | Description |
|---|---|
| **Make-or-buy** | Produce each part in-house or outsource to a specialist firm |
| **Scheduling** | Sequence jobs at each production option with fixture-dependent changeovers |
| **Routing** | Route a vehicle fleet to collect finished parts from external firms |

Every order requires three part types — Mechanical (M), PCB/Electronic (E), and
Wiring harness (W) — all of which must reach the depot before assembly begins
(*synchronisation constraint*). The objective is **minimum total tardiness**.

The solution method is an **Adaptive Large Neighbourhood Search (ALNS)** with three
destroy operators (Random, Worst-Tardiness, Bottleneck) and two repair operators
(Greedy, Regret-2), embedded in a Simulated Annealing acceptance framework.

---

## Repository Structure

```
EMU676/
├── arp_sa_v2.py           # Core ALNS solver — run standalone for a quick demo
├── code1_instances.py     # Step 1 · Generates 22 test instances → Excel + JSON
├── code2_alns.py          # Step 2 · Runs ALNS on all instances  → Results Excel
├── code3_param_tuning.py  # Step 3 · Grid search parameter tuning → Tuning Excel
└── README.md
```

Auto-generated at runtime:
```
instances/                 # 22 JSON instance files (created by code1)
ARP_SA_Instances.xlsx      # Instance data for human inspection
ARP_SA_Results.xlsx        # Computational results
ARP_SA_ParamTuning.xlsx    # Parameter sensitivity analysis
```

---

## Requirements

- Python **3.9 or later**
- [`openpyxl`](https://pypi.org/project/openpyxl/) for Excel output

Install the only external dependency:

```bash
pip install openpyxl
```

No other third-party packages are needed.

---

## Quick Start

### Option A — Single demo run (30 seconds)

```bash
python arp_sa_v2.py
```

Runs ALNS on the hand-crafted **M01** reference instance (10 orders, 600 iterations)
and prints a formatted solution to the console.

Expected output:
```
============================================================
  ARP-SA v2 — ALNS Solver
============================================================
  Initial solution:  obj = 46.30 h
   Iter      Best      Curr       Temp
    100    0.0000    0.0000   0.28532
  ...
  En iyi: 0.0000  |  Başlangıç: 46.30  |  İyileştirme: 100.0%  |  CPU: ~2.0 s
```

### Option B — Full benchmark (15–30 minutes)

Run the three codes in order:

```bash
# Step 1: Generate 22 test instances
python code1_instances.py

# Step 2: Run ALNS on all instances (5 seeds each)
python code2_alns.py

# Step 3: Parameter tuning grid search  [~75 min]
python code3_param_tuning.py
```

---

## Running Each File

### `arp_sa_v2.py` — Core Solver

The self-contained ALNS solver. Contains:
- Instance data (M01 reference)
- Greedy initial solution constructor
- Three destroy operators + two repair operators
- ALNS main loop with SA acceptance and adaptive weights
- Output to `arp_sa_v2_results.json`

**Key function:**
```python
alns(n_iter=600, q=9, T0=6.0, alpha=0.97, rho=0.10, seed=42)
# Returns: best_assignment, best_sequences, best_routes, f_best, f_init, history, op_stats
```

---

### `code1_instances.py` — Instance Generator

Generates the 22-instance benchmark and saves results.

```bash
python code1_instances.py
```

**Outputs:**
- `instances/M01.json` … `instances/L06.json` (22 files)
- `ARP_SA_Instances.xlsx`

**Instance naming:**

| Group | Names | n (orders) | Difficulty range |
|---|---|---|---|
| Small | S01–S08 | 5, 7 | Easy → Hard |
| Medium | M01–M08 | 10, 12 | Easy → Hard |
| Large | L01–L06 | 15, 20 | Easy → Hard |

M01 is the hand-crafted reference instance; all others are generated with
tightness factor γ ∈ [0.78, 1.90] and complexity factor κ ∈ [0.80, 1.15].

---

### `code2_alns.py` — Batch ALNS Runner

Reads all instances from `instances/`, runs ALNS (5 seeds each), and writes
a detailed results workbook.

```bash
python code2_alns.py
```

**Prerequisites:** `instances/` folder must exist (run `code1_instances.py` first).

**Output — `ARP_SA_Results.xlsx` (6 sheets):**

| Sheet | Contents |
|---|---|
| `Ozet_Istatistikler` | f_init, f_best (mean/std/min/max), improvement %, CPU, service level |
| `Detayli_Calistirmalar` | Per-seed results with per-order tardiness |
| `En_Iyi_Atamalar` | Best solution: order ↔ option assignment + completion times |
| `Operator_Analizi` | Operator call counts, final weights, total improvement per operator |
| `q_Parametre_Analizi` | q sensitivity on M01 (q ∈ {2,3,4,5,6,9}) |
| `Yakinasama_Ozet` | Convergence data at sampled iteration points |

**Estimated runtime:**

| Instance group | Time per instance | Total |
|---|---|---|
| Small (n=5,7) | < 2 s | < 20 s |
| Medium (n=10,12) | ~7–9 s | ~75 s |
| Large (n=15,20) | ~15–55 s | ~250 s |

---

### `code3_param_tuning.py` — Parameter Tuning

Performs a full grid search over 180 parameter combinations on three
representative instances (M01, M03, L02).

```bash
python code3_param_tuning.py
```

**Grid:**

| Parameter | Values tested |
|---|---|
| q (remove count) | 2, 3, 4, 6, 9 |
| T₀ (initial SA temperature) | 2.0, 4.0, 6.0, 10.0 |
| α (cooling rate) | 0.95, 0.97, 0.99 |
| ρ (weight update step) | 0.10, 0.20, 0.30 |

Total: 180 combinations × 3 instances × 5 seeds = **2,700 ALNS runs**.

**Output — `ARP_SA_ParamTuning.xlsx` (8 sheets):**

| Sheet | Contents |
|---|---|
| `Tum_Kombinasyonlar` | All 180 combinations, colour-coded by performance |
| `q_Etkisi` | Effect of q (other parameters fixed) |
| `T0_Etkisi` | Effect of T₀ |
| `alpha_Etkisi` | Effect of α |
| `rho_Etkisi` | Effect of ρ |
| `En_Iyi_10_Kombinasyon` | Top 10 per instance with rankings |
| `q_alpha_Etkilesim_M01` | q × α interaction table on M01 |
| `Parametre_Tavsiyesi` | Final recommended parameter set |

**Estimated runtime:** ~60–90 minutes.

---

## Instance Data Format

Each `instances/XXX.json` file follows this schema:

```json
{
  "name":      "M01",
  "n_orders":  10,
  "n_vehicles": 2,
  "difficulty": "Kolay (Ref.)",
  "orders":    ["P01", "P02", "P03", "P04", "P05",
                "P06", "P07", "P08", "P09", "P10"],
  "due":       {"P01": 12, "P02": 24, "P03": 32, "P04": 8, "P05": 48,
                "P06": 25, "P07": 20, "P08": 28, "P09": 22, "P10": 34},
  "proc": {
    "P01": {
      "M": {"WM": 6.5, "sM1": 5.2, "sM2": 7.8},
      "E": {"WE": 3.8, "sE1": 3.1, "sE2": 5.4},
      "W": {"WW": 3.2, "sW1": 3.9, "sW2": 2.7}
    }
  },
  "fixtures":  {"P01": {"M": "F2", "E": "F1", "W": "F3"}, ...},
  "setup":     {"WM": 1.5, "sM1": 1.2, "sM2": 1.8,
                "WE": 0.8, "sE1": 0.6, "sE2": 1.0,
                "WW": 1.0, "sW1": 0.9, "sW2": 1.1},
  "tightness": "custom",
  "complexity": 1.0,
  "seed": 0
}
```

---

## Shared Infrastructure

All 22 instances use the same physical network.

**Production options per part type:**

| Part type | Internal workshop | External firms |
|---|---|---|
| Mechanical (M) | WM | sM1, sM2 |
| PCB/Electronic (E) | WE | sE1, sE2 |
| Wiring (W) | WW | sW1, sW2 |

**Travel times t_ij (hours, symmetric):**

|       | depot | sM1 | sM2 | sE1 | sE2 | sW1 | sW2 |
|-------|------:|----:|----:|----:|----:|----:|----:|
| depot |   0.0 | 1.3 | 2.1 | 0.9 | 1.7 | 1.5 | 2.4 |
| sM1   |   1.3 | 0.0 | 1.4 | 1.8 | 2.3 | 2.1 | 2.9 |
| sM2   |   2.1 | 1.4 | 0.0 | 2.5 | 1.9 | 2.7 | 1.6 |
| sE1   |   0.9 | 1.8 | 2.5 | 0.0 | 1.1 | 1.3 | 2.0 |
| sE2   |   1.7 | 2.3 | 1.9 | 1.1 | 0.0 | 1.8 | 1.4 |
| sW1   |   1.5 | 2.1 | 2.7 | 1.3 | 1.8 | 0.0 | 1.2 |
| sW2   |   2.4 | 2.9 | 1.6 | 2.0 | 1.4 | 1.2 | 0.0 |

**Setup times (hours) — triggered only on fixture change:**

| Option | s_us |   | Option | s_us |
|--------|-----:|---|--------|-----:|
| WM     | 1.5  |   | WE     | 0.8  |
| sM1    | 1.2  |   | sE1    | 0.6  |
| sM2    | 1.8  |   | sE2    | 1.0  |
| WW     | 1.0  |   | sW1    | 0.9  |
|        |      |   | sW2    | 1.1  |

---

## Recommended ALNS Parameters

From the grid search over 180 combinations (see `code3_param_tuning.py`):

| Parameter | Value | Description |
|---|---|---|
| **q** | **9** | Parts removed per destroy call |
| **T₀** | **6.0** | Initial SA temperature |
| **α** | **0.97** | Cooling rate |
| **ρ** | **0.10** | Operator weight update step |
| δ₁ / δ₂ / δ₃ | 3.0 / 2.0 / 1.0 | Reward scores (new best / improved / accepted) |
| n_iter | 300–700 | Iterations, scaled by problem size n |

---

## Key Results Summary

| Group | n | f_init range | f_best range | Improvement | CPU range |
|---|---|---|---|---|---|
| Small | 5–7 | 11–89 h | 0.6–52 h | 29–95% | 0.3–0.5 s |
| Medium | 10–12 | 46–252 h | 1–176 h | 21–98% | 1.0–1.8 s |
| Large | 15–20 | 161–769 h | 70–468 h | 20–56% | 2.5–10.3 s |

The ALNS achieves **19.9% to 97.7% improvement** over the initial greedy solution
across all 22 instances. CPU time scales near-linearly with problem size.

---

## Algorithm Summary

```
arp_sa_v2.py implements:

ALNS Loop (n_iter iterations):
  1. Select destroy operator d  (prob ∝ weight ω_d)
  2. Select repair operator r   (prob ∝ weight ω_r)
  3. σ' = Repair(r, Destroy(d, σ, q))
  4. Accept σ' with SA probability exp(−Δf / Θ)
  5. Update ω_d, ω_r based on outcome (δ1/δ2/δ3)
  6. Cool: Θ ← α·Θ

Destroy operators:
  - Random:          remove q random (p,τ) pairs
  - WorstTardiness:  remove q pairs with highest T_p score
  - Bottleneck:      remove q pairs with latest r_p^τ (assembly bottleneck)

Repair operators:
  - Greedy:   insert freed parts one-by-one, cheapest slot first (EDD order)
  - Regret-2: insert the part with largest (2nd_best − best) gap first
```

---

## File Dependencies

```
code1_instances.py   ──────────────────────────► instances/*.json
                                                        │
                                              ┌─────────┤
                                              ▼         ▼
code2_alns.py ──── imports arp_sa_v2.py   reads JSON   writes Excel
code3_param_tuning.py ── imports arp_sa_v2.py
```

`arp_sa_v2.py`, `code2_alns.py`, and `code3_param_tuning.py` must be in the
**same directory**. The `instances/` folder must be in the same directory as well.

---

## Citation

```bibtex
@techreport{ozdemir2026arpsa,
  author      = {Özdemir, Ayşe Melisa},
  title       = {Assembly Routing Problem with Subcontractor Allocation
                 in Manufacturing ({ARP-SA})},
  institution = {Hacettepe University, Department of Industrial Engineering},
  year        = {2026},
  note        = {EMU676 Final Report}
}
```

---

## License

Developed for academic purposes — EMU676, Hacettepe University, 2025–2026.
