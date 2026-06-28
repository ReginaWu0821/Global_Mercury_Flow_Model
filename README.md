# Global Mercury Flow

This project runs the per-module scripts under `Code/` in workbook-number order,
reading the original Excel workbooks from `input/` and writing the computed
results to copies under `output/`. **`input/` is read-only at all times** —
every write goes to the copy under `output/`.

Simply enter data in the blue-shaded input fields and run the code 
to obtain sector-level emissions by country and global mercury flow by nation.

---

## 1. Directory layout

```
Global_Mercury_Flow/
├── Code/                       # All computation scripts (one module per workbook)
│   ├── _hg_expand.py           # Shared helper: expand template-row formulas by row offset
│   ├── _wb_resolver.py         # Shared helper: input→output path resolution and first-time copy
│   ├── _sheet_formula_sync.py  # Shared helper: keep formulas in sync with the input workbook
│   ├── _sheet_runner.py        # Shared entry point: run a single module sheet-by-sheet (also a CLI)
│   ├── 1.1 COMCoal_CW_CFPP/
│   │   └── *.py                # One script per sheet
│   ├── 1.2 COMCoal_CW_CFIB/
│   ├── ...
│   └── 6. Global Hg Flow/
├── input/                      # Original input workbooks (read-only)
│   ├── 1.1 COMCoal_CW_CFPP.xlsx
│   ├── ...
│   └── 6. Global Hg Flow.xlsx
├── output/                     # Run output (cleaned before each run by default)
├── run_all_workbooks.py        # Orchestrator: run all modules in workbook-number order
├── requirement.txt             # Python dependencies
└── README.md                   # This file
```


## 2. Environment setup

- Python ≥ 3.10 (3.10–3.13 recommended)
- Dependencies: only `openpyxl` (see `requirement.txt`)

Check that Python is available:

| OS      | Command               |
| ------- | --------------------- |
| Windows | `py --version`        |
| Linux   | `python3 --version`   |
| macOS   | `python3 --version`   |

Install dependencies (upgrading `pip` first is recommended):

| OS      | Command                                                                                  |
| ------- | ---------------------------------------------------------------------------------------- |
| Windows | `py -m pip install --upgrade pip` <br> `py -m pip install -r requirement.txt`            |
| Linux   | `python3 -m pip install --upgrade pip` <br> `python3 -m pip install -r requirement.txt`  |
| macOS   | `python3 -m pip install --upgrade pip` <br> `python3 -m pip install -r requirement.txt`  |



## 3. One-command run

From the repository root (the directory containing `run_all_workbooks.py`):

| OS      | Command                         |
| ------- | ------------------------------- |
| Windows | `py run_all_workbooks.py`       |
| Linux   | `python3 run_all_workbooks.py`  |
| macOS   | `python3 run_all_workbooks.py`  |

Default behavior:

- Reads inputs from `./input/`
- Writes outputs to `./output/` (**cleaned before the run**)
- Executes workbooks in order `1.1 → 1.2 → … → 6.`

Sample output:

```
[INPUT]  /path/to/Global_Mercury_Flow/input
[OUTPUT] /path/to/Global_Mercury_Flow/output  (cleaned)

=== [01/20] 1.1 COMCoal_CW_CFPP.xlsx ===
... per-sheet logs ...

=== [02/20] 1.2 COMCoal_CW_CFIB.xlsx ===
...

[ALL DONE] all workbook runners completed
[OUTPUT] results saved under: /path/to/Global_Mercury_Flow/output

[TIMING] run_all_workbooks.main elapsed: 00:01:23.456 (83.456s)
```

If any module returns a non-zero exit code, the orchestrator stops immediately
and propagates that exit code.



### Common options

`run_all_workbooks.py` accepts the following command-line options:

| Option                 | Default   | Description                                                                                            |
| ---------------------- | --------- | ------------------------------------------------------------------------------------------------------ |
| `-i`, `--input-dir`    | `input`   | Directory of input workbooks (relative paths are resolved against the repo root). **Read-only**.       |
| `-o`, `--output-dir`   | `output`  | Output directory. Inputs are first copied here, and every write goes to those copies.                  |
| `--keep-output`        | _off_     | Do not clean the output directory before the run (useful for incremental debugging — old files stay).  |

Examples:

```bash
# Custom input/output directories
python3 run_all_workbooks.py -i my_input -o my_output

# Keep previous outputs (do not clean)
python3 run_all_workbooks.py --keep-output
```



## 4. Conventions you must keep (important)

The project depends on exact file paths and names. Do not change them casually:

1. **Top-level directory names**: `Code/` and `input/` must be kept as-is.
2. **Module directory names**: subdirectories under `Code/` (including their
   numbers and spaces) must match the input workbook names one-to-one — do not
   rename them.
3. **Workbook names**: file names under `input/` must match the
   `工作簿: …/<file>.xlsx` reference written at the top of each script.
4. **Sheet names**: sheet names inside each `.xlsx` must remain unchanged.
5. **Shared entry points**: the four underscore-prefixed files
   `Code/_sheet_runner.py`, `Code/_wb_resolver.py`, `Code/_hg_expand.py`, and
   `Code/_sheet_formula_sync.py` must be kept.

> ⚠️ Do not "Save As" over a file under `input/` from inside Excel — Excel will
> rewrite the workbook and may alter internal structures we rely on. To inspect
> results, open the copy under `output/` instead.

