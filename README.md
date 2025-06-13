# Lens Grid Generator

A Python GUI tool for generating lens power grids and exporting them to Excel. Supports drag-and-drop or file selection for loading input parameters from JSON or Excel.

## Features

- Easy input for SPH, CYL, and Axis grid parameters
- Load input values from `.json` or `.xlsx` via drag-and-drop or file dialog
- Generates all combinations in grid form
- Exports results to Excel
- Auto-opens the generated file

## Requirements

- Python 3.x
- `pandas`
- `openpyxl`
- `tkinterdnd2`

Install dependencies:

```bash
pip install -r requirements.txt
```

## Usage

```bash
python lens_grid_generator.py
```

## Input File Format

- **JSON**: Should be a dict with keys matching the input fields (`sph_start`, `sph_step`, etc.)
- **Excel**: Should have columns `Field` and `Value` (see example below)

| Field      | Value |
|------------|-------|
| sph_start  | -6    |
| sph_step   | 0.25  |
| sph_max    | 6     |
| cyl_start  | -2    |
| cyl_step   | 0.25  |
| cyl_max    | 2     |
| axis_start | 0     |
| axis_step  | 10    |
| axis_max   | 180   |

---

**Generate complete lens grids easily!**
