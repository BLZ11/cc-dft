# Supporting Information: Improving coupled cluster theory for strongly correlated molecules with Kohn-Sham density encoding

[![DOI](https://zenodo.org/badge/DOI/10.5281/zenodo.17958091.svg)](https://doi.org/10.5281/zenodo.17958091)

Supporting code and analysis for the manuscript:

> **Improving coupled cluster theory for strongly correlated molecules with Kohn-Sham density encoding**  
> Abdulrahman Y. Zamani, Barbaro Zulueta, Andrew M. Ricciuti, John A. Keith, and Kevin Carter-Fenk  
> (2025)

---

## Overview

This repository contains Jupyter notebooks and Python scripts for analyzing bond dissociation energies (BDEs) of first-row transition metal diatomics using CCSD(T) with various DFT reference orbitals (CCSD(T)@DFA) and standard KS-DFT methods, plus the code that reproduces the supplementary figures of the manuscript.

### Key Features

- **CCSD(T)@DFA Analysis**: Benchmark of CCSD(T)/CBS with HF, SVWN5, PBE, PW91, R²SCAN, and PBE0 reference orbitals
- **KS-DFT Comparison**: Performance evaluation of 8 density functionals (SVWN5, PBE, PW91, R²SCAN, B3LYP, PBE0, ωB97X-V, ωB97M-V)
- **ph-AFQMC Reference**: Comparison with phaseless auxiliary-field quantum Monte Carlo data
- **Cr₂ PES Analysis**: Detailed potential energy surface study of the challenging Cr₂ dimer
- **Supplementary Figures**: Single self-contained notebook reproducing Supplementary Figures S3–S7, S10–S13, and S15–S53

---

## Repository Structure

```
.
├── README.md
├── LICENSE                              # MIT License
├── calc_bde_tm_analysis.ipynb           # Main BDE analysis notebook (49 species)
├── cr2_pes_analysis.ipynb               # Cr₂ potential energy surface analysis
├── si_figures_all.ipynb                 # Supplementary Figures S3–S53 (self-contained)
├── generate_tables.py                   # Standalone Excel export script
├── spin_state_energetics_nned_data.xlsx # Source data workbook (figures and tables)
│
├── ref_data/                            # Reference data (included)
│   ├── species_ref_data.csv             # Experimental BDEs, spin-orbit corrections
│   └── ph-afqmc_data.csv                # ph-AFQMC reference values
│
├── cr2_multi_diag/                      # Output directory for Cr₂ CSV files
│
├── species/                             # ORCA output files (download from Zenodo)
│   ├── Sc-H/
│   ├── Sc-O/
│   ├── ...
│   └── Zn-Zn/
│
└── cr2_pes/                             # Cr₂ PES scan data (download from Zenodo)
    ├── ccsdt_hf/
    ├── ccsdt_pbe/
    ├── ...
    └── exp.txt
```

---

## Data Availability

The ORCA output files (`species/` and `cr2_pes/` directories) are available on Zenodo:

**https://zenodo.org/records/17958091**

Download and extract the archive:

```bash
# Download from Zenodo
wget https://zenodo.org/records/17958091/files/data.tar.gz

# Extract
tar -xzf data.tar.gz
```

The `si_figures_all.ipynb` notebook and the `spin_state_energetics_nned_data.xlsx` workbook require no external downloads.

---

## Requirements

### Dependencies

```
python>=3.9
numpy>=1.20
pandas>=1.3
scipy>=1.7
matplotlib>=3.5
plotly>=5.0
seaborn>=0.12
openpyxl>=3.0
```

`seaborn` is used only by `si_figures_all.ipynb` (Figures S50 and S52).

### Installation

```bash
# Clone repository
git clone https://github.com/BLZ11/cc-dft.git
cd cc-dft

# Install dependencies
pip install numpy pandas scipy matplotlib plotly seaborn openpyxl

# Download and extract data from Zenodo
wget https://zenodo.org/records/17958091/files/data.tar.gz
tar -xzf data.tar.gz
```

---

## Usage

### 1. BDE Analysis Notebook (`calc_bde_tm_analysis.ipynb`)

Main analysis notebook for 49 first-row transition metal diatomics:

- **M–H**: ScH, TiH, VH, CrH, MnH, FeH, CoH, NiH, CuH, ZnH
- **M–O**: ScO, TiO, VO, CrO, MnO, FeO, CoO, NiO, CuO, ZnO
- **M–Cl**: TiCl, VCl, CrCl, MnCl, FeCl, CoCl, NiCl, CuCl, ZnCl
- **M–H⁺**: ScH⁺, TiH⁺, VH⁺, CrH⁺, MnH⁺, FeH⁺, CoH⁺, NiH⁺, CuH⁺, ZnH⁺
- **M–M**: Sc₂, Ti₂, V₂, Cr₂, Mn₂, Fe₂, Co₂, Ni₂, Cu₂, Zn₂

**Features:**
- Interactive Plotly figures for data exploration
- Publication-quality matplotlib figures (PDF/PNG export)
- Error statistics (RMSE, MAE, MAX) by bond type
- Comparison with ph-AFQMC reference data

The publication figures are signed-error box-and-whisker plots with every individual species overlaid, per Nature Communications editorial policy on bar graphs. Boxes span the interquartile range with the median marked, whiskers extend to 1.5 times the interquartile range, white diamonds mark the MAE, and gray shading marks the chemical accuracy range (±3 kcal/mol ≈ ±0.13 eV). The M–M and Overall panels of the second figure use a broken y-axis to isolate the −10.7 eV Mn₂ CCSD(T)@HF error.

```bash
jupyter notebook calc_bde_tm_analysis.ipynb
```

**Outputs:**
- `fig_benchmark_dist_with_qmc_signed_grid.pdf/.png` — Distribution plots for M–O, M–Cl, M–H, and Overall (with ph-AFQMC)
- `fig_benchmark_dist_no_qmc_signed_grid.pdf/.png` — Distribution plots for M–H⁺, M–M, and Overall
- `source_data_benchmark_distributions.csv` — Source data: one row per plotted point (figure, panel, method, bond type, species, signed error in eV)
- `spark_consolidated.pdf/.png` — 5×2 matrix of BDE curves

### 2. Cr₂ PES Analysis Notebook (`cr2_pes_analysis.ipynb`)

Detailed analysis of the Cr₂ potential energy surface:

- PES curves for all CCSD(T)@DFA methods
- Comparison with experimental PES (Larsson et al., 2022)
- Comparison with best theoretical estimate (BTE)
- T1 diagnostic analysis

```bash
jupyter notebook cr2_pes_analysis.ipynb
```

**Outputs:**
- `Cr-Cr_pes_main.pdf/.png` — Main PES comparison figure
- `Cr-Cr_pes_SI.pdf/.png` — SI PES comparison figure
- `cr2_multi_diag/Cr2_{method}.csv` — CSV files with PES data and T1 diagnostics for each method

### 3. Supplementary Figures Notebook (`si_figures_all.ipynb`)

Single notebook reproducing Supplementary Figures S3–S7, S10–S13, and S15–S53 (48 figures from 77 originally standalone notebooks). All data are embedded in the code cells, so the notebook runs top to bottom with no external files and completes in a few minutes. Each figure has its own markdown section header; multi-panel figures (S3–S5 orbital plots, S25–S26 per-molecule plots) carry sub-headers.

```bash
jupyter notebook si_figures_all.ipynb
```

**Outputs:** every figure displays inline and is also saved as SVG and PNG under its original filename. Two minimal adjustments were made relative to the standalone notebooks: saved-file names in the Figure S3–S5 sections carry a `figure_sN_` prefix so the three sections do not overwrite each other's files, and `.astype(int)` was appended to the `df.replace` calls in Figures S50 and S52 for pandas ≥ 3 compatibility. Plot titles are unchanged.

### 4. Excel Export Script (`generate_tables.py`)

Standalone script to generate formatted Excel tables:

```bash
python generate_tables.py
# Select unit: 1 (eV), 2 (kcal/mol), 3 (kJ/mol)
```

**Output:** `bde_results_<unit>.xlsx` with sheets:
- **CC** — CCSD(T)@DFA results by bond type
- **KS-DFT** — DFT results by bond type
- **ph-AFQMC** — QMC reference data
- **Overall (M-H, M-O, M-Cl)** — Summary statistics with QMC
- **Overall (M-H+, M-M)** — Summary statistics without QMC

### 5. Source Data Workbook (`spin_state_energetics_nned_data.xlsx`)

Excel workbook with the numerical data behind the manuscript figures and tables, organized as one sheet per item (for example `Figure1-Data-1`, `Table1-main`, `SFigure7`, `SFigure28-32`, `STable8`). Contents include singlet-triplet gap energetics at CCSD(T)/cc-pVTZ with restricted, unrestricted, and broken-symmetry references, NNED metrics, and T1 diagnostics for the main group and metal datasets.

---

## Computational Details

- **Basis Set (Geometry Optimization)**: def2-TZVP
- **Basis Set (CCSD(T) CBS Extrapolation)**: def2-nZVPP (n = T, Q)
- **Relativistic Treatment**: X2C Hamiltonian; ZORA was used for the BDE calculation of Mn₂
- **Spin-Orbit Corrections**: Applied from experimental/theoretical references
- **Software**: ORCA 6.0

---

## Citation

If you use this code or data, please cite:

```bibtex
@article{Zamani2025ms,
  title={Improving coupled cluster theory for strongly correlated molecules with Kohn-Sham density encoding},
  author={Zamani, Abdulrahman Y. and Zulueta, Barbaro and Ricciuti, Andrew M. and Keith, John A. and Carter-Fenk, Kevin},
  year={2025}
}
```

And the data repository:

```bibtex
@dataset{Zamani2025data,
  title={Supporting Data: Improving coupled cluster theory for strongly correlated molecules with Kohn-Sham density encoding},
  author={Zamani, Abdulrahman Y. and Zulueta, Barbaro and Ricciuti, Andrew M. and Keith, John A. and Carter-Fenk, Kevin},
  year={2025},
  publisher={Zenodo},
  doi={10.5281/zenodo.17958091}
}
```

---

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

## Contact

For questions or issues, please open a GitHub issue or submit a pull request.
