# Supporting Information: Kohn-Sham density encoding rescues coupled cluster theory for strongly correlated molecules 

[![DOI](https://zenodo.org/badge/DOI/10.5281/zenodo.17958091.svg)](https://doi.org/10.5281/zenodo.17958091)

Supporting code and analysis for the manuscript:

> **Kohn-Sham density encoding rescues coupled cluster theory for strongly correlated molecules**  
> Abdulrahman Y. Zamani, Barbaro Zulueta, Andrew M. Ricciuti, John A. Keith, and Kevin Carter-Fenk  
> (2025)

---

## Overview

This repository contains Jupyter notebooks and Python scripts for analyzing bond dissociation energies (BDEs) of first-row transition metal diatomics using CCSD(T) with various DFT reference orbitals (CCSD(T)@DFA) and standard KS-DFT methods.

### Key Features

- **CCSD(T)@DFA Analysis**: Benchmark of CCSD(T)/CBS with HF, SVWN5, PBE, PW91, R²SCAN, and PBE0 reference orbitals
- **KS-DFT Comparison**: Performance evaluation of 8 density functionals (SVWN5, PBE, PW91, R²SCAN, B3LYP, PBE0, ωB97X-V, ωB97M-V)
- **ph-AFQMC Reference**: Comparison with phaseless auxiliary-field quantum Monte Carlo data
- **Cr₂ PES Analysis**: Detailed potential energy surface study of the challenging Cr₂ dimer

---

## Repository Structure

```
.
├── README.md
├── calc_bde_tm_analysis.ipynb    # Main BDE analysis notebook (49 species)
├── cr2_pes_analysis.ipynb        # Cr₂ potential energy surface analysis
├── generate_tables.py            # Standalone Excel export script
│
├── ref_data/                     # Reference data (included)
│   ├── species_ref_data.csv      # Experimental BDEs, spin-orbit corrections
│   └── ph-afqmc_data.csv         # ph-AFQMC reference values
│
├── cr2_multi_diag/               # Output directory for Cr₂ CSV files
│
├── species/                      # ORCA output files (download from Zenodo)
│   ├── Sc-H/
│   ├── Sc-O/
│   ├── ...
│   └── Zn-Zn/
│
└── cr2_pes/                      # Cr₂ PES scan data (download from Zenodo)
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
openpyxl>=3.0
```

### Installation

```bash
# Clone repository
git clone https://github.com/BLZ11/cc-dft.git
cd cc-dft

# Install dependencies
pip install numpy pandas scipy matplotlib plotly openpyxl

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

```bash
jupyter notebook calc_bde_tm_analysis.ipynb
```

**Outputs:**
- `fig_benchmark_bar_mpl_with_qmc.pdf` — Bar chart for M–O, M–Cl, M–H (with ph-AFQMC)
- `fig_benchmark_bar_mpl_no_qmc.pdf` — Bar chart for M–H⁺, M–M
- `spark_consolidated.pdf` — 5×2 matrix of BDE curves

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
- `Cr-Cr_pes_main.pdf` — Main PES comparison figure
- `Cr-Cr_pes_SI.pdf` — SI PES comparison figure
- `cr2_multi_diag/Cr2_{method}.csv` — CSV files with PES data and T1 diagnostics for each method

### 3. Excel Export Script (`generate_tables.py`)

Standalone script to generate formatted Excel tables:

```bash
python generate_tables.py
```

**Output:** `bde_benchmark_results.xlsx` with sheets:
- **CC** — CCSD(T)@DFA results by bond type
- **KS-DFT** — DFT results by bond type
- **ph-AFQMC** — QMC reference data
- **Overall (M-H, M-O, M-Cl)** — Summary statistics with QMC
- **Overall (M-H+, M-M)** — Summary statistics without QMC

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
  title={Kohn-Sham density encoding rescues coupled cluster theory for strongly correlated molecules},
  author={Zamani, Abdulrahman Y. and Zulueta, Barbaro and Ricciuti, Andrew M. and Keith, John A. and Carter-Fenk, Kevin},
  year={2025}
}
```

And the data repository:

```bibtex
@dataset{Zamani2025data,
  title={Supporting Data: Kohn-Sham density encoding rescues coupled cluster theory for strongly correlated molecules},
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
