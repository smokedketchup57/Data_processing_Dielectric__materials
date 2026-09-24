# Dielectric Materials Data Processing

Python pipeline for batch impedance spectroscopy data processing, dielectric parameter extraction, and scientific plotting for biomass pellet samples (Miracle Fruit and Coffee Parchment).

---

## 📌 Overview

This repository provides an automated workflow to process multi-frequency capacitance (C_p) and conductance (G_p) experimental measurements obtained from impedance/dielectric analyzers. The pipeline:
- Filters high-frequency regimes (f = 8KHz) to mitigate low-frequency instrumentation noise.
- Averages multiple measurement sweeps across internal spreadsheet sheets (`Data` + replicates `Append1` to `Append9`).
- Normalizes raw electrical responses based on specimen geometry (circular pellet cross-section area A and thickness D).
- Extracts fundamental dielectric properties with error propagation (standard deviation $\sigma$ and standard error of the mean $\text{SEM}$):
  - Real Relative Permittivity (epsilon_r)
  - Imaginary Permittivity (epsilon_r)
  - Loss Tangent (tan\delta)
  - AC Conductivity (sigma_{ac})
- Exports consolidated multi-sample tables to Microsoft Excel (`.xlsx`).
- Generates automated diagnostic plots (C_p, G_p) and high-resolution publication-ready figures 

---

## 📂 Repository Structure

```text
.
├── Procesamiento_BienFM.py         # Main automated processing and plotting script
├── requirements.txt               # Required Python packages
├── README.md                      # Project documentation
├── Fruta_Milagrosa/               # Raw measurement files (.xls) for Miracle Fruit
│   ├── *.xls
│   ├── Diagnostico/               # Output: raw Cp/Gp diagnostic plots
│   └── Imagenes/                  # Output: scientific figures
└── Coffee_Parchment/              # Raw measurement files (.xls) for Coffee Parchment
    ├── *.xls
    ├── Diagnostico/               # Output: raw Cp/Gp diagnostic plots
    └── Imagenes_PC/               # Output: scientific figures
