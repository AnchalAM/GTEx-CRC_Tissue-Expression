# GTEx CRC Tissue Expression Analysis

## About the Project

This project was developed as part of my M.Sc. Big Data Biology coursework to study tissue-specific gene expression of colorectal cancer (CRC) biomarker genes using the GTEx database.

The program extracts TPM (Transcripts Per Million) values for selected genes and saves the results in Excel format. These results can be used for further analysis of gene expression across different human tissues.

## Objectives

- Extract TPM values from the GTEx dataset.
- Study the expression of important colorectal cancer biomarker genes.
- Save the processed data in Excel files for further analysis.

## Biomarker Genes

- TP53
- APC
- KRAS
- BRAF
- PIK3CA
- PTEN
- NRAS
- FBXW7
- SMAD4
- MLH1
- MSH2
- MSH6
- PMS2

## Files Included

main.py: Python script used for the analysis 
gene_list.txt: List of biomarker genes 
GTEx_TPM_tissue_expression.xlsx: Tissue-wise TPM values 
CRC_NAC_GTEx_Biomarkers.xlsx: Final biomarker results 

## Requirements

Python 3.10 or above
pandas
openpyxl
Install the required packages using:```pip install pandas openpyxl```

## Data Source

The gene expression data used in this project is obtained from the GTEx (Genotype-Tissue Expression) Portal.
Website: https://gtexportal.org

**Note:** The original GTEx dataset (`.gct.gz`) is not included in this repository because it is too large to upload on GitHub. It can be downloaded from the GTEx Portal.

