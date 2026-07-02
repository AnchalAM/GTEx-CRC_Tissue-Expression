import pandas as pd
import gzip


# File paths
GTEX_FILE = "GTEx_Analysis_2025-08-22_v11_RNASeQCv2.4.3_gene_median_tpm.gct.gz"
GENE_FILE = "gene_list.txt"
OUTPUT_FILE = "GTEx_TPM_tissue_expression.xlsx"

# Read gene list
with open(GENE_FILE, "r") as f:
    genes = [line.strip() for line in f]

print(f"Total genes: {len(genes)}")
# Read GTEx file
# -----------------------------
with gzip.open(GTEX_FILE, "rt") as f:

    # Skip first two lines, file is not a normal CSV/TSV file.the first two lines are metadata, not actual data.
    next(f)
    next(f)

    gtex = pd.read_csv(f, sep="\t")

print("GTEx file loaded successfully.")

# Rename first column
gtex.rename(columns={"Description": "Gene"}, inplace=True)

# Keep only required genes
result = gtex[gtex["Gene"].isin(genes)]

print(f"Genes found: {len(result)}")

# Save to Excel
result.to_excel(OUTPUT_FILE, index=False)

print("Done!")
print("Output saved as:", OUTPUT_FILE)