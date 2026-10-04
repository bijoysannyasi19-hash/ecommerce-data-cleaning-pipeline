import nbformat as nbf

nb = nbf.v4.new_notebook()

cells = []

# Section 1
cells.append(nbf.v4.new_markdown_cell("""# Week 1: Data Acquisition, Cleaning, and Preprocessing

## 1. Project Objective
The objective of this project is to acquire a publicly available dataset, explore its structure, identify data quality issues, and apply systematic data cleaning and preprocessing techniques. All decisions are documented, focusing on handling missing values, duplicates, outliers, and invalid data, while preparing a clean dataset for downstream analysis."""))

# Section 2
cells.append(nbf.v4.new_markdown_cell("""## 2. Dataset Selection and Source
**Dataset:** Online Retail
**Source:** UCI Machine Learning Repository
**Dataset Page:** https://archive.ics.uci.edu/dataset/352/online+retail

The dataset contains transactional records occurring between 01/12/2010 and 09/12/2011 for a UK-based and registered non-store online retail. It is an excellent dataset for learning data cleaning because it contains real-world data quality issues such as missing customer IDs, cancelled transactions, and extreme outlier values."""))

# Section 3
cells.append(nbf.v4.new_markdown_cell("""## 3. Data Acquisition
The dataset is downloaded directly from the UCI repository. We ensure that the original raw data is kept unmodified in the `data/raw/` directory."""))

cells.append(nbf.v4.new_code_cell("""import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as plt_sns
import os
import requests
import warnings
warnings.filterwarnings('ignore')

plt_sns.set_theme(style="whitegrid")

# Ensure directories exist
os.makedirs('../data/raw', exist_ok=True)
os.makedirs('../data/processed', exist_ok=True)
os.makedirs('../figures', exist_ok=True)

raw_path = '../data/raw/Online_Retail.xlsx'
url = "https://archive.ics.uci.edu/ml/machine-learning-databases/00352/Online%20Retail.xlsx"

if not os.path.exists(raw_path):
    print("Downloading dataset...")
    response = requests.get(url)
    with open(raw_path, 'wb') as f:
        f.write(response.content)
    print("Download complete.")
else:
    print("Dataset already exists locally.")

df_raw = pd.read_excel(raw_path)
df = df_raw.copy()
print(f"Dataset successfully loaded. Shape: {df.shape}")
print(f"File size on disk: {os.path.getsize(raw_path) / (1024 * 1024):.2f} MB")"""))

# Section 4
cells.append(nbf.v4.new_markdown_cell("""## 4. Dataset Overview and Data Dictionary
Let's look at the first few rows to understand the structure."""))

cells.append(nbf.v4.new_code_cell("""display(df.head())
display(df.info())"""))

cells.append(nbf.v4.new_markdown_cell("""### Data Dictionary
| Column Name | Meaning | Data Type | Expected Characteristics | Potential Data Quality Concern |
| :--- | :--- | :--- | :--- | :--- |
| InvoiceNo | Invoice number. A 6-digit integral number. If it starts with 'C', it indicates a cancellation. | Nominal/Object | Alphanumeric | May contain 'C' prefix for cancellations. |
| StockCode | Product code. A 5-digit integral number. | Nominal/Object | Alphanumeric | Non-product codes like 'POST', 'M' might exist. |
| Description | Product name. | Nominal/Object | Text | Inconsistent cases, missing descriptions. |
| Quantity | Quantities of each product per transaction. | Numeric | Integer > 0 | Negative values for returns/cancellations. |
| InvoiceDate | Invice Date and time. | Datetime | Date and Time | Correct parsing required. |
| UnitPrice | Product price per unit in sterling. | Numeric | Float > 0 | Zero or negative prices indicating bad data or adjustments. |
| CustomerID | Customer number. A 5-digit integral number. | Nominal/Float | Numeric | High rate of missing values for anonymous checkouts. |
| Country | Country name. | Nominal/Object | Text | Inconsistent spellings or 'Unspecified' country. |"""))

# Section 5
cells.append(nbf.v4.new_markdown_cell("""## 5. Initial Data Exploration
We will generate summary statistics and observe the distributions of numerical columns and categorical cardinality."""))

cells.append(nbf.v4.new_code_cell("""display(df.describe(include='all', datetime_is_numeric=True))

print("\\nNumber of unique values per column:")
display(df.nunique())"""))

cells.append(nbf.v4.new_markdown_cell("""Let's visualize the Top 10 Countries by number of transactions."""))

cells.append(nbf.v4.new_code_cell("""plt.figure(figsize=(10, 5))
df['Country'].value_counts().head(10).plot(kind='bar', color='skyblue')
plt.title('Top 10 Countries by Transaction Count')
plt.ylabel('Number of Transactions')
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig('../figures/top_countries.png')
plt.show()"""))

# Section 6
cells.append(nbf.v4.new_markdown_cell("""## 6. Missing Value Analysis
Now, let's identify missing values in the dataset and decide how to treat them."""))

cells.append(nbf.v4.new_code_cell("""missing_data = df.isnull().sum()
missing_percent = (missing_data / len(df)) * 100
missing_df = pd.DataFrame({'Missing Count': missing_data, 'Missing Percentage': missing_percent})
missing_df = missing_df[missing_df['Missing Count'] > 0].sort_values(by='Missing Percentage', ascending=False)
display(missing_df)"""))

cells.append(nbf.v4.new_markdown_cell("""**Observations:**
1. `CustomerID` is missing in a substantial portion of the records (approx 24.9%).
2. `Description` is missing in a small fraction of records (approx 0.26%).

**Decisions:**
- `CustomerID`: Since a customer ID uniquely identifies an entity making a purchase, imputing this with median/mode or a machine learning model would fabricate customer identities, completely invalidating customer-level analysis. I considered keeping these as a single 'Unknown' category, but for most cohort or CRM analyses, these rows add noise. However, because we might just want to analyze overall revenue or product popularity, completely dropping 24.9% of data is drastic. I will create a boolean flag `Is_Anonymous` and fill missing CustomerIDs with a placeholder like `-1` to retain the transaction data for revenue analysis without fabricating a real customer.
- `Description`: Since it's only 0.26%, and descriptions can often be inferred from `StockCode` or are genuinely missing for bad records, I will fill missing descriptions with the string "UNKNOWN" instead of dropping them, so we don't lose the associated numerical transaction data."""))

cells.append(nbf.v4.new_code_cell("""df['Is_Anonymous'] = df['CustomerID'].isnull()
df['CustomerID'].fillna(-1, inplace=True)
df['Description'].fillna("UNKNOWN", inplace=True)

print("Missing values after treatment:")
print(df.isnull().sum())"""))


# Section 7
cells.append(nbf.v4.new_markdown_cell("""## 7. Duplicate Analysis
Let's check for duplicate records. In retail data, a user might legitimately purchase the same item twice in separate carts, but an exact duplicate row at the exact same timestamp with the exact same invoice number is highly likely to be a system glitch."""))

cells.append(nbf.v4.new_code_cell("""duplicate_count = df.duplicated().sum()
print(f"Total duplicate rows: {duplicate_count}")
print(f"Duplicate percentage: {(duplicate_count / len(df)) * 100:.2f}%")"""))

cells.append(nbf.v4.new_markdown_cell("""**Decision:**
I will remove exact duplicate rows because there is no unique transaction line ID in this dataset. It's safer to assume exact matches (same Invoice, StockCode, Quantity, and exact timestamp) are data entry or system errors."""))

cells.append(nbf.v4.new_code_cell("""df.drop_duplicates(inplace=True)
print(f"Shape after removing duplicates: {df.shape}")"""))

# Section 8
cells.append(nbf.v4.new_markdown_cell("""## 8. Data Type Validation and Conversion
Checking if data types are appropriate."""))

cells.append(nbf.v4.new_code_cell("""print(df.dtypes)"""))

cells.append(nbf.v4.new_markdown_cell("""`CustomerID` is stored as float64 because it originally had missing values (NaN). Since we replaced NaN with -1, we can convert it to an integer. `InvoiceDate` is already datetime64. `InvoiceNo` is object, which is correct as it contains 'C' for cancellations. `StockCode` is also object."""))

cells.append(nbf.v4.new_code_cell("""df['CustomerID'] = df['CustomerID'].astype(int)
print(df.dtypes)"""))

nb.cells = cells
with open('c:/Users/sanny/Desktop/isro_project/project/src/generate_notebook_1.py', 'w') as f:
    pass # Wait, I shouldn't execute generation from strings in write_to_file directly if I can just write out the complete cells. Let's output it.
