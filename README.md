# Week 1 Project: Data Acquisition, Cleaning, and Preprocessing

## 1. Project Objective
This project demonstrates the complete workflow of acquiring a raw dataset, assessing its quality, and systematically cleaning and preprocessing it for future analytical use. The objective is to identify and treat missing values, duplicates, and outliers in a real-world e-commerce dataset.

## 2. Dataset
**Name:** Online Retail  
**Source:** UCI Machine Learning Repository  
**URL:** [https://archive.ics.uci.edu/dataset/352/online+retail](https://archive.ics.uci.edu/dataset/352/online+retail)

This dataset contains transactional records for a UK-based non-store online retail company between 2010 and 2011.

## 3. Project Structure
```
project/
├── data/
│   ├── raw/                 # Original, unmodified dataset
│   └── processed/           # Final cleaned dataset
├── notebooks/
│   └── week1_data_cleaning.ipynb  # Executed Jupyter Notebook with full analysis
├── reports/
│   └── Week1_Data_Acquisition_Cleaning_Preprocessing.docx # Final Project Report
├── figures/                 # Visualizations generated during EDA
├── src/                     # Python scripts used for generation
└── README.md                # This file
```

## 4. Installation Requirements
To reproduce the environment, you will need Python 3.8+ and the following packages:
```bash
pip install pandas openpyxl matplotlib seaborn jupyter python-docx requests
```

## 5. How to Run
1. Open a terminal in the `project/` root directory.
2. Ensure you have the required libraries installed.
3. Run the Jupyter Notebook `notebooks/week1_data_cleaning.ipynb`. The notebook will automatically download the raw dataset into `data/raw/` if it does not exist, perform all cleaning steps, and save the final cleaned dataset to `data/processed/`.
4. (Optional) Re-run the report generation script to generate the DOCX report: `python src/build_report.py`

## 6. Preprocessing Workflow
The cleaning process involved:
- **Missing Value Handling**: Kept incomplete `CustomerID` rows to preserve transaction revenue, replacing them with a -1 flag for `Is_Anonymous`. Imputed missing `Description` values with "UNKNOWN".
- **Duplicate Removal**: Eliminated exact system duplicate records.
- **Data Type Corrections**: Ensured numeric columns were correctly formatted, including IDs.
- **Invalid Records**: Identified negative `Quantity` values as legitimate cancellations, and flagged them instead of dropping. Eliminated records with negative or zero `UnitPrice`.
- **Outlier Filtering**: Applied conservative percentile clipping (0.1% and 99.9%) to trim extremely large administrative and postage adjustments without blindly dropping genuine bulk orders.

## 7. Outputs
- `data/processed/online_retail_cleaned.csv`
- `notebooks/week1_data_cleaning.ipynb`
- `reports/Week1_Data_Acquisition_Cleaning_Preprocessing.docx`
- Generated PNG figures in `figures/`
