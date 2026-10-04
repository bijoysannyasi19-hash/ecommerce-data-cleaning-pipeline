import os
import requests
import pandas as pd
import warnings

warnings.filterwarnings('ignore')

def download_and_save():
    # URL for Online Retail dataset from UCI
    url = "https://archive.ics.uci.edu/ml/machine-learning-databases/00352/Online%20Retail.xlsx"
    raw_path = "data/raw/Online_Retail.xlsx"
    
    if not os.path.exists(raw_path):
        print("Downloading dataset...")
        response = requests.get(url, stream=True, verify=False)
        response.raise_for_status()
        with open(raw_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)
        print("Download complete.")
    else:
        print("Dataset already downloaded.")

    print("Loading data to verify...")
    df = pd.read_excel(raw_path)
    print(f"Shape: {df.shape}")
    print("Columns:", df.columns.tolist())
    print(df.head())

if __name__ == "__main__":
    download_and_save()
