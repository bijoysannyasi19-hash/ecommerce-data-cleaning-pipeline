import ssl
import urllib.request

url = "https://archive.ics.uci.edu/ml/machine-learning-databases/00352/Online%20Retail.xlsx"
path = "data/raw/Online_Retail.xlsx"

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

print("Downloading via urllib...")
with urllib.request.urlopen(url, context=ctx) as response, open(path, 'wb') as out_file:
    data = response.read()
    out_file.write(data)
print("Done.")
