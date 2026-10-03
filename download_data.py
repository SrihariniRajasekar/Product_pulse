"""Download the Olist dataset from Kaggle and copy the CSVs into data/raw."""
import shutil
from pathlib import Path

import kagglehub

dest = Path(__file__).resolve().parent / "data" / "raw"
dest.mkdir(parents=True, exist_ok=True)

src = kagglehub.dataset_download("olistbr/brazilian-ecommerce")
for f in Path(src).glob("*.csv"):
    shutil.copy(f, dest / f.name)

print("Copied:", sorted(p.name for p in dest.glob("*.csv")))
