# %% [markdown]
# # Data Cleaning and Preprocessing in Pandas
# **Course:** Lab Space Discussion - Chapter 2.4
# **Dataset:** Spotify 2023 Song Metrics (works with iris.csv, movieprofit.csv and
# students.csv too - every step checks whether the column exists before touching it).
#
# Run this in Google Colab from top to bottom.

# %%
# ---------------------------------------------------------------
# STEP 0 - Load the dataset
# ---------------------------------------------------------------
import pandas as pd
import numpy as np

FILENAME = "spotify-2023.csv"   # <-- change this to your own file name

# Option A: upload the file from your computer (uncomment in Colab)
# from google.colab import files
# uploaded = files.upload()

# Option B: read it from Google Drive (uncomment in Colab)
# from google.colab import drive
# drive.mount('/content/drive')
# FILENAME = '/content/drive/MyDrive/your_folder/spotify-2023.csv'

# The Spotify file is NOT saved as UTF-8, so plain read_csv() throws
# UnicodeDecodeError. Trying encodings in order is what finally worked for me.
data = None
for enc in ["utf-8", "latin-1"]:
    try:
        data = pd.read_csv(FILENAME, encoding=enc)
        print(f"Loaded with encoding = {enc}")
        break
    except UnicodeDecodeError:
        print(f"encoding {enc} failed, trying the next one...")

print("Rows and columns:", data.shape)


# %%
# ---------------------------------------------------------------
# STEP 1 - First look at the raw data
# ---------------------------------------------------------------
print(data.head())
print()
data.info()


# %%
# ---------------------------------------------------------------
# STEP 2 - Diagnose the problems BEFORE fixing anything
# ---------------------------------------------------------------
print("--- Missing values per column (only columns that have some) ---")
missing = data.isna().sum()
print(missing[missing > 0] if (missing > 0).any() else "No missing values")

print("\n--- Exact duplicate rows ---")
print(data.duplicated().sum())

# A column that should be a number but pandas read as 'object' is a red flag:
# it means at least one value in there is text.
print("\n--- Columns stored as text instead of numbers ---")
print([c for c in data.columns if not pd.api.types.is_numeric_dtype(data[c])])

rows_before = len(data)


# %%
# ---------------------------------------------------------------
# STEP 3 - Rename columns so they are easier to type
# lowercase, no spaces, no '%' signs
# ---------------------------------------------------------------
data.columns = (data.columns
                .str.strip()
                .str.lower()
                .str.replace("%", "pct", regex=False)
                .str.replace(r"[^\w]+", "_", regex=True)
                .str.replace(r"_+", "_", regex=True)
                .str.strip("_"))
print(list(data.columns))


# %%
# ---------------------------------------------------------------
# STEP 4 - Turn text columns into real numbers
# Two problems in this dataset:
#   a) numbers written with a thousands separator -> "1,438"
#   b) one row where a whole block of text landed inside `streams`
# errors="coerce" turns anything that is not a number into NaN,
# so I can count exactly how many bad values there were.
# ---------------------------------------------------------------
should_be_numeric = [c for c in data.columns
                     if any(k in c for k in
                            ["streams", "playlists", "charts", "bpm",
                             "year", "month", "day", "count",
                             "revenue", "budget", "profit", "score",
                             "length", "width", "grade", "age"])]

for col in should_be_numeric:
    if not pd.api.types.is_numeric_dtype(data[col]):
        cleaned = (data[col].astype(str)
                   .str.replace(",", "", regex=False)     # 1,438 -> 1438
                   .str.replace("$", "", regex=False)     # $120M -> 120M
                   .str.strip())
        converted = pd.to_numeric(cleaned, errors="coerce")
        lost = converted.isna().sum() - data[col].isna().sum()
        print(f"{col:<25} -> numeric   ({lost} value(s) could not be converted)")
        data[col] = converted


# %%
# ---------------------------------------------------------------
# STEP 5 - Repair the mangled characters in the text columns
# Song titles came out as "Se├▒orita" because the file was encoded twice.
# Re-encoding latin-1 -> utf-8 puts the accents back.
# ---------------------------------------------------------------
def fix_text(value):
    if not isinstance(value, str):
        return value
    try:
        return value.encode("latin-1").decode("utf-8")
    except (UnicodeEncodeError, UnicodeDecodeError):
        return value          # already fine, leave it alone

text_columns = [c for c in data.columns
                if not pd.api.types.is_numeric_dtype(data[c])]

for col in text_columns:
    data[col] = data[col].str.strip().map(fix_text, na_action="ignore")
    # blank cells are missing values, not real text
    data[col] = data[col].replace(["", "nan", "None"], np.nan)


# %%
# ---------------------------------------------------------------
# STEP 6 - Duplicates and missing values
# ---------------------------------------------------------------
data = data.drop_duplicates()
print(f"Dropped {rows_before - len(data)} duplicate row(s)")

# The row whose `streams` cell held text is now NaN there. Inventing a median
# play count for it would be making data up, so that row goes.
if "streams" in data.columns:
    dropped = data["streams"].isna().sum()
    data = data.dropna(subset=["streams"])
    print(f"Dropped {dropped} row(s) with an unusable `streams` value")

# I did NOT use dropna() on the whole table, though: `key` alone is blank in
# many songs, and dropna() would have deleted those complete rows for the sake
# of one empty cell. So each remaining column gets its own treatment.
for col in data.columns:
    n_missing = data[col].isna().sum()
    if n_missing == 0:
        continue
    if pd.api.types.is_numeric_dtype(data[col]):
        fill = data[col].median()                  # median resists outliers
        data[col] = data[col].fillna(fill)
        print(f"{col:<25} {n_missing} missing -> filled with median {fill}")
    else:
        data[col] = data[col].fillna("Unknown")
        print(f"{col:<25} {n_missing} missing -> filled with 'Unknown'")

data_cleaned = data.reset_index(drop=True)


# %%
# ---------------------------------------------------------------
# STEP 7 - Build one real date column out of year / month / day
# ---------------------------------------------------------------
date_parts = ["released_year", "released_month", "released_day"]
if all(c in data_cleaned.columns for c in date_parts):
    data_cleaned["release_date"] = pd.to_datetime(
        data_cleaned[date_parts].rename(columns={
            "released_year": "year", "released_month": "month", "released_day": "day"}),
        errors="coerce")
    print(data_cleaned[["release_date"]].head())


# %%
# ---------------------------------------------------------------
# STEP 8 - Look for outliers with the IQR rule
# ---------------------------------------------------------------
for col in data_cleaned.select_dtypes(include="number").columns[:6]:
    q1, q3 = data_cleaned[col].quantile([0.25, 0.75])
    iqr = q3 - q1
    low, high = q1 - 1.5 * iqr, q3 + 1.5 * iqr
    n_out = ((data_cleaned[col] < low) | (data_cleaned[col] > high)).sum()
    print(f"{col:<25} {n_out} possible outlier(s)")
# Note: I am only *counting* them, not deleting them. A song with a huge number
# of streams is a real hit, not a typo.


# %%
# ---------------------------------------------------------------
# STEP 9 - Before / after and summary statistics
# ---------------------------------------------------------------
print(f"Rows before : {rows_before}")
print(f"Rows after  : {len(data_cleaned)}")
print(f"Missing left: {data_cleaned.isna().sum().sum()}")
print()
print(data_cleaned.describe())


# %%
# ---------------------------------------------------------------
# STEP 10 - Save the cleaned dataset
# ---------------------------------------------------------------
data_cleaned.to_csv("cleaned_dataset.csv", index=False, encoding="utf-8")
print("Saved as cleaned_dataset.csv")

# In Colab, download it with:
# from google.colab import files
# files.download('cleaned_dataset.csv')
