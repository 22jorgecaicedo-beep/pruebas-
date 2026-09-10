# Discussion post (paste this into D2L)

> Replace every `[ ]` with the real number your own run prints. The story below
> is the one this code actually produces; only the counts depend on your file.

---

**Dataset: Spotify 2023 Song Metrics**

I picked the Spotify file because I listen to a lot of this music and I wanted to
see whether the songs that end up on the most playlists are also the ones with
the highest number of streams. Before I could ask that, the file needed work.

**The first thing that happened was an error.** `pd.read_csv('spotify-2023.csv')`
threw `UnicodeDecodeError: 'utf-8' codec can't decode byte 0xf1`. The file is not
saved as UTF-8. Adding `encoding='latin-1'` loaded it, so I wrote a small loop
that tries UTF-8 first and falls back to latin-1:

```python
for enc in ["utf-8", "latin-1"]:
    try:
        data = pd.read_csv(FILENAME, encoding=enc)
        break
    except UnicodeDecodeError:
        continue
```

**Then `.info()` showed me the real problem.** `streams` was listed as `object`
instead of a number. A column of play counts should never be text, and that told
me at least one value in there was not a number. Same with
`in_deezer_playlists` and `in_shazam_charts`. Two different causes:

- Deezer and Shazam counts are written with a comma — `1,438` — so pandas read
  them as strings.
- In `streams` there is one row where an entire block of text
  (`BPM110KeyAModeMajor...`) got written into the cell. The row is shifted.

I stripped the commas and used `pd.to_numeric(..., errors="coerce")`, which
turns anything that is still not a number into `NaN` so I can count it:

```python
cleaned = data[col].astype(str).str.replace(",", "", regex=False).str.strip()
data[col] = pd.to_numeric(cleaned, errors="coerce")
```

That converted [X] values and left exactly [X] that could not be saved.

**The decision I spent the most time on was what to do about missing values.**
The tutorial shows `dropna()`, and my first version used it — I lost [X] rows.
That felt wrong, because most of those rows were only missing `key` (the musical
key of the song), and I was throwing away complete stream counts to fix one
blank cell. So I split it up:

- The row with unusable `streams` → **dropped.** Filling in a median play count
  would be inventing data.
- Missing numeric values → **median**, not mean, because a few massive hits pull
  the mean upward.
- Missing `key` → **"Unknown"**, because it is a category, not a quantity.

**I also left the outliers alone on purpose.** The IQR rule flagged [X] songs in
`streams`, but I checked and those are just genuinely huge hits. An outlier is
not automatically an error, and deleting the biggest songs would have destroyed
the exact thing I wanted to study.

Last steps: renamed the columns to lowercase with underscores (`danceability_%`
became `danceability_pct`, which is much easier to type), dropped [X] exact
duplicate rows, and combined `released_year`, `released_month` and
`released_day` into a single `release_date` column with `pd.to_datetime()`.

**Result:** [X] rows before → [X] rows after, and 0 missing values left.

**What I'm still unsure about:** filling `key` with "Unknown" adds a category
that does not really exist in music. Would it be better to leave those as `NaN`
so they are visibly missing? I couldn't decide.
