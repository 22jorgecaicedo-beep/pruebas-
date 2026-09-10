# Discussion post — listo para pegar en D2L

**Dataset: Spotify 2023 Song Metrics (953 songs, 24 columns)**

I picked this one because I wanted to check whether the songs that appear on the
most playlists are also the ones with the most streams. Before I could ask that,
the file needed work.

`.info()` is where I found the real problem. `streams` came back as `object`
instead of a number. A column of play counts should never be text, and that
told me something non-numeric was hiding in there. Same with
`in_deezer_playlists` and `in_shazam_charts`. Two different causes:

- The Deezer and Shazam numbers are written with a thousands separator
  (`1,438`), so pandas read the whole column as strings.
- In `streams` there is one row where a block of text
  (`BPM110KeyAModeMajor...`) ended up inside the cell. That row is shifted.

I stripped the commas and used `pd.to_numeric(..., errors="coerce")`, which
turns whatever is still not a number into `NaN` so I can count exactly how
many bad values there were:

```python
cleaned = data[col].astype(str).str.replace(",", "", regex=False).str.strip()
data[col] = pd.to_numeric(cleaned, errors="coerce")
```

`in_deezer_playlists` and `in_shazam_charts` converted with **0** values lost —
it really was just the commas. `streams` lost exactly **1**: the broken row.

Missing values: **95** songs with no `key` and **50** with no
`in_shazam_charts`. **0** exact duplicate rows, which honestly surprised me.

**The decision I spent the most time on was `dropna()`.** The tutorial uses it
and my first version did too. But `dropna()` deletes an entire row if *any*
column in it is blank, so it was throwing out the 95 rows missing `key` — and
those rows had perfectly good stream counts. I was destroying real data to fix
one empty cell. So I handled each column separately instead:

- The row with the unusable `streams` value → **dropped**. Filling in a median
  play count there would be inventing data.
- Missing numbers → **median**, not mean, because a handful of enormous hits
  drag the mean upward.
- Missing `key` → **"Unknown"**, because it is a category, not a quantity.

**953 rows → 952 rows, and 0 missing values left.**

I also left the outliers alone on purpose. The IQR rule flagged 109 songs in
`in_spotify_playlists` and 150 in `released_year`, but those are not errors.
`describe()` showed me the oldest song in the file is from **1930**, and streams
run from **2,762 up to 3.7 billion**. Those are old songs that came back into
the charts and genuine mega-hits. Deleting them would have destroyed the exact
thing I wanted to study.

(One thing that did *not* happen: I had written a fallback in case the file was
not saved as UTF-8, since encoding errors are common with this dataset. Mine
loaded on the first try, so the fallback never ran. I left it in anyway.)

**What I'm still unsure about:** filling the 50 missing `in_shazam_charts` with
the median gave me **2.5**. That is the correct median, but "appeared on 2.5
charts" is not a real thing — it's a count, so it should be a whole number. I
don't know whether the right move is to round it, use 0, or just leave those as
`NaN` and be upfront that they are unknown. Has anyone run into medians on a
count column?
