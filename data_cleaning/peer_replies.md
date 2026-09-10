# Two peer replies

Adapt them — swap the bracketed parts for what the classmate actually posted.
A reply that names something specific from their code is what earns the points.

## Reply 1 — for someone who used `dropna()`

Hi [name], your notebook runs cleanly and the `.describe()` output at the end is
a nice way to close it. One thing I'd check: `dropna()` deletes an entire row if
*any* single column in it is blank. Try running this before and after your
cleaning cell:

```python
print(len(data), "->", len(data_cleaned))
print(data.isna().sum()[data.isna().sum() > 0])
```

When I did that on my dataset I found I was deleting whole rows of perfectly
good data because one non-essential column was empty. If that's happening to
you too, `data.dropna(subset=['the_column_that_matters'])` drops only the rows
that are actually unusable. Did you check how many rows you lost?

## Reply 2 — for someone who hit an error

Hi [name], thanks for posting the error instead of hiding it — I got a similar
one. [If it's `UnicodeDecodeError`:] that file isn't saved as UTF-8, so
`pd.read_csv(file, encoding='latin-1')` should get you past it. [If it's
`KeyError`:] that usually means the column name has a space or different
capitalization than what you typed — `print(list(data.columns))` shows the exact
names, and `data.columns = data.columns.str.strip().str.lower()` makes them
predictable from then on.

Also, a trick that saved me a lot of time: after loading, run `data.info()` and
look at the Dtype column. Anything that should be a number but says `object`
means there's text hiding somewhere in it. That's how I found a row in my file
where a whole block of text had landed in a numeric column.
