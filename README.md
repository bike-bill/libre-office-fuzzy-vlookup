# libre-office-fuzzy-vlookup

A LibreOffice Basic macro that provides fuzzy (approximate) string matching, similar to Excel's `VLOOKUP` but tolerant of typos and minor variations.

Features:
- **Jaro-Winkler similarity** (default) — excellent for name matching and typo tolerance
- **Levenshtein distance** — simple edit-distance alternative
- **Automatic caching** — normalized strings and fuzzy scores cached for performance
- **Blocking optimization** — groups strings by prefix to reduce comparisons on large datasets
- **Tokenization** — handles "John Smith" vs "Smith, John" by sorting name parts

---

## Performance

The macro includes several optimizations for large datasets:

| Dataset Size | Comparisons | Uncached Time | Cached Time* |
|--------------|-------------|---------------|--------------|
| 1,000 × 10,000 | 10M | ~1.4 hrs | ~10 min |
| 3,600 × 10,000 | 36M | ~5 hrs | ~1.5 hrs |

*Assumes some repetition in lookup values (cache hits)

**Blocking optimization** reduces comparisons by grouping strings with matching prefixes (first 3 characters). For example, "William" only compares against table entries starting with "wil".

**Caches:**
- Normalized strings: up to 500,000 entries (~50 MB)
- Fuzzy scores: up to 10,000,000 entries (~2.5 GB)
- Caches persist until document closes
- **Caches start empty** and grow on-demand — a small spreadsheet with 100 names only uses ~100 cache entries, not the full limit
- **Score cache eviction**: when the score cache reaches 10M entries, it clears entirely and starts fresh (simple strategy suitable for most use cases)

---

## Installation

1. Open LibreOffice Calc.
2. Go to **Tools → Macros → Edit Macros…** (or **Tools → Basic IDE**).
3. In the Basic IDE, go to **File → Import Basic Source…** and select `FuzzyVlookup.bas`.  
   Alternatively, paste the contents of `FuzzyVlookup.bas` into any module.
4. The functions are now available as spreadsheet cell formulas.

---

## Functions

### `FuzzyVLookup`

```
=FuzzyVLookup(LookupValue, TableArray, IndexNum [, NFPercent] [, Rank] [, Algorithm])
```

Searches the first column of `TableArray` for the best fuzzy match to `LookupValue` and returns a value from that row.

| Parameter    | Type        | Default | Description |
|--------------|-------------|---------|-------------|
| `LookupValue`| String      | —       | The value to search for. |
| `TableArray` | Cell range  | —       | The range to search. The first column is searched; `IndexNum` selects which column to return. |
| `IndexNum`   | Integer     | —       | Column number to return (1 = first column of `TableArray`). Pass `0` to return the matched row's offset within the range instead. |
| `NFPercent`  | Single      | `0.05`  | Minimum match percentage (0–1). Matches below this threshold are ignored. |
| `Rank`       | Integer     | `1`     | Which match to return: `1` = best, `2` = second-best, etc. |
| `Algorithm`  | Integer     | `1`     | Matching algorithm: `1` = Jaro-Winkler (default, best for names/typos), `2` = Levenshtein distance. |

Returns the matched cell value, or a `#N/A` error if no match meets `NFPercent`.

**Examples**

```
=FuzzyVLookup("Willam", A2:C5, 2)              ' best match, Age column, default settings
=FuzzyVLookup("Willam", A2:C5, 2, 0.6)         ' require ≥ 60 % match
=FuzzyVLookup("Willam", A2:C5, 2, 0.5, 2)      ' return the 2nd-best match
=FuzzyVLookup("Willam", A2:C5, 0)              ' return matched row offset (1-based)
```

---

### `FuzzyPercent`

```
=FuzzyPercent(String1, String2 [, Algorithm] [, Normalised])
```

Returns a match score between 0 and 1 for two strings.

| Parameter    | Type    | Default | Description |
|--------------|---------|---------|-------------|
| `String1`    | String  | —       | First string. |
| `String2`    | String  | —       | Second string. |
| `Algorithm`  | Integer | `1`     | `1` = Jaro-Winkler (default), `2` = Levenshtein distance. |
| `Normalised` | Boolean | `False` | Pass `True` if strings are already lowercased/trimmed to skip normalization. |

**Examples**

```
=FuzzyPercent("William", "Willam")        ' → ~0.91
=FuzzyPercent("cat", "dog")              ' → low score
```

---

## Algorithms

| Value | Name | Description |
|-------|------|-------------|
| `1` | Character match | Scores how many individual characters from `String1` appear (in order) in `String2`. |
| `2` | Substring match | Scores how many substrings of increasing length (pairs, triplets, …) from `String1` appear in `String2`. |
| `3` | Combined | Runs both algorithms plus edit-distance similarity, then uses the strongest score. Recommended for general use and typo tolerance. |

---

## Testing

Run `TestFuzzyVLookup` from the Basic IDE (**Tools → Macros → Run Macro…**).  
The macro creates (or reuses) a sheet named **FuzzyVLookupTest**, populates it with sample data, runs a lookup, and displays the result in a message box.  Your other sheets are not affected.
