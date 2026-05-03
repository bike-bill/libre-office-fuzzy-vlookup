# libre-office-fuzzy-vlookup

A LibreOffice Basic macro that provides fuzzy (approximate) string matching, similar to Excel's `VLOOKUP` but tolerant of typos and minor variations.

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
| `Algorithm`  | Integer     | `3`     | Matching algorithm: `1` = character matching only, `2` = substring matching only, `3` = both combined. |

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
| `Algorithm`  | Integer | `3`     | `1` = characters, `2` = substrings, `3` = both. |
| `Normalised` | Boolean | `False` | Pass `True` if strings are already lowercased/trimmed to skip normalisation. |

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
