# Methodology: odour distance vs. environmental category

## Question

Is there an association between the environmental category (*milieucategorie*) of an activity and its guidance distance for odour (*geurafstand*) in the VNG 2009 *Staat van bedrijfsactiviteiten*?

- Odour distance: column E (`GEUR`) in both source files.
- Category: column R (`CATEGORIE`) in both source files.

**Column mapping.** You named column N for bijlage 1. In that file column N is empty, and columns L, M, N and P are hidden. Counted over visible columns only, `CATEGORIE` (R) is the 14th column, i.e. the "N" you see on screen. Column R is therefore used for both files.

## Sources (read-only)

| Code | File | Sheet used |
|---|---|---|
| B1 | `doc/milieuzonering_bijlage_1_deel_1_en_2.xlsx` | `EINDTABEL 1 met formule` |
| B2 | `doc/bijlage_2_staat_bedrijfsactiviteiten_bedrijventerreinen_versie_maart_2009_def.xlsx` | `EINDTABEL 1 met formule` |

The files in `doc/` are never written. `build_analysis.py` checks their SHA-256 hashes before and after each run, and records them on the `Statistics` sheet.

## How the category is computed in the sources

In both files:

1. `O` (`GROOTSTE AFSTAND`) = `MAX(J, G, F, E)`, the largest of the distances for gevaar, geluid, stof and geur.
2. `P` (`CAT`) = `VLOOKUP(O, codetabel!A1:B12, 2, FALSE)`, which maps that distance to a category.
3. `R` (`CATEGORIE`) = `IF(ISNA(P), "", P)`, which shows nothing for header rows.

The code table: 10 m → 1, 30 → 2, 50 → 3.1, 100 → 3.2, 200 → 4.1, 300 → 4.2, 500 → 5.1, 700 → 5.2, 1000 → 5.3, 1500 → 6. This matches the VNG rule that the largest distance determines the category.

The category is therefore not independent of the odour distance, because odour is one of the four inputs to the maximum. This structural dependency is central to how the correlation should be read (see the analyses in step 5).

Deviations from the pattern, all preserved and checked:

- B1 has 15 included rows, and B2 has 7, where R is a hardcoded value rather than a formula. Examples are B1 rows 430, 454, 465 and 472–474 (wind turbines) and 618 (helipads).
- Some rows have O hardcoded or written as `MAX(E:J)`, or have distances stored as text (`'0'`, `'10'`).
- B2 rows 423–424 look up the category through an external link (`[1]codetabel`). The cached values are used.
- B1 rows 763–837 have no SBI code, description, or other distances. They contain only a geur value (sometimes `-`) and a hardcoded category. According to the owner of the file, they are a targeted cut-and-paste of columns E and R from an updated version of bijlage 1. That version was published on the VNG website for the former guideline and can no longer be found. See step 3.

## Processing steps

All steps are in `analysis/build_analysis.py`. You can reproduce them with:

```
python3 analysis/build_analysis.py
python3 <xlsx-skill>/scripts/recalc.py analysis/geur_categorie_analyse.xlsx 120
```

The second command recalculates all formulas with LibreOffice so that cached values exist. It reported 29,219 formulas and 0 errors.

### 1. Extraction (intermediate sheets `B1_extract`, `B2_extract`)

The script takes every non-empty row from row 3 down. For each row it copies:

- the identifiers (A–D);
- the raw values of E, F, G and J, exactly as stored (number or text);
- the source content of O and R (the formula text or the hardcoded value);
- the cached value shown in Excel, which is the value used in the analysis. The category is normalised to a text label, e.g. integer `2` → `"2"`.

### 2. Verification (yellow formula columns on the extract sheets)

Live formulas recompute the source logic independently:

- `Geur/Stof/Geluid/Gevaar num`: text to number. `-` and empty cells become empty.
- `Max distance recalc` = `MAX` of the four numbers.
- `Category recalc` = `INDEX/MATCH` on the `Codetabel` sheet.
- `O check` and `R check` compare the cached source values with the recalculated ones. They show `n/a (excluded)` for rows outside the dataset.

Result: for all 559 included B1 rows and all 315 included B2 rows, both checks read `OK`. This includes the hardcoded R values. A summary is on `Statistics`, under "Verification of source values".

### 3. Inclusion rule

A row is included if it has (a) an activity description, (b) a numeric odour distance, where `0` counts as a value and `-` and empty do not, and (c) a category. The column `Exclusion reason` states why any other row is excluded. Excluded rows are mostly section headers, for example "Slachterijen en overige vleesverwerking:".

| | B1 | B2 |
|---|---|---|
| rows extracted | 824 | 467 |
| no numeric odour distance (headers) | 149 | 117 |
| no activity description | 116 | 35 |
| **included** | **559** | **315** |

B1 rows 763–837 come from the updated version of bijlage 1 (see above). They are excluded because they cannot be linked to an activity:

- They have no SBI code, description, or stof/geluid/gevaar distances, so their category cannot be verified. In several rows the category cannot follow from the geur value alone, e.g. geur 0 m with category 6, which is consistent with the category being set by another factor in the updated list.
- The block of 75 rows contains 11 empty rows, which are probably section headers in the updated list. I compared the pattern of empty rows and of the E and R values against every 75-row window of the main table. No window matches: the best window agrees on 60 of 75 row types and on only 12 categories. So the block is not an updated copy of a contiguous part of the 2009 table, and the rows cannot be assigned to activities from that table.

They are reported separately as a sensitivity analysis (README sheet and `FINDINGS.md`). If the updated version is found, they can be added as a separate dataset.

### 4. Final datasets (values only)

`Data_B1`, `Data_B2` and `Data_B1_union_B2` contain values only: one row per activity, with odour distance (m), category label, ordinal rank (1–10), and the derived variables:

- `Geur decisive`: whether the geur distance equals the maximum distance, i.e. whether odour determines the category.
- `Max other distances` and `Category rank excl. geur`: the category the activity would have on stof, geluid and gevaar alone (rank 0 if all three are 0 m).
- `In both B1 and B2`: matched on old SBI code, sub-number and description. 312 of the 315 B2 activities also appear in B1, with identical geur and category values, so B2 is almost a subset of B1. The union adds only 3 activities (n = 562). The two lists are therefore analysed separately and should not be pooled as independent samples.

Columns N–R on the dataset sheets are formulas that compute mid-ranks (average rank for ties): `COUNTIF(range,"<"&x) + (COUNTIF(range,x)+1)/2`.

### 5. Statistics (`Statistics` sheet)

Both variables are ordinal with many ties. The odour distance is measured on the same stepped, roughly logarithmic scale as the category. The analysis therefore uses rank correlations:

- **Spearman's ρ**: computed as `CORREL` of the mid-rank columns, as a live formula, and cross-checked against `scipy.stats.spearmanr`. All seven analyses match to below 10⁻⁴.
- **Kendall's τ-b**: `scipy.stats.kendalltau` with `variant="b"`, which corrects for ties. Not available as a formula.
- 95% CI for ρ: percentile bootstrap, 2000 resamples, seed 1.
- Pearson's r on metres vs. rank: reference only. It is not suitable here because of the ordinal, skewed scale.

Seven analyses are run:

1. and 4. Odour vs. category, all activities (B1, B2).
2. and 5. Only activities where odour is **not** decisive (odour < max distance), so the category was set by another factor. A residual truncation effect remains, because odour ≤ category distance still holds.
3. and 6. Odour vs. the category recomputed **without** odour, i.e. from max(stof, geluid, gevaar). This analysis has no structural coupling. It tests whether activities with a large odour distance also tend to have large other distances.
7. Union of B1 and B2, duplicates removed.

**Permutation baseline** (analyses 1, 4 and 7). The script shuffles odour distances across activities, rebuilds the category as max(shuffled odour, other distances), and computes ρ. This is repeated 1000 times with seed 1. The mean is the correlation that the `MAX` construction alone produces if odour were unrelated to the other nuisances.

### 6. Correlation diagrams

Two count heatmaps of category × odour distance, both also embedded on the `Figure` sheet:

- `analysis/figures/correlatiediagram_geur_categorie.png`: odour distance on the x-axis, category on the y-axis.
- `analysis/figures/correlatiediagram_categorie_geur.png`: the same data with the axes reversed, i.e. category on the x-axis and odour distance on the y-axis. Both variables are discrete, so a scatter plot would hide thousands of overlapping points. A count heatmap shows the joint distribution directly.

- Outlined cells: odour distance = category distance (odour is decisive).
- Hatched cells: impossible by construction.

`Crosstab_B1` and `Crosstab_B2` contain the underlying counts as `COUNTIFS` formulas, with a check that the totals equal n.

## Workbook conventions

- Blue text: a value copied from a source, or computed in Python.
- Black text: a live formula.
- Yellow fill: a verification column.
- Font: Arial.

The `README` sheet describes every sheet.

## Software

Python 3 with openpyxl, scipy 1.18 and matplotlib 3.11, and LibreOffice for recalculation.

## References

- VNG (2009). *Bedrijven en milieuzonering*. Den Haag: VNG uitgeverij. Bijlage 1 (Staat van bedrijfsactiviteiten) and Bijlage 2 (Staat van bedrijfsactiviteiten bedrijventerreinen); the largest distance determines the category.
- Agresti, A. (2010). *Analysis of Ordinal Categorical Data* (2nd ed.). Wiley. Rank-based association for ordinal data with ties.
- Kendall, M. G. (1945). The treatment of ties in ranking problems. *Biometrika*, 33(3), 239–251. Kendall's τ-b.
- SciPy documentation: [`scipy.stats.spearmanr`](https://docs.scipy.org/doc/scipy/reference/generated/scipy.stats.spearmanr.html) and [`scipy.stats.kendalltau`](https://docs.scipy.org/doc/scipy/reference/generated/scipy.stats.kendalltau.html).
