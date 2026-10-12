# Methodology: odour distance vs. environmental category

## Question

Is there an association between the environmental category (*milieucategorie*) of an activity and its guidance distance for odour (*geurafstand*) in the VNG 2009 *Staat van bedrijfsactiviteiten*?

## Sources (read-only)

All three files are the original VNG 2009 Excel appendices. They are listed on the VNG publication page [Handreiking Bedrijven en milieuzonering](https://vng.nl/publicaties/handreiking-bedrijven-en-milieuzonering).

| Code | File in `doc/` | Sheet | Geur | Stof | Geluid | Gevaar | Grootste afstand | Categorie |
|---|---|---|---|---|---|---|---|---|
| T1 | `bijlage_1_richtafstandenlijst_tabel_1_versie_maart_2009_def.xls`: bijlage 1, tabel 1, bedrijfsactiviteiten | `EINDTABEL 1 met formule` | E | F | G | J | O | R |
| T2 | `bijlage_1_richtafstandenlijst_tabel_2_versie_maart_2009_def.xls`: bijlage 1, tabel 2, opslagen en installaties | `TABEL2` | D | E | F | I | N | P |
| B2 | `bijlage_2_staat_bedrijfsactiviteiten_bedrijventerreinen_versie_maart_2009_def.xlsx`: bijlage 2, bedrijventerreinen | `EINDTABEL 1 met formule` | E | F | G | J | O | R |

**Datasets analysed.** T1, T2 and B2 separately, plus:

- **B1 = T1 + T2**, bijlage 1 as a whole;
- **B1 ∪ B2**: B1 plus the bijlage 2 activities that are not in T1.

**Column notes.**

- In T1 and B2, columns L, M, N and P are hidden. `CATEGORIE` (R) is the 14th visible column, i.e. the "N" seen on screen.
- In T2, the category is in column P (`CATEGORIE`), as specified. Column Q, also headed `CATEGORIE`, is empty.

**Provenance.**

- **T1 and T2:** downloaded from the VNG publication page. T2 was downloaded manually by the owner, because the VNG server refuses automated downloads of that file.
- **B2:** supplied by the owner as `.xlsx`. Its name matches the VNG file `bijlage_2_staat_bedrijfsactiviteiten_bedrijventerreinen_versie_maart_2009_def.xls`; this was not checked further.
- **Earlier combined file:** an earlier version of this analysis used a combined file, `milieuzonering_bijlage_1_deel_1_en_2.xlsx`. It has been removed from `doc/`; Git history keeps a copy.
  - Its rows 1–762 are identical to T1, in both values and formulas; this was checked cell by cell.
  - Its rows 763–837 were columns D and P of T2, pasted into columns E and R.

The files in `doc/` are never written. `build_analysis.py` checks their SHA-256 hashes before and after each run, and records them on the `Statistics` sheet.

## How the category is computed in the sources

**T1 and B2** use formulas:

1. `O` (`GROOTSTE AFSTAND`) = `MAX(J, G, F, E)`, the largest of the distances for gevaar, geluid, stof and geur.
2. `P` (`CAT`) = `VLOOKUP(O, codetabel!A1:B12, 2, FALSE)`.
3. `R` (`CATEGORIE`) = `IF(ISNA(P), "", P)`.

**T2** contains no formulas. Its distances, its `GROOTSTE AFSTAND` (N) and its category (P) are typed values. Most are stored as text, e.g. `'30'`, and a `-` means the aspect is not relevant. The same rule nevertheless applies to T2.

The code table is identical in all three files: 10 m → 1, 30 → 2, 50 → 3.1, 100 → 3.2, 200 → 4.1, 300 → 4.2, 500 → 5.1, 700 → 5.2, 1000 → 5.3, 1500 → 6. This matches the VNG rule that the largest distance determines the category.

The category is therefore not independent of the odour distance, because odour is one of the four inputs to the maximum. This structural dependency is central to how the correlation should be read (see step 5).

Deviations from the pattern, all preserved and checked:

- **Hardcoded categories.** T1 has 15 included rows, and B2 has 7, where R is a hardcoded value rather than a formula. Examples are T1 rows 430, 454, 465 and 472–474 (wind turbines) and 618 (helipads). All of these match the recalculation.
- **Irregular O and text distances.** Some rows have O hardcoded or written as `MAX(E:J)`, or have distances stored as text.
- **External link.** B2 rows 423–424 look up the category through an external link (`[1]codetabel`). The cached values are used.
- **One inconsistency in T2.** Row 16, "bovengronds, K3-klasse: < 10 m3", has geur 10, gevaar 10 and stof/geluid `-`. It nevertheless lists `GROOTSTE AFSTAND` 30 and category 2, where the rule gives 10 and category 1. The category as published (2) is used, and the mismatch is flagged on `T2_extract`.

## Processing steps

All steps are in `analysis/build_analysis.py`. You can reproduce them with:

```
python3 analysis/build_analysis.py
python3 <xlsx-skill>/scripts/recalc.py analysis/geur_categorie_analyse.xlsx 180
```

The second command recalculates all formulas with LibreOffice so that cached values exist. It reported 38,901 formulas and 0 errors.

### 1. Extraction (intermediate sheets `T1_extract`, `T2_extract`, `B2_extract`)

**Reading `.xls` files (T1, T2).** The values that Excel displays are read with `xlrd`. These are Excel's own stored results, so no recalculation is involved. The formula text is read from a temporary LibreOffice conversion to `.xlsx`, which is discarded after the run.

**Reading `.xlsx` (B2).** The file is read with `openpyxl` twice: once for the formulas and once for the cached values.

For every non-empty row from row 3 down, the extract sheet holds:

- the identifiers: SBI codes and number for T1 and B2; Nr. and subnr. for T2;
- the description;
- the four raw distances, exactly as stored (number, text or `-`);
- the source content of the max-distance and category columns (formula text or typed value);
- the displayed (cached) values. The cached category is the value used in the analysis, normalised to a text label, e.g. `2` → `"2"`.

The header row of each extract sheet names the source column letters.

### 2. Verification (yellow formula columns on the extract sheets)

Live formulas recompute the source logic independently:

- `Geur/Stof/Geluid/Gevaar num`: text to number. `-` and empty cells become empty.
- `Max distance recalc` = `MAX` of the four numbers.
- `Category recalc` = `INDEX/MATCH` on the `Codetabel` sheet.
- `O check` and `R check` compare the cached source values with the recalculated ones. They show `n/a (excluded)` for rows outside the dataset.

**Results:**

- **T1:** 559 of 559 OK.
- **B2:** 315 of 315 OK.
- **T2:** 46 of 47 OK. The exception is row 16 (see above).

A summary is on `Statistics`, under "Verification of source values".

### 3. Inclusion rule

A row is included if it has:

- (a) a description;
- (b) a numeric odour distance, where `0` counts as a value and `-` and empty do not;
- (c) a category.

The column `Exclusion reason` states why any other row is excluded. Excluded rows are section headers, for example "Slachterijen en overige vleesverwerking:" or "OPSLAGEN". In T2 they also include storage types whose odour distance is `-`, such as LPG tanks and explosives, which only have a hazard distance.

| | T1 | T2 | B2 |
|---|---|---|---|
| rows extracted | 760 | 76 | 467 |
| no numeric odour distance (headers, `-`) | 149 | 29 | 117 |
| no description | 52 | 0 | 35 |
| **included** | **559** | **47** | **315** |

### 4. Final datasets (values only)

`Data_T1`, `Data_T2`, `Data_B1`, `Data_B2` and `Data_B1_union_B2` contain values only: one row per activity, with odour distance (m), category label and ordinal rank (1–10). They also hold these derived variables:

- `Geur decisive`: whether the geur distance equals the maximum distance, i.e. whether odour determines the category.
- `Max other distances` and `Category rank excl. geur`: the category the activity would have on stof, geluid and gevaar alone (rank 0 if all three are 0 m or `-`).
- `Also in bijlage 2 / tabel 1`: matched on SBI code, sub-number and description.
  - 312 of the 315 B2 activities also appear in T1, with identical geur and category values, so B2 is almost a subset of T1.
  - T2 has no counterpart in B2.
  - The union B1 ∪ B2 therefore adds only 3 activities (n = 609). B1 and B2 are analysed separately and should not be pooled as independent samples.

Columns on the dataset sheets that hold formulas:

- **N–R:** mid-ranks (average rank for ties), `COUNTIF(range,"<"&x) + (COUNTIF(range,x)+1)/2`, used by `Statistics`.
- **S–T:** step positions, used by `Regression`.

### 5. Statistics (`Statistics` sheet)

Both variables are ordinal with many ties. The odour distance is measured on the same stepped, roughly logarithmic scale as the category. The analysis therefore uses rank correlations:

- **Spearman's ρ:** computed as `CORREL` of the mid-rank columns, as a live formula, and cross-checked against `scipy.stats.spearmanr`. All 13 analyses match to below 10⁻⁴.
- **Kendall's τ-b:** `scipy.stats.kendalltau` with `variant="b"`, which corrects for ties. Not available as a formula.
- **95% CI for ρ:** percentile bootstrap, 2000 resamples, seed 1.
- **Pearson's r** on metres vs. rank: reference only. It is not suitable here because of the ordinal, skewed scale.

Three analyses are run for each of B1, T1, T2 and B2:

1. **Odour vs. category, all activities.**
2. **Only activities where odour is not decisive** (odour < max distance), so the category was set by another factor. A residual truncation effect remains, because odour ≤ category distance still holds.
3. **Odour vs. the category recomputed without odour**, i.e. from max(stof, geluid, gevaar). This analysis has no structural coupling. It tests whether activities with a large odour distance also tend to have large other distances.

A 13th analysis covers the union B1 ∪ B2.

**Permutation baseline** (each "all activities" analysis):

- The script shuffles odour distances across activities, rebuilds the category as max(shuffled odour, other distances), and computes ρ.
- This is repeated 1000 times with seed 1.
- The mean is the correlation that the `MAX` construction alone produces if odour were unrelated to the other nuisances.

### 6. Correlation diagrams

Two kinds of diagram, each in two orientations, all with four panels: T1, T2, B1 and B2. All are in `analysis/figures/` and are also embedded on the `Figure` sheet.

**Count heatmaps.**

- `correlatiediagram_geur_categorie.png`: odour distance on the x-axis, category on the y-axis.
- `correlatiediagram_categorie_geur.png`: the axes reversed.
- Cell value: number of activities. Outlined cells: odour distance = category distance (odour is decisive). Hatched cells: impossible by construction.
- `Crosstab_T1`, `Crosstab_T2`, `Crosstab_B1` and `Crosstab_B2` contain the counts as `COUNTIFS` formulas, with a check that the totals equal n.

**Bubble diagrams with regression line.**

- `bubbel_geur_categorie.png` has odour distance on the x-axis; `bubbel_categorie_geur.png` has the axes reversed.
- Circle area is proportional to the number of activities. Counts of 10 or more are printed in the circle.
- The regression line is ordinary least squares (`scipy.stats.linregress`) on the step positions of both axes: odour steps 0–10 for 0, 10, 30, … 1500 m, and category steps 0–9 for 1, 2, 3.1, … 6.
  - Each activity is one observation, so the fit is implicitly weighted by the circle sizes.
  - The step positions are on the dataset sheets as formulas (columns S and T; the odour steps come from `Codetabel` G:H).
  - The `Regression` sheet recomputes slope, intercept and R² with `SLOPE`, `INTERCEPT` and `RSQ`, and checks them against Python.
- Steps rather than metres are used because the axes are drawn on the step scale. A fit in metres would plot as a curve and be dominated by the few activities at 700–1500 m.

## Workbook conventions

- Blue text: a value copied from a source, or computed in Python.
- Black text: a live formula.
- Yellow fill: a verification column.
- Font: Arial.

The `README` sheet describes every sheet.

## Software

Python 3 with openpyxl, xlrd, scipy 1.18 and matplotlib 3.11, and LibreOffice for the `.xls` formula text and for recalculation.

## References

- VNG (2009). *Bedrijven en milieuzonering*. Den Haag: VNG uitgeverij. Bijlage 1 (Richtafstandenlijsten: tabel 1 bedrijfsactiviteiten, tabel 2 opslagen en installaties) and bijlage 2 (Staat van bedrijfsactiviteiten bedrijventerreinen); the largest distance determines the category. Excel appendices: https://vng.nl/publicaties/handreiking-bedrijven-en-milieuzonering
- Agresti, A. (2010). *Analysis of Ordinal Categorical Data* (2nd ed.). Wiley. Rank-based association for ordinal data with ties.
- Kendall, M. G. (1945). The treatment of ties in ranking problems. *Biometrika*, 33(3), 239–251. Kendall's τ-b.
- SciPy documentation: [`scipy.stats.spearmanr`](https://docs.scipy.org/doc/scipy/reference/generated/scipy.stats.spearmanr.html) and [`scipy.stats.kendalltau`](https://docs.scipy.org/doc/scipy/reference/generated/scipy.stats.kendalltau.html).
