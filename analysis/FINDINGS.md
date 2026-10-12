# Findings: odour distance vs. environmental category

Data: `analysis/geur_categorie_analyse.xlsx`. Method: `METHODOLOGY.md`. Figures (all in `figures/`, also on the `Figure` sheet):

- count heatmaps: `correlatiediagram_geur_categorie.png` (odour on the x-axis) and `correlatiediagram_categorie_geur.png` (axes reversed);
- bubble diagrams with regression line: `bubbel_geur_categorie.png` (odour on the x-axis) and `bubbel_categorie_geur.png` (axes reversed).

## Summary

1. There is a moderately strong, positive, monotonic association between the environmental category and the odour distance: Spearman ρ ≈ 0.62 and Kendall τ-b ≈ 0.52 in both lists.
2. Part of this association is built in. The category is the maximum of the geur, stof, geluid and gevaar distances, so the category distance can never be smaller than the odour distance. The `MAX` construction alone would give ρ ≈ 0.39 (B1) and 0.46 (B2). The observed values lie clearly above that baseline.
3. Beyond the built-in part, there is a real co-variation. The odour distance correlates with the category that the other three factors alone would produce: ρ = 0.47 (B1) and 0.39 (B2). Activities with large odour distances also tend to have large dust, noise or hazard distances.
4. In practice, the category gives a firm **upper bound** on the odour distance, but a poor estimate of it. Within one category, odour distances range from 0 m up to the category distance. Odour determines the category in only about a third of activities.

## Core results

| # | Analysis | n | Spearman ρ (95% CI) | Kendall τ-b | Baseline ρ from `MAX` alone (95% range) |
|---|---|---|---|---|---|
| 1 | B1: odour vs. category | 559 | 0.618 (0.553–0.675) | 0.517 | 0.393 (0.334–0.447) |
| 2 | B1: only where odour is not decisive | 379 | 0.640 (0.572–0.707) | 0.537 | – |
| 3 | B1: odour vs. category without odour | 559 | 0.474 (0.404–0.540) | 0.373 | – |
| 4 | B2: odour vs. category | 315 | 0.610 (0.521–0.692) | 0.529 | 0.459 (0.391–0.528) |
| 5 | B2: only where odour is not decisive | 198 | 0.561 (0.440–0.665) | 0.487 | – |
| 6 | B2: odour vs. category without odour | 315 | 0.385 (0.287–0.484) | 0.311 | – |
| 7 | B1 ∪ B2 (duplicates removed) | 562 | 0.618 (0.556–0.675) | 0.517 | 0.394 (0.331–0.451) |

All p-values are below 10⁻¹¹; the exact values are on the `Statistics` sheet. For completeness, Pearson r on metres vs. rank is 0.51 (B1) and 0.60 (B2). It is not the preferred measure here.

## Distribution (B1)

| Category | n | Mean odour distance (m) | Odour decisive | Odour = 0 m |
|---|---|---|---|---|
| 1 | 43 | 1 | 5 | 38 |
| 2 | 91 | 12 | 25 | 33 |
| 3.1 | 82 | 26 | 26 | 15 |
| 3.2 | 100 | 51 | 34 | 14 |
| 4.1 | 64 | 118 | 29 | 3 |
| 4.2 | 84 | 173 | 37 | 8 |
| 5.1 | 44 | 172 | 10 | 4 |
| 5.2 | 24 | 286 | 8 | 0 |
| 5.3 | 15 | 279 | 3 | 0 |
| 6 | 12 | 404 | 3 | 2 |
| **Total** | **559** | | **180 (32%)** | **117 (21%)** |

In B2, odour is decisive in 117 of 315 activities (37%). B2 contains no activities in categories 5.1–5.3. Its maximum is 4.2, apart from six activities in category 6.

## Interpretation

- **Direction and strength.** Higher categories go with larger odour distances. The mean rises from about 1 m in category 1 to about 400 m in category 6. The ordering is not strict at the top: categories 4.2 and 5.1, and 5.2 and 5.3, have nearly equal means.
- **Built-in part.** Because the category is a maximum, one triangle of the diagram (hatched) is necessarily empty, which by itself produces a positive correlation. The permutation baseline (ρ ≈ 0.39–0.46) estimates the size of that effect. The observed ρ of about 0.61–0.62 lies above the baseline's 95% range in both lists.
- **Substantive part.** Analyses 3 and 6 are free of this coupling. They show a moderate positive correlation between odour and the category from stof, geluid and gevaar. This reflects how the activity types are composed (e.g. food industry, waste processing and chemistry are high on several factors at once). It is not an arithmetic effect.
- **Subset where odour is not decisive (analyses 2 and 5).** ρ stays at about 0.56–0.64. This subset still carries a truncation effect, because odour < category distance, so it cannot be read as a clean "independent" correlation. Analyses 3 and 6 are the cleaner test.
- **Heterogeneity.** Odour determines the category mainly in the middle range (3.1–4.2). From 5.1 upward, it is mostly another factor (stof, geluid or gevaar): in B1, odour is decisive in 24 of the 95 activities in categories 5.1–6. Odour distances above 300 m are rare (24 activities in B1).

## Regression lines (bubble diagrams)

Ordinary least squares on the step positions of both axes: odour 0, 10, 30, … 1500 m = steps 0–10; category 1, 2, 3.1, … 6 = steps 0–9. Each activity counts once. Values are on the `Regression` sheet, as live formulas cross-checked against Python.

| Dataset | Regression | Intercept | Slope (SE) | R² |
|---|---|---|---|---|
| B1 | category step on odour step | 1.82 | 0.60 (0.035) | 0.36 |
| B1 | odour step on category step | 0.60 | 0.59 (0.034) | 0.36 |
| B2 | category step on odour step | 1.98 | 0.47 (0.035) | 0.37 |
| B2 | odour step on category step | 0.26 | 0.78 (0.058) | 0.37 |

Reading:

- In B1, one odour step up (e.g. 30 → 50 m) goes with about 0.6 category step up, and vice versa. The line explains about 36% of the variance in step position (r ≈ 0.60 on the step scale, close to Spearman ρ = 0.62).
- The two lines per dataset are not each other's inverse; this is regression to the mean, and the reason both orientations are shown. In B2 the slopes differ more, because B2 has no activities in categories 5.1–5.3.
- The line is descriptive only. It treats the steps as equally spaced, which they are not in metres. It also ignores the built-in constraint (odour ≤ category distance): in the odour-on-x diagram the B1 line runs into the impossible region beyond 700 m. The rank correlations above remain the primary measure; the regression line shows the trend visually.

## Caveats

- **This is not a sample.** The lists are the full catalogue of the VNG guideline. The p-values are therefore not inference to a population. ρ and τ-b are best read as descriptive measures of the catalogue's internal structure.
- **B2 is not independent of B1.** 312 of the 315 B2 activities appear in B1 with identical values. The two results confirm each other, but they are not independent replications.
- **Ties.** Both variables have few distinct values: 11 odour steps and 10 categories. τ-b and mid-rank ρ correct for ties. Even so, the CI and p-values remain approximations.
- **Excluded rows in B1.** Rows 763–837 are a targeted cut-and-paste of columns E and R from an updated version of bijlage 1. That version was published on the VNG website for the former guideline and can no longer be found. The rows have no description and no other distances, so they cannot be linked to an activity or verified, and they are excluded. They do not match any contiguous part of the 2009 table. Sensitivity check: those 47 rows alone give ρ = 0.33 (p = 0.02); B1 including them gives n = 606 and ρ = 0.61. The conclusion does not change.
- **Hardcoded categories** (15 in B1, 7 in B2) all match the category recomputed from the distances. Cached O and R values match recalculation for every included row.

## References

- VNG (2009). *Bedrijven en milieuzonering*. Den Haag: VNG uitgeverij.
- Agresti, A. (2010). *Analysis of Ordinal Categorical Data* (2nd ed.). Wiley.
- Kendall, M. G. (1945). The treatment of ties in ranking problems. *Biometrika*, 33(3), 239–251.
- Efron, B., & Tibshirani, R. J. (1993). *An Introduction to the Bootstrap*. Chapman & Hall.
