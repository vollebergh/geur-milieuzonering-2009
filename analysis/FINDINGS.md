# Findings: odour distance vs. environmental category

Data: `analysis/geur_categorie_analyse.xlsx`. Method: `METHODOLOGY.md`.

Figures, all in `figures/` and also on the `Figure` sheet, each with panels for T1, T2, B1 and B2:

- count heatmaps: `correlatiediagram_geur_categorie.png` (odour on the x-axis) and `correlatiediagram_categorie_geur.png` (axes reversed);
- bubble diagrams with regression line: `bubbel_geur_categorie.png` (odour on the x-axis) and `bubbel_categorie_geur.png` (axes reversed).

Datasets:

- **T1:** bijlage 1, tabel 1 (bedrijfsactiviteiten);
- **T2:** bijlage 1, tabel 2 (opslagen en installaties);
- **B1:** T1 + T2, bijlage 1 as a whole;
- **B2:** bijlage 2 (bedrijventerreinen).

All three are the original VNG 2009 Excel files.

## Summary

1. **Association.** Across bijlage 1 as a whole there is a moderately strong, positive, monotonic association between the environmental category and the odour distance: Spearman ρ = 0.61 and Kendall τ-b = 0.51 (n = 606). Tabel 1 (ρ = 0.62) and bijlage 2 (ρ = 0.61) give the same picture.
2. **Built-in part.** The category is the maximum of the geur, stof, geluid and gevaar distances, so the category distance can never be smaller than the odour distance. The `MAX` construction alone would give ρ ≈ 0.39 in bijlage 1 and 0.46 in bijlage 2. The observed values lie clearly above that baseline.
3. **Real co-variation.** Beyond the built-in part, odour distance correlates with the category that the other three factors alone would produce: ρ = 0.46 in bijlage 1 and 0.39 in bijlage 2. Activities with large odour distances also tend to have large dust, noise or hazard distances.
4. **Tabel 2 is different.** In tabel 2 (opslagen en installaties, n = 47) the association is weak: ρ = 0.33, with a CI of 0.03–0.61. That is no larger than the `MAX` construction alone produces (0.31). Odour there is unrelated to the other nuisances (ρ = −0.03). For storage and installations, the link between odour and category is entirely a product of how the category is calculated.
5. **Practical meaning.** The category gives a firm **upper bound** on the odour distance, but a poor estimate of it. Within one category, odour distances range from 0 m up to the category distance. Odour determines the category in only about a third of activities.

## Answering the research question

**Question:** does a higher environmental category lead to a larger odour distance, and how do the two correlate?

**Calculation.** Spearman's rank correlation coefficient (ρ). Both variables are ordinal: the category, and the odour distance, which is a ladder of 11 fixed steps. ρ measures whether a higher category goes with a higher odour distance (a monotonic association) and assumes neither equal spacing nor normality. Pearson's r is therefore not suitable as the primary measure.

**The two numbers (bijlage 1, tabel 1 + tabel 2, n = 606):**

- **ρ = 0.61** (95% CI 0.55–0.67): the direction and strength of the association.
- **p < 10⁻⁶²**: the probability of a ρ at least this large if there were no association at all.

Other values:

- Tabel 1 alone gives ρ = 0.62, p < 10⁻⁵⁹.
- Tabel 2 alone gives ρ = 0.33, p = 0.02.
- Bijlage 2 gives ρ = 0.61, p < 10⁻³².
- Kendall's τ-b (0.51 for bijlage 1) is a robustness check on ρ, not a third key number.

**What the numbers say:**

- **Direction and strength.** ρ is positive: the higher the category, the larger the odour distance tends to be. A value of 0.61 is moderate to strong. Common cut-offs put 0.40–0.69 at moderate and 0.70 and up at strong (Schober et al., 2018). ρ² ≈ 0.37, so about 37% of the variation in ranks is shared.
- **The p-value.** It rules out chance, but has limited meaning here. The lists are the complete VNG catalogue, not a sample, so ρ itself is the main result.

**"Leads to" cannot be answered this way.** There are two reasons:

1. A correlation does not establish cause and effect.
2. In this data the cause runs the other way, by construction. The category is calculated from the largest of four distances, odour among them, so a large odour distance produces a high category, not the reverse.

Part of ρ = 0.61 comes from that calculation alone; the permutation baseline puts it at about 0.39. The remainder is real co-variation: odour also correlates with the category based on stof, geluid and gevaar alone (ρ = 0.46).

**Suggested wording:** "Higher categories are associated with larger odour distances (Spearman ρ = 0.61). This is partly because odour helps determine the category, and partly because activities with a lot of odour nuisance also score high on other nuisances. For storage and installations (tabel 2) only the first applies."

## Core results

| # | Analysis | n | Spearman ρ (95% CI) | Kendall τ-b | Baseline ρ from `MAX` alone (95% range) |
|---|---|---|---|---|---|
| 1 | B1 (T1 + T2): odour vs. category | 606 | 0.610 (0.550–0.666) | 0.511 | 0.393 (0.335–0.447) |
| 2 | B1: only where odour is not decisive | 410 | 0.627 (0.559–0.693) | 0.528 | – |
| 3 | B1: odour vs. category without odour | 606 | 0.459 (0.389–0.522) | 0.362 | – |
| 4 | T1: odour vs. category | 559 | 0.618 (0.553–0.675) | 0.517 | 0.393 (0.334–0.447) |
| 5 | T1: only where odour is not decisive | 379 | 0.640 (0.572–0.707) | 0.537 | – |
| 6 | T1: odour vs. category without odour | 559 | 0.474 (0.404–0.540) | 0.373 | – |
| 7 | T2: odour vs. category | 47 | 0.334 (0.034–0.607) | 0.293 | 0.311 (0.087–0.523) |
| 8 | T2: only where odour is not decisive | 31 | 0.336 (−0.044–0.686) | 0.310 | – |
| 9 | T2: odour vs. category without odour | 47 | −0.029 (−0.346–0.313) | −0.014 | – |
| 10 | B2: odour vs. category | 315 | 0.610 (0.521–0.692) | 0.529 | 0.459 (0.391–0.528) |
| 11 | B2: only where odour is not decisive | 198 | 0.561 (0.440–0.665) | 0.487 | – |
| 12 | B2: odour vs. category without odour | 315 | 0.385 (0.287–0.484) | 0.311 | – |
| 13 | B1 ∪ B2 (duplicates removed) | 609 | 0.610 (0.551–0.666) | 0.511 | 0.394 (0.342–0.449) |

All p-values for B1, T1 and B2 are below 10⁻¹¹. For T2 they are 0.02 (analysis 7), 0.07 (analysis 8) and 0.85 (analysis 9). The exact values are on the `Statistics` sheet. For completeness, Pearson r on metres vs. rank is 0.51 (B1), 0.27 (T2) and 0.60 (B2). It is not the preferred measure here.

## Distribution (bijlage 1, tabel 1 + tabel 2)

| Category | n | of which T2 | Mean odour distance (m) | Odour = 0 m |
|---|---|---|---|---|
| 1 | 48 | 5 | 2 | 40 |
| 2 | 103 | 12 | 12 | 36 |
| 3.1 | 97 | 15 | 26 | 18 |
| 3.2 | 106 | 6 | 50 | 15 |
| 4.1 | 69 | 5 | 117 | 4 |
| 4.2 | 86 | 2 | 169 | 9 |
| 5.1 | 44 | 0 | 172 | 4 |
| 5.2 | 25 | 1 | 276 | 0 |
| 5.3 | 15 | 0 | 279 | 0 |
| 6 | 13 | 1 | 373 | 3 |
| **Total** | **606** | **47** | | **129 (21%)** |

**Odour decisive** (odour distance = maximum of the four recalculated distances):

- B1: 196 of 606 activities (32%), of which T1 180 of 559 and T2 16 of 47.
- B2: 117 of 315 (37%).

B2 contains no activities in categories 5.1–5.3. Its maximum is 4.2, apart from six activities in category 6.

**Tabel 2.** Most of its entries are low categories (1–3.1: 32 of 47) with small odour distances (0–30 m: 37 of 47). Its high categories are set by hazard or noise, not odour:

- radar installation: gevaar 1500 m, odour 0 m;
- storage above the BRZO threshold: gevaar 700 m, odour 30 m.

Storage types without any odour distance (`-`), such as LPG tanks and explosives, are excluded (29 rows, including headers).

## Interpretation

- **Direction and strength.** Higher categories go with larger odour distances. The mean rises from about 2 m in category 1 to about 370 m in category 6. The ordering is not strict at the top: categories 4.2 and 5.1, and 5.2 and 5.3, have nearly equal means.
- **Built-in part.** Because the category is a maximum, one triangle of the diagram (hatched) is necessarily empty, which by itself produces a positive correlation. The permutation baseline (ρ ≈ 0.39–0.46) estimates the size of that effect. For B1, T1 and B2, the observed ρ of about 0.61–0.62 lies above the baseline's 95% range.
- **Substantive part.** Analyses 3, 6 and 12 are free of this coupling. They show a moderate positive correlation between odour and the category from stof, geluid and gevaar. This reflects how the activity types are composed (e.g. food industry, waste processing and chemistry are high on several factors at once). It is not an arithmetic effect.
- **Tabel 2 has no substantive part.** Analysis 9 gives ρ ≈ 0. The observed ρ = 0.33 (analysis 7) equals the construction baseline (0.31). For storage and installations, a high category says nothing about odour beyond the upper bound.
- **Subset where odour is not decisive (analyses 2, 5, 8 and 11).** ρ is about 0.56–0.64 in B1, T1 and B2. This subset still carries a truncation effect, because odour < category distance, so it cannot be read as a clean "independent" correlation. The analyses without odour are the cleaner test.
- **Heterogeneity.** Odour determines the category mainly in the middle range (3.1–4.2). From 5.1 upward, it is mostly another factor (stof, geluid or gevaar): in T1, odour is decisive in 24 of the 95 activities in categories 5.1–6. Odour distances above 300 m are rare (24 activities in B1, all from T1).

## Regression lines (bubble diagrams)

Ordinary least squares on the step positions of both axes: odour 0, 10, 30, … 1500 m = steps 0–10; category 1, 2, 3.1, … 6 = steps 0–9. Each activity counts once. Values are on the `Regression` sheet, as live formulas cross-checked against Python.

| Dataset | Regression | Intercept | Slope (SE) | R² |
|---|---|---|---|---|
| B1 | category step on odour step | 1.78 | 0.60 (0.034) | 0.35 |
| B1 | odour step on category step | 0.62 | 0.57 (0.032) | 0.35 |
| T1 | category step on odour step | 1.82 | 0.60 (0.035) | 0.36 |
| T1 | odour step on category step | 0.60 | 0.59 (0.034) | 0.36 |
| T2 | category step on odour step | 1.80 | 0.29 (0.193) | 0.05 |
| T2 | odour step on category step | 1.20 | 0.16 (0.110) | 0.05 |
| B2 | category step on odour step | 1.98 | 0.47 (0.035) | 0.37 |
| B2 | odour step on category step | 0.26 | 0.78 (0.058) | 0.37 |

Reading:

- **B1 and T1.** One odour step up (e.g. 30 → 50 m) goes with about 0.6 category step up, and vice versa. The line explains about 35% of the variance in step position (r ≈ 0.59–0.60 on the step scale, close to Spearman ρ).
- **T2.** The slope is not significantly different from zero (p = 0.14), with R² = 0.05. This is consistent with the weak rank correlation.
- **Two lines per dataset.** The lines in the two orientations are not each other's inverse; this is regression to the mean, and the reason both orientations are shown. In B2 the slopes differ more, because B2 has no activities in categories 5.1–5.3.
- **Descriptive only.** The line treats the steps as equally spaced, which they are not in metres. It also ignores the built-in constraint (odour ≤ category distance): in the odour-on-x diagram the B1 line runs into the impossible region beyond 700 m. The rank correlations above remain the primary measure; the regression line shows the trend visually.

## Caveats

- **This is not a sample.** The lists are the full catalogue of the VNG guideline. The p-values are therefore not inference to a population. ρ and τ-b are best read as descriptive measures of the catalogue's internal structure.
- **B2 is not independent of B1.** 312 of the 315 B2 activities appear in tabel 1 with identical values. The results confirm each other, but they are not independent replications.
- **Tabel 2 is small.** With n = 47 its confidence intervals are wide. Its result mainly shows that the strong association in tabel 1 does not carry over to storage and installations.
- **Ties.** Both variables have few distinct values: 11 odour steps and 10 categories. τ-b and mid-rank ρ correct for ties. Even so, the CI and p-values remain approximations.
- **One inconsistency in tabel 2.** Row 16 (K3-class tank < 10 m³) lists category 2, while its distances (geur 10, gevaar 10) give category 1. The published category is used. Using category 1 instead changes ρ for tabel 2 from 0.334 to 0.333.
- **Hardcoded categories** (15 in tabel 1, 7 in bijlage 2, and all of tabel 2) match the recalculated category, except tabel 2 row 16. Cached values match recalculation for every other included row.
- **Earlier version.** An earlier version of this analysis used a hand-made combination of tabel 1 and tabel 2, in which tabel 2 had no descriptions and was excluded. Its tabel 1 part was identical to the VNG file, and its 47 usable tabel 2 rows give exactly the tabel 2 result above. The conclusions for tabel 1 and bijlage 2 are unchanged.

## References

- VNG (2009). *Bedrijven en milieuzonering*. Den Haag: VNG uitgeverij. Excel appendices: https://vng.nl/publicaties/handreiking-bedrijven-en-milieuzonering
- Agresti, A. (2010). *Analysis of Ordinal Categorical Data* (2nd ed.). Wiley.
- Schober, P., Boer, C., & Schwarte, L. A. (2018). Correlation coefficients: appropriate use and interpretation. *Anesthesia & Analgesia*, 126(5), 1763–1768. https://doi.org/10.1213/ANE.0000000000002864
- Kendall, M. G. (1945). The treatment of ties in ranking problems. *Biometrika*, 33(3), 239–251.
- Efron, B., & Tibshirani, R. J. (1993). *An Introduction to the Bootstrap*. Chapman & Hall.
