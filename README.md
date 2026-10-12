# Geur en milieucategorie (VNG 2009)

Hangt de milieucategorie van een bedrijfsactiviteit samen met de geurafstand? Dat is onderzocht in de *Staat van bedrijfsactiviteiten* uit VNG, *Bedrijven en milieuzonering* (2009).

## Bronnen (`doc/`)

| Bron | Bestand | n |
|---|---|---|
| Bijlage 1, tabel 1 (bedrijfsactiviteiten) | `bijlage_1_richtafstandenlijst_tabel_1_versie_maart_2009_def.xls` | 559 |
| Bijlage 1, tabel 2 (opslagen en installaties) | `bijlage_1_richtafstandenlijst_tabel_2_versie_maart_2009_def.xls` | 47 |
| Bijlage 2 (bedrijventerreinen) | `bijlage_2_staat_bedrijfsactiviteiten_bedrijventerreinen_versie_maart_2009_def.xlsx` | 315 |

De bestanden komen van de [VNG-publicatiepagina](https://vng.nl/publicaties/handreiking-bedrijven-en-milieuzonering). De n telt de activiteiten met een numerieke geurafstand en een categorie.

## Conclusies

| Bron | Spearman ρ | Conclusie |
|---|---|---|
| Tabel 1 | 0,62 | Matig tot sterk positief verband. Het verband is deels ingebouwd, deels een echte samenhang met de andere hinderaspecten. |
| Tabel 2 | 0,33 | Zwak verband, dat volledig door de berekening van de categorie verklaard wordt. Geur hangt hier niet samen met de andere aspecten (ρ = −0,03). |
| **Tabel 1 + 2** | **0,61** | Matig tot sterk positief verband (95%-BI 0,55–0,67; n = 606). Ongeveer 0,39 daarvan volgt alleen uit de berekening. |
| Bijlage 2 | 0,61 | Zelfde beeld als tabel 1; 312 van de 315 activiteiten staan ook in tabel 1. |

**Waarom het verband deels ingebouwd is.** De categorie is de grootste van vier afstanden: geur, stof, geluid en gevaar. Daardoor is een hoge categorie een gevolg van een grote geurafstand, en niet andersom. De categorie geeft een bovengrens voor de geurafstand, maar is er een slechte schatter van.

## Meer informatie

- [`analysis/METHODOLOGY.md`](analysis/METHODOLOGY.md): methode en bewerkingen.
- [`analysis/FINDINGS.md`](analysis/FINDINGS.md): volledige resultaten en kanttekeningen.
- `analysis/geur_categorie_analyse.xlsx` en `analysis/figures/`: data en figuren.
