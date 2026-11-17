# Pricer MTN et blotter de booking – Excel VBA

Outil d'assistant trader pour un desk de trésorerie long terme : pricer une émission de MTN
(taux fixe ou FRN) sur la courbe de financement de l'émetteur, puis l'ajouter à un blotter,
la contrôler avant booking et exporter les trades validés.

*English summary at the bottom.*

## Ce que fait le classeur

| Feuille | Contenu |
|---|---|
| **Market** | Courbe swap EUR (taux par 1–30 ans), spread de financement de l'émetteur par maturité, Euribor 6M. Cellules jaunes = saisies. |
| **Curve** | Bootstrap annuel des facteurs d'actualisation à partir des taux swap par, taux zéro, forwards 1 an, facteurs d'actualisation émetteur (swap + spread). |
| **Pricer** | Coupon fixe au pair, prix au coupon proposé, écart au pair, **coût tout compris swappé (Euribor + bp)**, **marge FRN au pair**, duration, DV01, valeur actuelle. |
| **Cashflows** | Échéancier du MTN : dates, flux, facteurs d'actualisation, valeurs actuelles. |
| **Blotter** | Trades à booker, statut (À valider / Validé / Erreur) et commentaire de contrôle, mis en couleur. |
| **Controles_VBA** | Formules Excel comparées aux fonctions VBA indépendantes. |

Tout le pricing fonctionne **en formules Excel natives**, sans macro. Les modules VBA ajoutent :

- `modPricing` – fonctions de feuille qui **re-calculent tout indépendamment** (bootstrap en VBA) :
  `MTN_COUPON_PAIR`, `MTN_PRIX`, `MTN_MARGE_FRN`, `MTN_TOUT_COMPRIS`, `MTN_DV01` (par bump de
  ±1 bp de toute la courbe et re-bootstrap).
- `modBlotter` – macros `AjouterAuBlotter` (copie le trade du pricer), `ValiderBlotter`
  (ID en double, devise, nominal, maturité 1–30 ans, coupon, écart au pair > 5 bp) et
  `ExporterBooking` (CSV des trades validés, séparateur « ; »).

## Exemple (données illustratives)

Émission 7 ans, 50 M€, coupon proposé 3,35 %, courbe swap et spreads d'exemple :

| Mesure | Valeur |
|---|---|
| Taux swap 7 ans | 2,81 % |
| Spread de financement 7 ans | 55 bp |
| Coupon fixe au pair | 3,359 % |
| Prix au coupon de 3,35 % | 99,94 % (−0,9 bp vs pair) |
| Coût tout compris swappé | Euribor + 54 bp |
| Marge FRN au pair | Euribor + 55,2 bp |
| Duration / DV01 | 6,35 ans / 31 726 € par bp |

> La courbe et les spreads sont des **valeurs d'exemple, pas des données de marché**. Remplacer
> la feuille *Market* par la courbe du desk (ex. Bloomberg EUSA) et la grille de spreads de l'émetteur.

## Contrôles

- `tests/test_mtn.py` (Python, 8 tests) ré-implémente le bootstrap et le pricing et les compare
  aux valeurs calculées par le classeur : chaque swap par est re-pricé à 1 ; le coupon au pair cote
  100 ; un FRN sans spread vaut le pair à Euribor flat ; DV01 par duration à moins de 5 % du bump
  de courbe.
- Les fonctions VBA donnent les mêmes résultats que les formules (écart ≤ 1e-13) ; testé dans
  LibreOffice en mode compatibilité VBA, avec la validation du blotter et l'export CSV.

```bash
pip install openpyxl pytest
pytest -q
```

## Limites

- Coupons annuels, maturités entières, base 30/360 simplifiée ; une seule courbe (pas de
  base 3M/6M ni d'actualisation €STR).
- Spread de financement appliqué en continu sur la courbe swap.
- Le CSV d'export est un format générique, pas un format d'import Summit.

---

## English summary

Excel VBA tool for a long-term funding desk: bootstraps a EUR swap curve, adds the issuer's
funding spread, and prices fixed-rate MTNs and FRNs (par coupon, price, all-in cost swapped to
Euribor, FRN margin, duration, DV01). VBA modules re-implement the pricing independently and run a
pre-booking blotter (validation rules and CSV export). Market inputs are illustrative.
