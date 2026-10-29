# Contrôle Stocks & Immobilisations

Logiciel Windows (.exe) : import du journal comptable, génération automatique des stocks (601/602 → 31/32) avec CMUP et fiches d'inventaire, et des fichiers d'immobilisations (comptes 21 à 24) avec registre imprimable.

## Obtenir le .exe via GitHub
1. Créez un dépôt GitHub et envoyez-y tout ce dossier (y compris `.github/workflows/build.yml`).
2. Onglet **Actions** → « Compiler le .exe Windows » (se lance à chaque push, ou bouton *Run workflow*).
3. À la fin, ouvrez l'exécution → **Artifacts** → `ControleStocksImmos-windows` → téléchargez le `.exe`.
4. Option : créez un tag `v1.0` (`git tag v1.0 && git push --tags`) et le .exe est publié dans **Releases**.

## Utilisation
1. **Journal** : importez un export Excel (.xlsx) ou CSV avec une ligne d'en-tête : Date, Journal, Pièce, Compte, Libellé, Débit, Crédit.
2. **Générer** : comptes 601xxx → stock 311xxx, 602xxx → stock 321xxx ; ventes 701xxx → sorties du stock 311xxx, 702xxx → sorties du stock 321xxx ; comptes 21/22/23/24 → fiches d'immobilisations.
   (Le compte de stock de chaque ligne reste modifiable. Si vos 702 correspondent à des produits finis, rattachez-les au stock 36 en éditant « Compte stock ».)
3. **Achats → Stocks** : saisissez les quantités (double-clic). Renommez « Article » (multi-sélection possible) pour regrouper plusieurs achats d'un même article.
4. **Ventes → Sorties de stock** : chaque vente 701/702 est rattachée automatiquement à l'article d'achat le plus proche (mots-clés du libellé ; boutons « Associer à un article » et « Rattachement automatique » pour corriger). Saisissez la quantité vendue : coût de sortie = quantité × CMUP, marge brute = prix de vente − coût.
5. **Inventaire & CMUP** : stock initial (optionnel), quantité inventoriée ; CMUP = (valeur initiale + achats) / (qté initiale + qté achetée) ; stock final = qté inventoriée × CMUP. Export PDF (fiches par compte, fiches individuelles) et Excel, avec écritures de variation de stocks proposées.
6. **Immobilisations** : pour chaque acquisition, composantes avec référence, sous-référence, quantité, valeur, durée, date de mise en service, localisation. « Éclater en unités » crée N sous-références.
7. **Registre** : par compte comptable ou par composante, situation à une date (amortissement linéaire prorata jours), impression PDF / Excel.

Les données sont enregistrées automatiquement dans `Documents/ControleStocksImmos/projet.db` (menu Fichier pour créer/ouvrir d'autres projets). Une nouvelle génération après un nouvel import conserve les quantités, articles, références et composantes déjà saisis.

## Test en local
```
pip install -r requirements.txt
python app.py
```
Un fichier `journal_exemple.csv` est fourni pour tester.
