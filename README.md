# MH Heures — Système complet de gestion des heures
## MH Maler & Gipser AG · Gstaad

---

## FICHIERS INCLUS

```
MH_PWA/
├── index.html        ← Application PWA (ouvrir sur mobile)
├── manifest.json     ← Config PWA (icône, nom, couleurs)
├── sw.js             ← Service Worker (mode hors-ligne)
├── icon-192.svg      ← Icône app
├── icon-512.svg      ← Icône app (grande)
├── import_csv.py     ← Script Python d'import CSV→Excel
└── README.md         ← Ce fichier
```

---

## UTILISATION MOBILE (PWA)

### Installation sur iPhone/iPad
1. Ouvrir `index.html` dans **Safari**
2. Appuyer sur l'icône de partage ↑
3. Sélectionner **"Sur l'écran d'accueil"**
4. L'app s'installe comme une vraie application

### Configuration initiale (Patron)
- Appuyer sur ⚙ sur l'écran PIN
- Entrer le PIN admin : **1234** ← _à changer dans index.html_
- Créer les employés avec leurs PINs personnels

### Saisie quotidienne (Employé)
1. Entrer son PIN personnel
2. Sélectionner la date
3. Entrer les heures de début/fin (3 périodes max)
4. Optionnel: code absence, chantier, km, commentaire
5. **Enregistrer**

### Export mensuel
1. Aller dans l'onglet **Export**
2. Sélectionner le mois/année
3. Appuyer **"Exporter CSV → Excel"**
4. Envoyer/télécharger le fichier `.csv`

---

## IMPORT CSV → EXCEL

### Prérequis
```bash
pip install openpyxl
```

### Usage simple
```bash
# Import automatique (cherche l'Excel dans le dossier courant)
python import_csv.py MH_Import_Luis_Fruela_Juin_2026.csv

# Import avec Excel spécifié
python import_csv.py MH_Import_Luis_Fruela_Juin_2026.csv MH_Gestion_Heures_2026_PRO.xlsx
```

### Ce que fait l'import
- ✅ Lit le CSV exporté depuis la PWA
- ✅ Détecte automatiquement le mois, l'année, l'employé
- ✅ Trouve l'onglet mensuel correspondant (Jan, Fév…)
- ✅ Écrit les heures dans les bonnes cellules
- ✅ Écrit les codes d'absence (V, M, A…)
- ✅ Écrit chantier, km, commentaires
- ✅ Crée un backup avant modification
- ✅ Colorie les cellules importées (bleu clair)
- ⚠️ Les formules Excel restent intactes (H. nettes calculées par Excel)

### Après l'import
Dans Excel, forcer le recalcul avec **Ctrl+Alt+F9** (Windows) ou **Cmd+Alt+F9** (Mac).

---

## FORMAT CSV

Le CSV généré a ce format (reconnu par le script) :
```
## MH_IMPORT_V2
## EMPLOYE:Luis Fruela
## MOIS:Juin
## ANNEE:2026
## PERIODE:2026-06

DATE_ISO,JOUR,DEB1,FIN1,DEB2,FIN2,DEB3,FIN3,HEURES,CODE,CHANTIER,KM,COMMENTAIRE
2026-06-01,Lu,07:00,12:00,13:15,17:30,,,9.25,,Chalet Dubois,12,
2026-06-08,Lu,,,,,,,,V,,0,Vacances posées
```

---

## CORRESPONDANCE COLONNES

| CSV          | Excel colonne | Description            |
|-------------|---------------|------------------------|
| DATE_ISO    | A             | Date (format YYYY-MM-DD) |
| DEB1/FIN1   | C / D         | Période 1              |
| DEB2/FIN2   | E / F         | Période 2              |
| DEB3/FIN3   | G / H         | Période 3              |
| CODE        | J             | Code absence (V, M, A…)|
| COMMENTAIRE | K             | Commentaire libre      |
| CHANTIER    | L             | Nom du chantier        |
| KM          | N             | Kilomètres trajet      |

Les heures nettes (col. I) sont calculées **automatiquement** par les formules Excel.

---

## SÉCURITÉ

- Les données sont stockées **localement** sur l'appareil (localStorage)
- Aucune donnée n'est envoyée à un serveur externe
- Les PINs sont stockés en clair dans le localStorage — ne pas utiliser des PINs bancaires
- Pour supprimer toutes les données : vider le cache du navigateur

---

## CODES D'ABSENCE

| Code | Description                              |
|------|------------------------------------------|
| V    | Vacances                                 |
| M    | Maladie (avec certificat médical)        |
| A    | Accident (LAA)                           |
| Mil  | Militaire / Service civil / Maternité   |
| AP   | Absence payée                            |
| C    | Correction manuelle                      |
| F    | Jour férié officiel                      |

---

_MH Maler & Gipser AG · Eisbahnweg 3 · 3780 Gstaad · MWST CHE-163.519.509_
