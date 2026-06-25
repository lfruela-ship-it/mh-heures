#!/usr/bin/env python3
"""
import_csv.py — MH Maler & Gipser AG
======================================
Importe un CSV généré par la PWA "MH Gestion des Heures" 
dans le fichier Excel MH_Gestion_Heures_2026_PRO.xlsx

USAGE:
  python import_csv.py <fichier.csv> [fichier_excel.xlsx]

  Si le fichier Excel n'est pas spécifié, cherche automatiquement
  MH_Gestion_Heures_2026_PRO.xlsx dans le répertoire courant.

FONCTIONNEMENT:
  1. Lit le CSV exporté depuis la PWA (format MH_IMPORT_V2)
  2. Identifie le mois/année et l'employé concernés
  3. Trouve la feuille mensuelle correspondante dans l'Excel (Jan, Fév, etc.)
  4. Écrit les heures et codes dans les cellules correctes
  5. Sauvegarde l'Excel mis à jour

CORRESPONDANCE CSV → EXCEL:
  DATE_ISO  → colonne A (date)
  DEB1/FIN1 → colonnes C/D (début/fin période 1)
  DEB2/FIN2 → colonnes E/F (début/fin période 2)
  DEB3/FIN3 → colonnes G/H (début/fin période 3)
  CODE      → colonne J (code d'absence)
  CHANTIER  → colonne L (chantier)
  KM        → colonne N (km)
  COMMENTAIRE → colonne K (commentaire)
"""

import sys
import os
import csv
import datetime
import io
from pathlib import Path

try:
    from openpyxl import load_workbook
    from openpyxl.styles import PatternFill, Font
except ImportError:
    print("❌ openpyxl non installé. Exécutez: pip install openpyxl")
    sys.exit(1)

# ── Mapping mois FR → onglet Excel ────────────────────────────────
SHEET_MAP = {
    1: 'Jan', 2: 'Fév', 3: 'Mar', 4: 'Avr',
    5: 'Mai', 6: 'Jun', 7: 'Jul', 8: 'Aoû',
    9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Déc'
}
MOIS_FR = {
    'Janvier':1, 'Février':2, 'Mars':3, 'Avril':4,
    'Mai':5, 'Juin':6, 'Juillet':7, 'Août':8,
    'Septembre':9, 'Octobre':10, 'Novembre':11, 'Décembre':12
}

# ── Couleurs pour cellules importées ──────────────────────────────
FILL_IMPORTED = PatternFill("solid", fgColor="EAF4FB")   # bleu clair
FILL_SAT      = PatternFill("solid", fgColor="E8E8E8")   # gris samedi
FILL_HOL      = PatternFill("solid", fgColor="DDEEFF")   # bleu férié
FILL_CODE     = PatternFill("solid", fgColor="FFF9C4")   # jaune code

def time_str_to_excel(t_str):
    """Convertit '08:30' en valeur Excel (fraction de jour)."""
    if not t_str or t_str.strip() == '':
        return None
    try:
        h, m = map(int, t_str.strip().split(':'))
        return (h * 60 + m) / (24 * 60)
    except (ValueError, AttributeError):
        return None

def find_date_row(ws, target_date):
    """
    Trouve la ligne Excel correspondant à une date ISO (YYYY-MM-DD).
    La colonne A contient des dates Excel ou des strings formatées.
    """
    target = datetime.date.fromisoformat(target_date)
    for row in ws.iter_rows(min_row=5, max_row=ws.max_row):
        cell = row[0]  # colonne A
        val = cell.value
        if val is None:
            continue
        # Valeur datetime/date Python (openpyxl la lit souvent ainsi)
        if isinstance(val, (datetime.datetime, datetime.date)):
            cell_date = val.date() if isinstance(val, datetime.datetime) else val
            if cell_date == target:
                return cell.row
        # Valeur numérique (numéro de série Excel)
        elif isinstance(val, (int, float)):
            try:
                # Excel date serial: 1 = 1900-01-01 (avec ajustement leap bug)
                excel_date = datetime.date(1899, 12, 30) + datetime.timedelta(days=int(val))
                if excel_date == target:
                    return cell.row
            except (ValueError, OverflowError):
                pass
    return None

def import_csv_to_excel(csv_path, excel_path):
    """Fonction principale d'import."""
    
    print(f"\n{'='*55}")
    print(f"  MH Maler & Gipser AG — Import CSV → Excel")
    print(f"{'='*55}")
    print(f"  CSV   : {csv_path}")
    print(f"  Excel : {excel_path}")
    print(f"{'='*55}\n")
    
    # ── Lire le CSV ───────────────────────────────────────────────
    if not Path(csv_path).exists():
        print(f"❌ Fichier CSV introuvable: {csv_path}")
        sys.exit(1)
    
    with open(csv_path, 'r', encoding='utf-8-sig') as f:
        raw = f.read()
    
    # Vérifier la signature
    if '## MH_IMPORT_V2' not in raw:
        print("⚠️  Ce CSV ne semble pas généré par la PWA MH Heures (signature manquante).")
        print("   L'import continuera quand même si le format des colonnes est correct.")
    
    # Extraire les métadonnées
    employe = None
    periode = None
    annee = None
    mois_num = None
    
    for line in raw.split('\n'):
        line = line.strip()
        if line.startswith('## EMPLOYE:'):
            employe = line.replace('## EMPLOYE:', '').strip()
        elif line.startswith('## PERIODE:'):
            periode = line.replace('## PERIODE:', '').strip()  # ex: "2026-06"
        elif line.startswith('## ANNEE:'):
            try:
                annee = int(line.replace('## ANNEE:', '').strip())
            except ValueError:
                pass
        elif line.startswith('## MOIS:'):
            mois_str = line.replace('## MOIS:', '').strip()
            mois_num = MOIS_FR.get(mois_str)
    
    # Déduire mois/année depuis PERIODE si nécessaire
    if periode and not (annee and mois_num):
        try:
            y, m = map(int, periode.split('-'))
            annee = annee or y
            mois_num = mois_num or m
        except (ValueError, AttributeError):
            pass
    
    # Lire les lignes de données
    lines = [l for l in raw.split('\n') if not l.startswith('#') and l.strip()]
    if not lines:
        print("❌ Aucune donnée trouvée dans le CSV.")
        sys.exit(1)
    
    reader = csv.DictReader(lines)
    rows = list(reader)
    
    if not rows:
        print("❌ Aucune ligne de données dans le CSV.")
        sys.exit(1)
    
    # Déduire mois/année depuis les données si toujours manquant
    if not (annee and mois_num):
        first_date = rows[0].get('DATE_ISO', '').strip()
        if first_date:
            try:
                d = datetime.date.fromisoformat(first_date)
                annee = annee or d.year
                mois_num = mois_num or d.month
            except ValueError:
                pass
    
    if not (annee and mois_num):
        print("❌ Impossible de déterminer le mois/année du CSV.")
        sys.exit(1)
    
    sheet_name = SHEET_MAP.get(mois_num)
    if not sheet_name:
        print(f"❌ Mois invalide: {mois_num}")
        sys.exit(1)
    
    print(f"  Employé  : {employe or 'Non précisé'}")
    print(f"  Période  : {mois_num:02d}/{annee}")
    print(f"  Feuille  : {sheet_name}")
    print(f"  Lignes   : {len(rows)}\n")
    
    # ── Charger l'Excel ───────────────────────────────────────────
    if not Path(excel_path).exists():
        print(f"❌ Fichier Excel introuvable: {excel_path}")
        sys.exit(1)
    
    wb = load_workbook(excel_path)
    
    if sheet_name not in wb.sheetnames:
        print(f"❌ Onglet '{sheet_name}' non trouvé dans l'Excel.")
        print(f"   Onglets disponibles: {', '.join(wb.sheetnames)}")
        sys.exit(1)
    
    ws = wb[sheet_name]
    
    # ── Importer chaque ligne ────────────────────────────────────
    imported = 0
    skipped  = 0
    errors   = []
    
    for row_data in rows:
        date_iso = row_data.get('DATE_ISO', '').strip()
        if not date_iso:
            skipped += 1
            continue
        
        # Validation date
        try:
            target_date = datetime.date.fromisoformat(date_iso)
        except ValueError:
            errors.append(f"Date invalide: {date_iso}")
            skipped += 1
            continue
        
        # Trouver la ligne Excel
        excel_row = find_date_row(ws, date_iso)
        if excel_row is None:
            errors.append(f"Date {date_iso} non trouvée dans l'onglet {sheet_name}")
            skipped += 1
            continue
        
        # ── Écrire les valeurs ────────────────────────────────────
        is_saturday = target_date.weekday() == 5
        base_fill = FILL_SAT if is_saturday else FILL_IMPORTED
        
        # Périodes horaires (colonnes C..H = indices 2..7)
        time_cols = [
            ('DEB1', 2), ('FIN1', 3),
            ('DEB2', 4), ('FIN2', 5),
            ('DEB3', 6), ('FIN3', 7),
        ]
        for key, col_idx in time_cols:
            val_str = row_data.get(key, '').strip()
            excel_val = time_str_to_excel(val_str)
            cell = ws.cell(row=excel_row, column=col_idx+1)
            if excel_val is not None:
                cell.value = excel_val
                cell.number_format = 'H:MM'
                cell.fill = base_fill
            else:
                cell.value = None
                cell.fill = base_fill
        
        # Code absence (colonne J = 10)
        code = row_data.get('CODE', '').strip()
        code_cell = ws.cell(row=excel_row, column=10)
        if code:
            code_cell.value = code
            code_cell.fill = FILL_CODE
            code_cell.font = Font(bold=True, color='E8532B')
        else:
            code_cell.fill = base_fill
        
        # Commentaire (colonne K = 11)
        cmt = row_data.get('COMMENTAIRE', '').strip()
        ws.cell(row=excel_row, column=11).value = cmt or None
        ws.cell(row=excel_row, column=11).fill = base_fill
        
        # Chantier (colonne L = 12)
        chantier = row_data.get('CHANTIER', '').strip()
        ws.cell(row=excel_row, column=12).value = chantier or None
        ws.cell(row=excel_row, column=12).fill = base_fill
        
        # Km (colonne N = 14)
        km_str = row_data.get('KM', '').strip()
        try:
            km = int(km_str) if km_str else 0
        except ValueError:
            km = 0
        if km > 0:
            ws.cell(row=excel_row, column=14).value = km
        
        imported += 1
        status = f"✓ {date_iso} → ligne {excel_row}"
        if code:
            status += f" [{code}]"
        if chantier:
            status += f" — {chantier}"
        print(f"  {status}")
    
    # ── Sauvegarder ───────────────────────────────────────────────
    # Créer un backup
    backup = str(excel_path).replace('.xlsx', '_BACKUP_avant_import.xlsx')
    import shutil
    shutil.copy2(excel_path, backup)
    print(f"\n  Backup créé: {Path(backup).name}")
    
    # Recalculer (LibreOffice si disponible)
    wb.save(excel_path)
    
    # Résumé
    print(f"\n{'─'*55}")
    print(f"  ✅ Import terminé")
    print(f"  Importés : {imported} jours")
    if skipped:
        print(f"  Ignorés  : {skipped} lignes")
    if errors:
        print(f"\n  ⚠️  Avertissements:")
        for e in errors[:10]:
            print(f"     • {e}")
    print(f"{'─'*55}")
    print(f"\n  Fichier mis à jour: {Path(excel_path).name}")
    print(f"  Pensez à recalculer l'Excel (Ctrl+Alt+F9 sous Excel)")
    print(f"  ou via: python recalc.py {excel_path}\n")
    
    return imported

# ── Point d'entrée ────────────────────────────────────────────────
if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(0)
    
    csv_path = sys.argv[1]
    
    # Chercher l'Excel automatiquement
    if len(sys.argv) >= 3:
        excel_path = sys.argv[2]
    else:
        candidates = list(Path('.').glob('MH_Gestion_Heures*.xlsx'))
        candidates = [c for c in candidates if 'BACKUP' not in str(c)]
        if candidates:
            excel_path = str(candidates[0])
            print(f"ℹ️  Excel trouvé automatiquement: {excel_path}")
        else:
            print("❌ Fichier Excel non trouvé. Spécifiez-le en argument:")
            print("   python import_csv.py mon_fichier.csv MH_Gestion_Heures_2026_PRO.xlsx")
            sys.exit(1)
    
    n = import_csv_to_excel(csv_path, excel_path)
    sys.exit(0 if n > 0 else 1)
