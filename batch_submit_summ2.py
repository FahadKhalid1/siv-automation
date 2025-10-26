import pandas as pd
import requests
from bs4 import BeautifulSoup
import io
import json
from datetime import datetime
import smtplib
from email.message import EmailMessage
import os
from dotenv import load_dotenv  # ✅ NEW

# -----------------------
# Load environment variables
# -----------------------
load_dotenv()
SENDER_EMAIL = os.getenv("EMAIL_USER")
APP_PASSWORD = os.getenv("EMAIL_PASS")
RECEIVER_EMAIL = "catchfahad92@gmail.com"  # Final destination

# -----------------------
# Config
# -----------------------
TARGET_URL = "https://siv.interieur.gouv.fr/map-usg-ui/do/csa_retour_dem_certificat"
SUCCESS_KEYWORDS = ["Récapitulatif", "Certificat", "Titulaire Principal"]
KNOWN_LOGICAL_ERROR = "Aucun dossier ne correspond à la recherche. L'opération ne peut se poursuivre."

# Paths
INPUT_FILE = "entries_mini.xlsx"
RESULT_FILE = "results.xlsx"
SUMMARY_FILE = "summary_report.json"

# -----------------------
# 1) Load Excel
# -----------------------
df = pd.read_excel(INPUT_FILE)
original_columns = df.columns.tolist()

rename_map = {
    "Numéro d'immatriculation": "numero_immatriculation",
    "Date de première immatriculation du véhicule": "date_premiere_immat",
    "Date du certificat d'immatriculation": "date_certificat",
    "(Si personne physique) Nom et prénom": "nom_prenom",
    "ou (Si personne morale) Raison sociale": "raison_sociale",
    "Status": "result",
}
df.rename(columns=rename_map, inplace=True)
df = df.fillna("")

for col in ["date_premiere_immat", "date_certificat"]:
    if col in df.columns:
        df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=True).dt.strftime("%d/%m/%Y")
        df[col] = df[col].fillna("")

for col in ["numero_immatriculation", "date_premiere_immat", "date_certificat", "nom_prenom", "raison_sociale"]:
    if col not in df.columns:
        df[col] = ""

# -----------------------
# 2) Iterate and call SIV
# -----------------------
session = requests.Session()
session.headers.update({
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Content-Type": "application/x-www-form-urlencoded",
})

ok_list, logical_error_list, technical_error_list, skipped_list = [], [], [], []
results = []

try:
    session.get(TARGET_URL, timeout=30)
except Exception:
    pass

for i, row in df.iterrows():
    immat = str(row.get("numero_immatriculation", "")).strip()
    if not immat:
        results.append("skipped")
        skipped_list.append(immat)
        continue

    payload = {
        "rechercheDossier.numeroImmatriculation": immat,
        "rechercheDossier.datePremImmat": str(row.get("date_premiere_immat", "")).strip(),
        "rechercheDossier.dateCi": str(row.get("date_certificat", "")).strip(),
        "rechercheDossier.nomEtPrenom": str(row.get("nom_prenom", "")).strip(),
        "rechercheDossier.raisonSociale": str(row.get("raison_sociale", "")).strip(),
    }

    try:
        resp = session.post(TARGET_URL, data=payload, timeout=40, allow_redirects=True)
        html = resp.text
        soup = BeautifulSoup(html, "html.parser")
        text = soup.get_text(" ", strip=True)

        if any(k in html for k in SUCCESS_KEYWORDS):
            results.append("ok")
            ok_list.append(immat)
        elif KNOWN_LOGICAL_ERROR in text or "Aucun dossier ne correspond" in text:
            results.append("logical_error")
            logical_error_list.append(immat)
        else:
            results.append("technical_error")
            technical_error_list.append(immat)
    except Exception:
        results.append("technical_error")
        technical_error_list.append(immat)

df["result"] = results

# -----------------------
# 3) Build outputs
# -----------------------
ordered_cols = [c for c in original_columns if c in df.columns and c != "result"]
ordered_cols.append("result")
df_out = df[ordered_cols]

df_out.to_excel(RESULT_FILE, index=False)

summary = {
    "run_timestamp": datetime.now().isoformat(),
    "total_processed": int(len(df)),
    "ok": int(len(ok_list)),
    "logical_error": int(len(logical_error_list)),
    "technical_error": int(len(technical_error_list)),
    "skipped": int(len(skipped_list)),
    "lists": {
        "ok": ok_list,
        "logical_error": logical_error_list,
        "technical_error": technical_error_list,
        "skipped": skipped_list,
    }
}
with open(SUMMARY_FILE, "w", encoding="utf-8") as f:
    json.dump(summary, f, ensure_ascii=False, indent=2)

print("✅ Processing complete. Files generated:", RESULT_FILE, SUMMARY_FILE)

# -----------------------
# 4) Email the results
# -----------------------
msg = EmailMessage()
msg["Subject"] = f"🧾 Daily SIV Automation Report — {datetime.now().strftime('%d/%m/%Y')}"
msg["From"] = f"SIV Automation <{SENDER_EMAIL}>"
msg["To"] = RECEIVER_EMAIL
msg.set_content(
    f"""Bonjour,\n
Veuillez trouver ci-joint les résultats du traitement automatique du {datetime.now().strftime('%d/%m/%Y')}.\n
Résumé :
✅ OK: {len(ok_list)}
⚠️ Erreurs logiques: {len(logical_error_list)}
❌ Erreurs techniques: {len(technical_error_list)}
⏭️ Ignorés: {len(skipped_list)}\n
Bien cordialement,\n
SIV Automation System"""
)

# Attach both files
for file_path in [RESULT_FILE, SUMMARY_FILE]:
    with open(file_path, "rb") as f:
        file_data = f.read()
        msg.add_attachment(file_data, maintype="application", subtype="octet-stream", filename=os.path.basename(file_path))

try:
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(SENDER_EMAIL, APP_PASSWORD)
        smtp.send_message(msg)
    print(f"📧 Email sent successfully to {RECEIVER_EMAIL}")
except Exception as e:
    print("❌ Failed to send email:", e)
