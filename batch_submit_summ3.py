import pandas as pd
import requests
from bs4 import BeautifulSoup
import io
import json
from datetime import datetime
import smtplib
from email.message import EmailMessage
import os

# -----------------------
# Config
# -----------------------
TARGET_URL = "https://siv.interieur.gouv.fr/map-usg-ui/do/csa_retour_dem_certificat"
SUCCESS_KEYWORDS = ["Récapitulatif", "Certificat", "Titulaire Principal"]
KNOWN_ERROR = "Aucun dossier ne correspond à la recherche. L'opération ne peut se poursuivre."

# Paths
INPUT_FILE = "entries.xlsx"
RESULT_FILE = "results.xlsx"
SUMMARY_FILE = "summary_report.json"
ERROR_FILE = "errors.xlsx"

# Email configuration
SENDER_EMAIL = "siv.automation.report@gmail.com"
APP_PASSWORD = "rjtpxqrrlsoitqpo"  # Gmail App Password
RECEIVER_EMAIL = "catchfahad92@gmail.com"

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
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Content-Type": "application/x-www-form-urlencoded",
})

ok_list, error_list, skipped_list = [], [], []
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
        else:
            results.append("error")
            error_list.append(immat)
    except Exception:
        results.append("error")
        error_list.append(immat)

df["result"] = results

# -----------------------
# 3) Build outputs
# -----------------------
ordered_cols = [c for c in original_columns if c in df.columns and c != "result"]
ordered_cols.append("result")
df_out = df[ordered_cols]

# Save all results
df_out.to_excel(RESULT_FILE, index=False)

# Extract only errors
error_df = df_out[df_out["result"] == "error"]
error_df.to_excel(ERROR_FILE, index=False)

# Build summary JSON
summary = {
    "run_timestamp": datetime.now().isoformat(),
    "total_processed": int(len(df)),
    "ok": int(len(ok_list)),
    "error": int(len(error_list)),
    "skipped": int(len(skipped_list)),
    "lists": {
        "ok": ok_list,
        "error": error_list,
        "skipped": skipped_list,
    }
}
with open(SUMMARY_FILE, "w", encoding="utf-8") as f:
    json.dump(summary, f, ensure_ascii=False, indent=2)

print("✅ Processing complete. Files generated:", RESULT_FILE, ERROR_FILE, SUMMARY_FILE)

# -----------------------
# 4) Email the results (HTML)
# -----------------------
msg = EmailMessage()
msg["Subject"] = f"🧾 Daily SIV Automation Report — {datetime.now().strftime('%d/%m/%Y')}"
msg["From"] = f"SIV Automation <{SENDER_EMAIL}>"
msg["To"] = RECEIVER_EMAIL

# HTML body
html_body = f"""
<html>
  <body style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#333;">
    <p>Bonjour,</p>
    <p>Veuillez trouver ci-joint les résultats du traitement automatique effectué le 
    <b>{datetime.now().strftime('%d/%m/%Y')}</b>.</p>

    <h3 style="color:#222;">Résumé des résultats :</h3>
    <table style="border-collapse:collapse;font-family:Arial,Helvetica,sans-serif;font-size:14px;">
      <tr style="background-color:#e7f7ec;">
        <td style="padding:8px 12px;">✅ Dossiers traités avec succès</td>
        <td style="padding:8px 12px;"><b style="color:#228B22;">{len(ok_list)}</b></td>
      </tr>
      <tr style="background-color:#fde8e8;">
        <td style="padding:8px 12px;">❌ Vérification du certificat échouée</td>
        <td style="padding:8px 12px;"><b style="color:#B22222;">{len(error_list)}</b></td>
      </tr>
      <tr style="background-color:#f4f4f4;">
        <td style="padding:8px 12px;">⏭️ Dossiers ignorés</td>
        <td style="padding:8px 12px;"><b>{len(skipped_list)}</b></td>
      </tr>
    </table>

    <p><br>Les fichiers <b>results.xlsx</b>, <b>errors.xlsx</b> et <b>summary_report.json</b> sont joints à ce message.</p>
    <p>Bien cordialement,<br><b>SIV Automation System</b></p>
  </body>
</html>
"""
msg.add_alternative(html_body, subtype="html")

# Attach all files
for file_path in [RESULT_FILE, ERROR_FILE, SUMMARY_FILE]:
    if os.path.exists(file_path):
        with open(file_path, "rb") as f:
            file_data = f.read()
            msg.add_attachment(
                file_data,
                maintype="application",
                subtype="octet-stream",
                filename=os.path.basename(file_path)
            )

# Send email
try:
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(SENDER_EMAIL, APP_PASSWORD)
        smtp.send_message(msg)
    print(f"✅ Email sent successfully to {RECEIVER_EMAIL}")
except Exception as e:
    print("❌ Failed to send email:", e)
