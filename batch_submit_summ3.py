# batch_submit_summ3.py

import os
import pandas as pd
import asyncio
import json
import smtplib
from email.message import EmailMessage
from datetime import datetime
from dotenv import load_dotenv
from siv_submitter import submit_form

# ── Load .env (EMAIL_USER, EMAIL_PASS, etc.) ───────────────────────────────────
load_dotenv()

EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")
EMAIL_TO   = os.getenv("EMAIL_TO", "catchfahad92@gmail.com")  # default receiver if not set

# ── Paths ──────────────────────────────────────────────────────────────────────
# Use relative INPUT_FILE so it works locally AND on Render
INPUT_FILE   = "entries.xlsx"
OUTPUT_DIR   = "results_final"
OUTPUT_FILE  = os.path.join(OUTPUT_DIR, "entries_results.xlsx")
ERRORS_FILE  = os.path.join(OUTPUT_DIR, "errors_only.xlsx")
SUMMARY_FILE = os.path.join(OUTPUT_DIR, "summary_report.json")

async def main():
    # Ensure output folder
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # Load input
    if not os.path.exists(INPUT_FILE):
        raise FileNotFoundError(f"Input Excel not found: {os.path.abspath(INPUT_FILE)}")
    df = pd.read_excel(INPUT_FILE)
    print(f"📄 Loaded {len(df)} entries from {INPUT_FILE}")

    # Normalize headers (French → internal)
    rename_map = {
        "Numéro d'immatriculation": "numero_immatriculation",
        "Date de première immatriculation du véhicule": "date_premiere_immat",
        "Date du certificat d'immatriculation": "date_certificat",
        "(Si personne physique) Nom et prénom": "nom_prenom",
        "ou (Si personne morale) Raison sociale": "raison_sociale",
        "Status": "result",
    }
    df.rename(columns=rename_map, inplace=True)

    # Clean NaN
    df = df.fillna("")

    # Dates → dd/mm/YYYY (explicit dayfirst to avoid warnings)
    for col in ["date_premiere_immat", "date_certificat"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=True).dt.strftime("%d/%m/%Y")
            df[col] = df[col].fillna("")

    # Ensure expected columns exist
    for col in ["numero_immatriculation", "date_premiere_immat", "date_certificat", "nom_prenom", "raison_sociale"]:
        if col not in df.columns:
            df[col] = ""

    processed_count = 0
    all_results = []

    # Iterate all rows
    for i, row in df.iterrows():
        immat = str(row.get("numero_immatriculation", "")).strip()
        if not immat:
            df.loc[i, "result"] = "skipped"
            print(f"⏭️  Skipping row {i+1}: empty immatriculation.")
            continue

        print(f"\n🔹 Processing {i+1}/{len(df)} → {immat}")

        data = {
            "numero_immatriculation": immat,
            "date_premiere_immat": str(row.get("date_premiere_immat", "")).strip(),
            "date_certificat": str(row.get("date_certificat", "")).strip(),
            "nom_prenom": str(row.get("nom_prenom", "")).strip(),
            "raison_sociale": str(row.get("raison_sociale", "")).strip(),
        }

        # Submit one entry
        try:
            result = await submit_form(data)
        except Exception as e:
            result = {"status": "technical_error", "message": str(e)}

        status = result.get("status", "unknown")
        df.loc[i, "result"] = status
        all_results.append({"immat": immat, "status": status})
        processed_count += 1
        print(f"➡️  {immat}: {status}")

        # Save progress incrementally
        df.to_excel(OUTPUT_FILE, index=False)

    # Build errors_only.xlsx
    error_df = df[df["result"].isin(["logical_error", "technical_error", "error"])]
    if not error_df.empty:
        error_df.to_excel(ERRORS_FILE, index=False)
        print(f"❌ Saved {len(error_df)} error rows → {ERRORS_FILE}")
    else:
        # Produce an empty file with same columns for consistency
        pd.DataFrame(columns=df.columns).to_excel(ERRORS_FILE, index=False)
        print("✅ No error rows — errors_only.xlsx generated as empty.")

    # JSON summary
    grouped = df.groupby("result")["numero_immatriculation"].apply(list).to_dict()
    summary = {
        "run_timestamp": datetime.now().isoformat(),
        "total_processed": int(processed_count),
        "ok": int((df["result"] == "ok").sum()),
        "logical_error": int((df["result"] == "logical_error").sum()),
        "technical_error": int((df["result"] == "technical_error").sum()),
        "error": int((df["result"] == "error").sum()),  # in case submit_form returns "error"
        "skipped": int((df["result"] == "skipped").sum()),
        "detailed_lists": grouped,
        "output_files": {
            "results_excel": OUTPUT_FILE,
            "errors_excel": ERRORS_FILE,
            "summary_json": SUMMARY_FILE,
        },
    }
    with open(SUMMARY_FILE, "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=4)

    print(f"\n🏁 All done!")
    print(f"📊 Results saved to: {OUTPUT_FILE}")
    print(f"❌ Errors saved to: {ERRORS_FILE}")
    print(f"🧾 JSON summary saved to: {SUMMARY_FILE}")

    # ── Optional: Email results if credentials exist ────────────────────────────
    if EMAIL_USER and EMAIL_PASS:
        try:
            send_summary_email(
                summary=summary,
                attachments=[OUTPUT_FILE, ERRORS_FILE, SUMMARY_FILE],
                sender=EMAIL_USER,
                password=EMAIL_PASS,
                recipient=EMAIL_TO,
            )
            print(f"📧 Email sent to {EMAIL_TO}")
        except Exception as e:
            print(f"⚠️ Email send failed: {e}")
    else:
        print("ℹ️ EMAIL_USER/EMAIL_PASS not set — skipping email.")

def send_summary_email(summary, attachments, sender, password, recipient):
    """Send an HTML summary email with attachments (results, errors, json)."""
    ok = summary.get("ok", 0)
    logical_error = summary.get("logical_error", 0)
    technical_error = summary.get("technical_error", 0)
    generic_error = summary.get("error", 0)
    skipped = summary.get("skipped", 0)
    total = summary.get("total_processed", 0)

    # HTML body (clean + readable)
    html = f"""
    <html>
      <body style="font-family: Arial, sans-serif; color:#222;">
        <h2>Rapport quotidien — SIV</h2>
        <p>Date: {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>

        <table cellpadding="8" cellspacing="0" border="0" style="border-collapse:collapse; margin-top:10px;">
          <tr>
            <td style="background:#f6f6f6; border:1px solid #ddd;">Total traités</td>
            <td style="border:1px solid #ddd;"><strong>{total}</strong></td>
          </tr>
          <tr>
            <td style="background:#f6f6f6; border:1px solid #ddd;">✅ OK</td>
            <td style="border:1px solid #ddd;"><strong>{ok}</strong></td>
          </tr>
          <tr>
            <td style="background:#f6f6f6; border:1px solid #ddd;">⚠️ Erreurs logiques</td>
            <td style="border:1px solid #ddd;"><strong>{logical_error}</strong></td>
          </tr>
          <tr>
            <td style="background:#f6f6f6; border:1px solid #ddd;">❌ Erreurs techniques</td>
            <td style="border:1px solid #ddd;"><strong>{technical_error + generic_error}</strong></td>
          </tr>
          <tr>
            <td style="background:#f6f6f6; border:1px solid #ddd;">⏭️ Ignorés</td>
            <td style="border:1px solid #ddd;"><strong>{skipped}</strong></td>
          </tr>
        </table>

        <p style="margin-top:15px;">Les fichiers en pièce jointe:</p>
        <ul>
          <li><code>entries_results.xlsx</code> (tous les résultats)</li>
          <li><code>errors_only.xlsx</code> (uniquement les erreurs)</li>
          <li><code>summary_report.json</code> (résumé JSON)</li>
        </ul>

        <p style="margin-top:10px;">Cordialement,<br/>SIV Automation</p>
      </body>
    </html>
    """

    msg = EmailMessage()
    msg["Subject"] = "Rapport quotidien — SIV"
    msg["From"] = sender
    msg["To"] = recipient
    msg.set_content("Veuillez consulter le rapport HTML. Si vous voyez ce message en texte, utilisez un client compatible HTML.")
    msg.add_alternative(html, subtype="html")

    # Attach files
    for path in attachments:
        if not path or not os.path.exists(path):
            continue
        with open(path, "rb") as f:
            data = f.read()
            filename = os.path.basename(path)
            # lazy generic MIME
            msg.add_attachment(data, maintype="application", subtype="octet-stream", filename=filename)

    import ssl
    import smtplib

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as smtp:
        smtp.login(sender, password)
        smtp.send_message(msg)

if __name__ == "__main__":
    asyncio.run(main())
