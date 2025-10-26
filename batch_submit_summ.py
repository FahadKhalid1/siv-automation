import os
import pandas as pd
import asyncio
import json
from siv_submitter import submit_form

INPUT_FILE = "/Users/fahad/Desktop/car_stolen/entries.xlsx"
OUTPUT_DIR = "results_final"  # new output folder for clarity
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "entries_results.xlsx")
ERRORS_FILE = os.path.join(OUTPUT_DIR, "errors_only.xlsx")
SUMMARY_FILE = os.path.join(OUTPUT_DIR, "summary_report.json")


async def main():
    # === Ensure output directory exists ===
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # === Load and prepare input ===
    df = pd.read_excel(INPUT_FILE)
    print(f"📄 Loaded {len(df)} entries from {INPUT_FILE}")

    # Normalize columns (supports French headers)
    rename_map = {
        "Numéro d'immatriculation": "numero_immatriculation",
        "Date de première immatriculation du véhicule": "date_premiere_immat",
        "Date du certificat d'immatriculation": "date_certificat",
        "(Si personne physique) Nom et prénom": "nom_prenom",
        "ou (Si personne morale) Raison sociale": "raison_sociale",
        "Status": "result"
    }
    df.rename(columns=rename_map, inplace=True)

    # Clean NaN and normalize date format
    df = df.fillna("")
    for col in ["date_premiere_immat", "date_certificat"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce").dt.strftime("%d/%m/%Y")
            df[col] = df[col].fillna("")

    # Prepare result storage
    all_results = []
    processed_count = 0

    for i, row in df.iterrows():
        immat = str(row.get("numero_immatriculation", "")).strip()
        if not immat:
            df.loc[i, "result"] = "skipped"
            continue

        print(f"\n🔹 Processing entry {i+1}/{len(df)} ({immat})")

        data = {
            "numero_immatriculation": immat,
            "date_premiere_immat": str(row.get("date_premiere_immat", "")).strip(),
            "date_certificat": str(row.get("date_certificat", "")).strip(),
            "nom_prenom": str(row.get("nom_prenom", "")).strip(),
            "raison_sociale": str(row.get("raison_sociale", "")).strip(),
        }

        try:
            result = await submit_form(data)
        except Exception as e:
            result = {"status": "technical_error", "message": str(e)}

        status = result.get("status", "unknown")
        df.loc[i, "result"] = status
        all_results.append({"immat": immat, "status": status})
        processed_count += 1

        print(f"➡️  {immat}: {status}")

        # Incremental save
        df.to_excel(OUTPUT_FILE, index=False)

    # === Build error-only dataset ===
    error_df = df[df["result"].isin(["logical_error", "technical_error"])]
    if not error_df.empty:
        error_df.to_excel(ERRORS_FILE, index=False)
        print(f"❌ Saved {len(error_df)} error rows → {ERRORS_FILE}")
    else:
        print("✅ No error rows found — errors file will be empty.")
        pd.DataFrame(columns=df.columns).to_excel(ERRORS_FILE, index=False)

    # === Build JSON summary with detailed lists ===
    grouped = df.groupby("result")["numero_immatriculation"].apply(list).to_dict()
    summary = {
        "total_processed": int(processed_count),
        "ok": int((df["result"] == "ok").sum()),
        "logical_error": int((df["result"] == "logical_error").sum()),
        "technical_error": int((df["result"] == "technical_error").sum()),
        "skipped": int((df["result"] == "skipped").sum()),
        "detailed_lists": grouped,
        "output_files": {
            "results_excel": OUTPUT_FILE,
            "errors_excel": ERRORS_FILE,
            "summary_json": SUMMARY_FILE
        }
    }

    with open(SUMMARY_FILE, "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=4)

    print(f"\n🏁 All done!")
    print(f"📊 Results saved to: {OUTPUT_FILE}")
    print(f"❌ Errors saved to: {ERRORS_FILE}")
    print(f"🧾 JSON summary saved to: {SUMMARY_FILE}")


if __name__ == "__main__":
    asyncio.run(main())
