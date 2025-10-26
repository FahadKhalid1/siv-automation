import pandas as pd
import asyncio
from siv_submitter import submit_form

INPUT_FILE = "entries.xlsx"
OUTPUT_FILE = "entries_results.xlsx"

async def main():
    # Load all entries
    df = pd.read_excel(INPUT_FILE)
    print(f"📄 Loaded {len(df)} entries from {INPUT_FILE}")

    # Normalize column names (handles French Excel headers)
    rename_map = {
        "Numéro d'immatriculation": "numero_immatriculation",
        "Date de première immatriculation du véhicule": "date_premiere_immat",
        "Date du certificat d'immatriculation": "date_certificat",
        "(Si personne physique) Nom et prénom": "nom_prenom",
        "ou (Si personne morale) Raison sociale": "raison_sociale",
        "Status": "result"
    }
    df.rename(columns=rename_map, inplace=True)

    # Replace NaN with empty strings
    df = df.fillna("")

    # Format date columns to dd/mm/yyyy
    for col in ["date_premiere_immat", "date_certificat"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce").dt.strftime("%d/%m/%Y")
            df[col] = df[col].fillna("")

    results = []

    for i, row in df.iterrows():
        print(f"\n🔹 Processing entry {i+1}/{len(df)} ({row['numero_immatriculation']})")

        data = {
            "numero_immatriculation": str(row.get("numero_immatriculation", "")).strip(),
            "date_premiere_immat": str(row.get("date_premiere_immat", "")).strip(),
            "date_certificat": str(row.get("date_certificat", "")).strip(),
            "nom_prenom": str(row.get("nom_prenom", "")).strip(),
            "raison_sociale": str(row.get("raison_sociale", "")).strip(),
        }

        # Skip empty plate numbers
        if not data["numero_immatriculation"]:
            print("⚠️  Skipping empty row.")
            df.loc[i, "result"] = "skipped"
            continue

        try:
            result = await submit_form(data)
        except Exception as e:
            result = {"status": "error", "message": str(e)}

        results.append(result.get("status", "unknown"))

        # Save progress incrementally
        df.loc[i, "result"] = result.get("status", "unknown")
        df.to_excel(OUTPUT_FILE, index=False)

        print(f"✅ Result for {data['numero_immatriculation']}: {result.get('status', 'unknown')}")

    print(f"\n🏁 All done! Results saved to {OUTPUT_FILE}")

if __name__ == "__main__":
    asyncio.run(main())
