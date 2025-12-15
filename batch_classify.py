import os
import csv
import json
from typing import List, Dict

from extract import extract_setup_data_to_json
from classifier import BusinessClassifier


def list_excel_files(directory: str) -> List[str]:
    allowed_exts = (".xlsm", ".xlsx")
    files: List[str] = []
    for name in sorted(os.listdir(directory)):
        if name.startswith('.'):
            continue
        if name == 'Sub-Sectors_vf.xlsx':
            continue
        if not name.lower().endswith(allowed_exts):
            continue
        files.append(os.path.join(directory, name))
    return files


def get_company_name_from_extracted(data: Dict) -> str:
    return (
        data.get('i_Setup', {})
            .get('fields', {})
            .get('Business Name', {})
            .get('value', '')
            or ''
    ).strip()


def main() -> None:
    script_dir = os.path.dirname(os.path.abspath(__file__))
    excel_dir = os.path.join(script_dir, 'financial_analysis')  # Excel files live in the project root alongside this script

    # Initialize classifier once
    openai_api_key = os.getenv("OPENAI_API_KEY")
    reference_file_path = os.path.join(script_dir, 'categories ideas.xlsx')
    classifier = BusinessClassifier(openai_api_key, reference_file_path)

    excel_files = list_excel_files(excel_dir) # cap to first 20 files for testing
    if not excel_files:
        print("No Excel files found to process.")
        return

    output_csv = os.path.join(script_dir, 'batch_classification_results.csv')
    fieldnames = [
        'source_file',
        'company_name',
        'sector',
        'primary_subsector',
        'additional_subsectors',
        'top_products',
        'confidence_explanation',
    ]

    rows: List[Dict[str, str]] = []

    for excel_path in excel_files:
        try:
            print(f"Processing: {excel_path}")

            # Extract JSON structure (avoid writing JSON file by passing a dummy path if needed)
            extracted = extract_setup_data_to_json(excel_path, output_json_path=None)
            if not extracted:
                print(f"  Skipped (no data extracted)")
                continue

            # Classify
            result = classifier.classify_business(extracted)

            # Collect row
            company_name = get_company_name_from_extracted(extracted) or os.path.splitext(os.path.basename(excel_path))[0]
            additional = "; ".join(result.additional_subsectors) if result.additional_subsectors else ''
            top_products = "; ".join(result.top_products[:5]) if getattr(result, "top_products", None) else ''

            rows.append({
                'source_file': os.path.basename(excel_path),
                'company_name': company_name,
                'sector': result.sector,
                'primary_subsector': result.primary_subsector,
                'additional_subsectors': additional,
                'top_products': top_products,
                'confidence_explanation': result.confidence_explanation,
            })

        except Exception as e:
            print(f"  Error processing {excel_path}: {e}")
            continue

    # Write/append CSV
    if rows:
        file_exists = os.path.exists(output_csv)
        with open(output_csv, 'a', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            if not file_exists:
                writer.writeheader()
            writer.writerows(rows)
        action = "Appended" if file_exists else "Wrote"
        print(f"\n{action} {len(rows)} rows to {output_csv}")
    else:
        print("No classification results to write.")


if __name__ == '__main__':
    main()


