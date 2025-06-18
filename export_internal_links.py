import json
import os
from unified_docx_extractor import UnifiedDocxExtractor

DOCX_PATH = r"C:\Users\HP\Desktop\StrideAi\ds-nn-unified\Bullet points issue.docx"

OUTPUT_PATH = r"C:\Users\HP\Desktop\StrideAi\external_link_test\ds-nn-m7\scripts\unified_internal_links1.json"


print(f"--- Running Unified Extractor for ds-nn-m7 ---")
print(f"Analyzing document: {os.path.basename(DOCX_PATH)}")

try:
 
    extractor = UnifiedDocxExtractor(DOCX_PATH)
    results = extractor.run()

    internal_links = results.get('internal_links', [])
    cross_references = results.get('cross_references', [])

    combined_list = internal_links + cross_references

    with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
        json.dump(combined_list, f, indent=4, ensure_ascii=False)

    print(f"\nSuccess! Extracted and saved {len(combined_list)} total internal links and cross-references to:")
    print(f"{OUTPUT_PATH}")

except FileNotFoundError:
    print(f"\nERROR: Input DOCX file not found.")
    print(f"Please make sure this path is correct: '{DOCX_PATH}'")
except Exception as e:
    print(f"\nAn unexpected error occurred during extraction: {e}")
