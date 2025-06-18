from unified_docx_extractor import UnifiedDocxExtractor
import json
import os

DOCX_PATH = r"C:\Users\HP\Desktop\StrideAi\ds-nn-unified\Test 10 BNSY (1) (1).docx"

OUTPUT_PATH = r"C:\Users\HP\Desktop\StrideAi\external_link_test\ds-nn-m6\scripts\unified_output4.json"

print(f"Running Unified Extractor on: {os.path.basename(DOCX_PATH)}...")
extractor = UnifiedDocxExtractor(DOCX_PATH)
results = extractor.run()

external_links_data = results.get('external_links', [])

with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
    json.dump(external_links_data, f, indent=4, ensure_ascii=False)

print(f"Successfully extracted {len(external_links_data)} external links.")
print(f"Output saved to: {OUTPUT_PATH}")
