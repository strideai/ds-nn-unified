import json
import os

INPUT_PATH = r"C:\Users\HP\Desktop\StrideAi\ds-nn-unified\Bulletissue.json"

OUTPUT_PATH = r"C:\Users\HP\Desktop\StrideAi\external_link_test\ds-nn-bulletpoint\scripts\unified_bullet_points.json"

print(f"--- Filtering unified output for ds-nn-bulletpoint ---")
print(f"Reading master data from: {os.path.basename(INPUT_PATH)}")

try:
    with open(INPUT_PATH, 'r', encoding='utf-8') as f:
        data = json.load(f)

    bullets_data = data.get('bullets', [])
 
    with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
        json.dump(bullets_data, f, indent=4, ensure_ascii=False)

    print(f"\n  Saved {len(bullets_data)} bullet points to:")
    print(f"{OUTPUT_PATH}")

except FileNotFoundError:
    print(f"\n ERROR: Input file not found at the specified INPUT_PATH.")
    print(f"Please make sure this path is correct: '{INPUT_PATH}'")
except json.JSONDecodeError:
    print(f"\n ERROR: Could not read the input JSON file. It may be empty or corrupt.")
except Exception as e:
    print(f"\n An unexpected error occurred: {e}")
