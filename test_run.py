# from unified_docx_extractor import UnifiedDocxExtractor
# import pprint

# extractor = UnifiedDocxExtractor(r"C:\Users\HP\Desktop\Data\ds-nn-unified\Bullet points issue.docx")
# results = extractor.run()
# pprint.pprint(results)

# import json
# import pandas as pd

# # # Save as JSON
# # with open('extraction_results.json', 'w', encoding='utf-8') as f:
# #     json.dump(results, f, indent=2, ensure_ascii=False)

# # Save as Excel
# # import pandas as pd

# # with pd.ExcelWriter('extraction_report.xlsx') as writer:
# #     # Headings
# #     pd.DataFrame(results['headings'], columns=['Heading Text', 'Style']).to_excel(writer, sheet_name='Headings', index=False)
# #     # Bullets
# #     pd.DataFrame(results['bullets'], columns=['Bullet Item']).to_excel(writer, sheet_name='Bullets', index=False)
# #     # Abbreviations
# #     pd.DataFrame(results['abbreviations'], columns=['Abbreviation']).to_excel(writer, sheet_name='Abbreviations', index=False)
# #     # Font Issues
# #     pd.DataFrame(results['font_issues']).to_excel(writer, sheet_name='Font Issues', index=False)
# #     # Cross References
# #     pd.DataFrame(results['cross_references'], columns=['Reference']).to_excel(writer, sheet_name='Cross References', index=False)
# #     # Tables
# #     for idx, table in enumerate(results['tables']):
# #         pd.DataFrame(table).to_excel(writer, sheet_name=f'Table_{idx+1}', index=False)
    

# # Save as Markdown

# def write_markdown_report(results, filename='extraction_report.md'):
#     with open(filename, 'w', encoding='utf-8') as f:
#         f.write("# Document Extraction Report\n\n")
#         f.write("## Headings\n")
#         for h, style in results.get('headings', []):
#             f.write(f"- **{h}** ({style})\n")
#         f.write("\n## Bullets\n")
#         for b in results.get('bullets', []):
#             f.write(f"- {b}\n")
#         f.write("\n## Abbreviations\n")
#         for abbr in results.get('abbreviations', []):
#             f.write(f"- {abbr}\n")
#         f.write("\n## Font Issues\n")
#         for issue in results.get('font_issues', []):
#             f.write(f"- Page {issue['page']}: \"{issue['text']}\" (Current: {issue['current_size']}, Expected: {issue['expected_size']}, Style: {issue['style']})\n")
#         f.write("\n## Cross References\n")
#         for ref in results.get('cross_references', []):
#             f.write(f"- {ref}\n")
#         f.write("\n## Tables\n")
#         for idx, table in enumerate(results.get('tables', [])):
#             f.write(f"\n### Table {idx+1}\n")
#             for row in table:
#                 f.write(" | ".join(row) + "\n")
        

# write_markdown_report(results, 'extraction_report.md')

from unified_docx_extractor import UnifiedDocxExtractor
import json
import pandas as pd
import re

# Function to clean illegal characters for Excel
def clean_for_excel(data):
    """Clean data for Excel export by removing illegal characters"""
    if isinstance(data, str):
        # Remove control characters that Excel can't handle
        return re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F\uF000-\uFFFF]', '', data)
    elif isinstance(data, dict):
        return {k: clean_for_excel(v) for k, v in data.items()}
    elif isinstance(data, list):
        return [clean_for_excel(item) for item in data]
    else:
        return data

extractor = UnifiedDocxExtractor(r"C:\Users\HP\Desktop\StrideAi\ds-nn-unified\Test5_PKBW.docx")
results = extractor.run()

clean_results = clean_for_excel(results)

# import pprint
# pprint.pprint(clean_results)

with open('test5.json', 'w', encoding='utf-8') as f:
    json.dump(clean_results, f, indent=2, ensure_ascii=False)

# try:
#     with pd.ExcelWriter('Bulletissue.xlsx', engine='openpyxl') as writer:
#         # Headings
#         if clean_results['headings']:
#             pd.DataFrame(clean_results['headings']).to_excel(writer, sheet_name='Headings', index=False)
        
#         # Bullets
#         if clean_results['bullets']:
#             pd.DataFrame(clean_results['bullets']).to_excel(writer, sheet_name='Bullets', index=False)
        
#         # Content Chunks
#         if clean_results['content_chunks']:
#             pd.DataFrame(clean_results['content_chunks']).to_excel(writer, sheet_name='Content Chunks', index=False)
        
#         # Hyperlinks
#         if clean_results['hyperlinks']:
#             pd.DataFrame(clean_results['hyperlinks']).to_excel(writer, sheet_name='Hyperlinks', index=False)
        
#         # Cross References
#         if clean_results['cross_references']:
#             pd.DataFrame(clean_results['cross_references']).to_excel(writer, sheet_name='Cross References', index=False)
    
#     print(" Excel file created successfully!")
    
# except Exception as e:
#     print(f" Excel export error: {e}")

print(" Final extraction complete! Check 'test5.json' and 'test5.xlsx'")
