import os
import re
import zipfile
import pythoncom
import win32com.client
import docx
import pandas as pd
import time
from lxml import etree
from pathlib import Path
 
class UnifiedDocxExtractor:
    def __init__(self, docx_path, abbr_reference_path=None):
        self.docx_path = docx_path
        self.abbr_reference_path = abbr_reference_path
        self.results = {}

    def run(self):
        # self._extract_headings_with_numbers()
        # self._extract_bullet_points()
        # self._extract_content_chunks()
        # self._extract_links_and_cross_references()
        # return self.results
        print("Starting extraction...")
        start_time = time.time()

        print("Extracting headings...")
        t1 = time.time()
        self._extract_headings_with_numbers()
        t2 = time.time()
        print(f"Headings extracted in {t2 - t1:.2f} seconds")

        print("Extracting bullet points...")
        t3 = time.time()
        self._extract_bullet_points()
        t4 = time.time()
        print(f"Bullets extracted in {t4 - t3:.2f} seconds")

        print("Extracting content chunks...")
        t5 = time.time()
        self._extract_content_chunks()
        t6 = time.time()
        print(f"Content chunks extracted in {t6 - t5:.2f} seconds")

        print("Extracting links and cross-references...")
        t7 = time.time()
        self._extract_links_and_cross_references()
        t8 = time.time()
        print(f"Links and cross-references extracted in {t8 - t7:.2f} seconds")

        total_time = time.time() - start_time
        print(f"Total extraction time: {total_time:.2f} seconds")
        print("Extraction complete.\n")
        return self.results

    def _clean_text(self, text):
        """Remove illegal characters and clean text"""
        if not text:
            return ""
        # Remove control characters, tabs, and page numbers
        cleaned = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F\uF000-\uFFFF]', '', str(text))
        # Remove tab characters and trailing page numbers
        cleaned = re.sub(r'\t\d+$', '', cleaned)  
        cleaned = re.sub(r'\t+', ' ', cleaned)    
        return cleaned.strip()
    

    def _extract_headings_with_numbers(self):
        """Extract headings with improved detection and number extraction"""
        pythoncom.CoInitialize()
        word = None
        doc = None
        headings = []
        
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(self.docx_path)
            word.ActiveDocument.Repaginate()
            
            # Method 1: Extract from Word's built-in heading styles
            for para in doc.Paragraphs:
                try:
                    style_name = para.Range.Style.NameLocal
                    text = self._clean_text(para.Range.Text)
                    
                    if not text or len(text) < 2:
                        continue
                    
                    # Check for heading styles
                    is_heading = ('Heading' in style_name and 'Table' not in style_name)
                    
                    # Also check for numbered paragraphs that look like headings
                    if not is_heading:
                        # Look for patterns like "1. Purpose", "5.1. Strategy", etc.
                        if re.match(r'^\d+(\.\d+)*\.\s+[A-Z]', text):
                            is_heading = True
                    
                    if is_heading:
                        heading_number = None
                        heading_name = text
                        
                        # Extract heading number and name
                        number_match = re.match(r'^(\d+(?:\.\d+)*)\.\s*(.+)', text)
                        if number_match:
                            heading_number = number_match.group(1)
                            heading_name = self._clean_text(number_match.group(2))
                            heading_name = re.sub(r'\s+\d+$', '', heading_name)

                        
                        # Get page number
                        page_number = para.Range.Information(3)  # wdActiveEndPageNumber
                        
                        # Get font size with error handling
                        try:
                            font_size = para.Range.Font.Size
                            if font_size and font_size > 100:  # Fix unrealistic font sizes
                                font_size = 12  # Default reasonable size
                        except:
                            font_size = 'Unknown'
                        
                        headings.append({
                            'Heading Number': heading_number,
                            'Heading Name': heading_name,
                            'Page No': page_number,
                            'Font Size': font_size,
                            'Style': style_name
                        })
                except Exception as e:
                    continue
            
            # Method 2: Fallback - scan for numbered headings in all paragraphs
            if len(headings) < 3:  # If we didn't find many headings, use fallback
                for para in doc.Paragraphs:
                    try:
                        text = self._clean_text(para.Range.Text)
                        if not text:
                            continue
                        
                        # Look for numbered headings
                        if re.match(r'^\d+(\.\d+)*\.\s+[A-Za-z]', text) and len(text) < 100:
                            number_match = re.match(r'^(\d+(?:\.\d+)*)\.\s*(.+)', text)
                            if number_match:
                                heading_number = number_match.group(1)
                                heading_name = self._clean_text(number_match.group(2))
                                heading_name = re.sub(r'\s+\d+$', '', heading_name)
                                page_number = para.Range.Information(3)
                                
                                try:
                                    font_size = para.Range.Font.Size
                                    if font_size and font_size > 100:
                                        font_size = 12
                                except:
                                    font_size = 'Unknown'
                                
                                # Check if already exists
                                exists = any(h['Heading Number'] == heading_number and 
                                           h['Heading Name'] == heading_name for h in headings)
                                
                                if not exists:
                                    headings.append({
                                        'Heading Number': heading_number,
                                        'Heading Name': heading_name,
                                        'Page No': page_number,
                                        'Font Size': font_size,
                                        'Style': 'Detected Heading'
                                    })
                    except Exception:
                        continue
            
            self.results['headings'] = headings
            
        except Exception as e:
            print(f"Error extracting headings: {str(e)}")
            self.results['headings'] = []
        finally:
            if doc is not None:
                try:
                    doc.Close(False)
                except Exception:
                    pass
            if word is not None:
                try:
                    word.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()

    # def _extract_bullet_points(self):
    #     """Extract actual bullet symbols/characters and their levels - UPDATED"""
    #     pythoncom.CoInitialize()
    #     word = None
    #     doc = None
    #     bullets = []
    #     try:
    #         word = win32com.client.Dispatch("Word.Application")
    #         word.Visible = False
    #         doc = word.Documents.Open(self.docx_path)
    #         word.ActiveDocument.Repaginate()
            
    #         for para in doc.Paragraphs:
    #             try:
    #                 text = self._clean_text(para.Range.Text)
    #                 if not text or len(text) < 3:
    #                     continue
                    
    #                 # Skip headings
    #                 if re.match(r'^\d+(\.\d+)*\.\s+', text):
    #                     continue
                    
    #                 # Extract only Word-recognized list items
    #                 if para.Range.ListFormat.ListType > 0:
    #                     bullet_symbol = para.Range.ListFormat.ListString.strip()
    #                     level = para.Range.ListFormat.ListLevelNumber + 1  # 1-based index
    #                     page_number = para.Range.Information(3)
                        
    #                     bullets.append({
    #                         'Bullet Point': bullet_symbol,
    #                         'Bullet Point Text': text,
    #                         'Level': level,
    #                         'Page No': page_number
    #                     })
                        
    #             except Exception as e:
    #                 continue
            
    #         self.results['bullets'] = bullets
            
    #     except Exception as e:
    #         print(f"Error extracting bullets: {str(e)}")
    #         self.results['bullets'] = []
    #     finally:
    #         if doc is not None:
    #             try:
    #                 doc.Close(False)
    #             except Exception:
    #                 pass
    #         if word is not None:
    #             try:
    #                 word.Quit()
    #             except Exception:
    #                 pass
    #         pythoncom.CoUninitialize()
    
    def _extract_bullet_points(self):
        """Extract and validate bullet points using managers logic."""
        # Bullet symbol mappings and expected levels
        BULLET_FONT_MAP = {
            61623: '•',
            9679: "•",
            8226: '•',  # Filled round bullet (•)
            8229: '▪',  # Filled square bullet (▪)
            10003: '✔',  # Check mark (✔)
            9675: '○',  # Hollow circle bullet (○)
            61607: '',  # Square bullet ()
            61692: '',
            61656: '',  # Checkmark ()
            111: 'o'
        }
        EXPECTED_BULLETS = {
            1: ['•'],
            2: ['○', 'o'],
            3: ['▪', ''],
        }

        pythoncom.CoInitialize()
        word = None
        doc = None
        bullet_points = []
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(self.docx_path)
            word.ActiveDocument.Repaginate()

            expected_bullet_chars = {symbol for lvl in EXPECTED_BULLETS for symbol in EXPECTED_BULLETS[lvl]}
            check_list = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "0"]

            for para in doc.Paragraphs:
                bullet_marker = None
                level = None

                # Native bullet list
                if para.Range.ListFormat.ListType in [2, 4]:
                    bullet_marker = para.Range.ListFormat.ListString.strip()
                    if any(item in bullet_marker for item in check_list):
                        bullet_marker = None
                    else:
                        level = para.Range.ListFormat.ListLevelNumber
                else:
                    # Manual bullet: check if first char is a known bullet
                    text = self._clean_text(para.Range.Text)
                    if text and text[0] in expected_bullet_chars:
                        bullet_marker = text[0]
                        if any(item in bullet_marker for item in check_list):
                            bullet_marker = None
                        else:
                            level = 1  # Manual bullets default to level 1

                if bullet_marker:
                    try:
                        ascii_value = ord(bullet_marker)
                    except Exception:
                        ascii_value = None

                    # Use mapping if available
                    if ascii_value is not None and ascii_value in BULLET_FONT_MAP:
                        bullet_symbol_ascii = BULLET_FONT_MAP[ascii_value]
                    else:
                        bullet_symbol_ascii = bullet_marker

                    text = self._clean_text(para.Range.Text)
                    page_num = para.Range.Information(3)

                    # Get the expected symbols for this bullet level
                    expected_symbols = EXPECTED_BULLETS.get(level, [])
                    if bullet_symbol_ascii not in expected_symbols:
                        bullet_points.append({
                            "error_type": "Incorrect bullet symbol",
                            "page": page_num,
                            "text": text,
                            "incorrect_symbol": f'{bullet_symbol_ascii} (ASCII: {ascii_value if ascii_value is not None else "Not available"})',
                            "level": level,
                            "expected_symbol": f'{expected_symbols} (ASCII: {[ord(symbol) for symbol in expected_symbols if symbol]})'
                        })

            self.results['bullets'] = bullet_points

        except Exception as e:
            print(f"Error extracting bullets: {str(e)}")
            self.results['bullets'] = []
        finally:
            if doc is not None:
                try:
                    doc.Close(False)
                except Exception:
                    pass
            if word is not None:
                try:
                    word.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()

    

    def _extract_content_chunks(self):
        """Extract content c    hunks with better filtering"""
        pythoncom.CoInitialize()
        word = None
        doc = None
        chunks = []
        
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(self.docx_path)
            word.ActiveDocument.Repaginate()
            
            bullet_symbols = '•●◦▪–-→➤o'
            
            for para in doc.Paragraphs:
                try:
                    text = self._clean_text(para.Range.Text)
                    if not text or len(text) < 10:
                        continue
                    
                    style_name = para.Range.Style.NameLocal
                    
                    # Skip headings
                    if ('Heading' in style_name or 
                        re.match(r'^\d+(\.\d+)*\.\s+', text)):
                        continue
                    
                    # Skip bullets (enhanced filtering)
                    if ('List' in style_name or 'Bullet' in style_name or 
                        (text and text[0] in bullet_symbols) or
                        (text.startswith('\t') or text.startswith('    ')) or
                        re.search(r'\[\d+\]$', text) or
                        (text.lower().startswith('server') or 
                         ('.' in text and text.count('.') >= 2 and len(text.split()) <= 3)) or
                        any(text.lower().startswith(word) for word in ['it systems', 'it security', 'data integrity', 'infrastructure', 'ownership'])):
                        continue
                    
                    # Skip TOC entries
                    if re.match(r'^\d+(\.\d+)*\.\s*\w+\s*\d+$', text):
                        continue
                    
                    # Skip very short content
                    if len(text.split()) < 3:
                        continue
                    
                    page_number = para.Range.Information(3)
                    
                    try:
                        font_size = para.Range.Font.Size
                        if font_size and font_size > 100:
                            font_size = 10
                    except:
                        font_size = 'Unknown'
                    
                    chunks.append({
                        'Content': text,
                        'Page No': page_number,
                        'Chunk': f"Paragraph on page {page_number}",
                        'Font Size': font_size
                    })
                except Exception:
                    continue
            
            self.results['content_chunks'] = chunks
            
        except Exception as e:
            print(f"Error extracting content chunks: {str(e)}")
            self.results['content_chunks'] = []
        finally:
            if doc is not None:
                try:
                    doc.Close(False)
                except Exception:
                    pass
            if word is not None:
                try:
                    word.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()

    # def _extract_links_and_cross_references(self):
    #     """Extract hyperlinks and cross-references with improved cleaning"""
    #     pythoncom.CoInitialize()
    #     word = None
    #     doc = None
    #     hyperlinks = []
    #     cross_references = []
        
    #     try:
    #         word = win32com.client.Dispatch("Word.Application")
    #         word.Visible = False
    #         doc = word.Documents.Open(self.docx_path)
    #         word.ActiveDocument.Repaginate()
            
    #         # Extract hyperlinks with better filtering
    #         seen_links = set()
    #         if doc.Hyperlinks.Count > 0:
    #             for hyperlink in doc.Hyperlinks:
    #                 try:
    #                     link_text = self._clean_text(hyperlink.TextToDisplay)
    #                     address = hyperlink.Address if hyperlink.Address else f"Internal: {hyperlink.SubAddress}"
    #                     page_number = hyperlink.Range.Information(3)
                        
    #                     # Better filtering for meaningful links
    #                     if (not link_text or len(link_text) < 1 or 
    #                         link_text.isspace() or link_text in ['\t', ' ']):
    #                         continue
                        
    #                     # Create unique key
    #                     link_key = (link_text, address, page_number)
    #                     if link_key not in seen_links:
    #                         seen_links.add(link_key)
    #                         hyperlinks.append({
    #                             "type": "Hyperlink",
    #                             "page_number": page_number,
    #                             "link_text": link_text,
    #                             "target": self._clean_text(address)
    #                         })
    #                 except Exception:
    #                     continue
            
    #         # Extract cross-references
    #         for field in doc.Fields:
    #             try:
    #                 if field.Type == 3:  # wdFieldRef
    #                     ref_text = self._clean_text(field.Result.Text)
    #                     target_text = self._clean_text(field.Code.Text) if field.Code else "No target text"
    #                     page_number = field.Code.Information(3) if field.Code else 1
                        
    #                     if ref_text and len(ref_text) > 0:
    #                         cross_references.append({
    #                             "type": "Cross-reference",
    #                             "page_number": page_number,
    #                             "ref_text": ref_text,
    #                             "target_text": target_text
    #                         })
    #             except Exception:
    #                 continue
            
    #         # Extract [n] style references
    #         seen_refs = set()
    #         for para in doc.Paragraphs:
    #             try:
    #                 text = para.Range.Text
    #                 if text:
    #                     ref_matches = re.findall(r'\[(\d+)\]', text)
    #                     if ref_matches:
    #                         page_number = para.Range.Information(3)
    #                         for ref_num in ref_matches:
    #                             ref_key = (ref_num, page_number)
    #                             if ref_key not in seen_refs:
    #                                 seen_refs.add(ref_key)
    #                                 cross_references.append({
    #                                     "type": "Cross-reference",
    #                                     "page_number": page_number,
    #                                     "ref_text": f"[{ref_num}]",
    #                                     "target_text": f"Reference {ref_num}"
    #                                 })
    #             except Exception:
    #                 continue
            
    #         self.results['hyperlinks'] = hyperlinks
    #         self.results['cross_references'] = cross_references
            
    #     except Exception as e:
    #         print(f"Error extracting links and references: {str(e)}")
    #         self.results['hyperlinks'] = []
    #         self.results['cross_references'] = []
    #     finally:
    #         if doc is not None:
    #             try:
    #                 doc.Close(False)
    #             except Exception:
    #                 pass
    #         if word is not None:
    #             try:
    #                 word.Quit()
    #             except Exception:
    #                 pass
    #         pythoncom.CoUninitialize()
    def _extract_links_and_cross_references(self):
        pythoncom.CoInitialize()
        word = None
        doc = None
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(self.docx_path)
            word.ActiveDocument.Repaginate()

            self.results['external_links'] = []
            self.results['internal_links'] = []
            self.results['cross_references'] = []

            #for hyperlinks
            for hyperlink in doc.Hyperlinks:
                link_text = hyperlink.TextToDisplay
                # If Address is present, it's an external link; otherwise, it's internal (bookmark/anchor)
                if hyperlink.Address:
                    address = hyperlink.Address
                else:
                    address = f"Internal: {hyperlink.SubAddress}" if hyperlink.SubAddress else ""
                try:
                    page_number = hyperlink.Range.Information(3)
                except Exception:
                    page_number = None

                # External link: has an Address and is not a Word internal anchor
                if hyperlink.Address and hyperlink.Address.lower().startswith(('http', 'https', 'www')):
                    self.results['external_links'].append({
                        "type": "Hyperlink",
                        "page_number": page_number,
                        "link_text": link_text,
                        "target": address
                    })
                # Internal link: has SubAddress (bookmark/anchor)
                elif hyperlink.SubAddress:
                    self.results['internal_links'].append({
                        "type": "Hyperlink",
                        "page_number": page_number,
                        "link_text": link_text,
                        "target": address
                    })

            # for cross-references
            for field in doc.Fields:
                if field.Type == 3:  # wdFieldRef (Cross-reference)
                    ref_text = field.Result.Text
                    target_text = field.Code.Text if field.Code else "No target text"
                    try:
                        page_number = field.Code.Information(3)
                    except Exception:
                        page_number = None
                    self.results['cross_references'].append({
                        "type": "Cross-reference",
                        "page_number": page_number,
                        "ref_text": ref_text,
                        "target_text": target_text
                    })

        except Exception as e:
            print(f"Error extracting links: {str(e)}")
            self.results['external_links'] = []
            self.results['internal_links'] = []
            self.results['cross_references'] = []
        finally:
            if doc is not None:
                try:
                    doc.Close(False)
                except Exception:
                    pass
            if word is not None:
                try:
                    word.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()

# Example usage
# if __name__ == "__main__":
#     extractor = UnifiedDocxExtractor(r"C:\Users\HP\Desktop\StrideAi\ds-nn-unified\Test5_PKBW.docx")
#     results = extractor.run()
#     import pprint
#     pprint.pprint(results)