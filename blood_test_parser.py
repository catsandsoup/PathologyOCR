import pytesseract
from PIL import Image
import pandas as pd
import re
from typing import Dict, List, Tuple

class TemplateBloodTestParser:
    def __init__(self, tesseract_path: str = r"C:\Program Files\Tesseract-OCR\tesseract.exe"):
        pytesseract.pytesseract.tesseract_cmd = tesseract_path
        
        # Test name mapping - maps OCR text variations to standard test names
        self.test_name_mapping = {
            # Exact matches
            'sodium': 'Sodium',
            'potassium': 'Potassium', 
            'chloride': 'Chloride',
            'bicarbonate': 'Bicarbonate',
            'urea': 'Urea',
            'creatinine': 'Creatinine',
            'egfr': 'eGFR',
            'calcium': 'Calcium',
            'corrected calcium': 'Calcium',  # Map corrected calcium to calcium
            'corr calcium': 'Calcium',
            'magnesium': 'Magnesium',
            'phosphate': 'Phosphate',
            'bilirubin total': 'Bili.Total',
            'bili.total': 'Bili.Total',
            'bili total': 'Bili.Total',
            'alp': 'ALP',
            'ggt': 'GGT', 
            'ld': 'LD',
            'ast': 'AST',
            'alt': 'ALT',
            'total protein': 'Total Protein',
            'totalprotein': 'Total Protein',
            'albumin': 'Albumin',
            'globulin': 'Globulin',
            'cholesterol': 'Cholesterol',
            'triglycerides': 'Triglycerides'
        }
        
    def extract_text_from_image(self, image_path: str) -> str:
        """Extract text using OCR with optimized settings"""
        image = Image.open(image_path)
        
        # Enhanced OCR config for better table recognition
        custom_config = r'--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz./()\-<>: '
        raw_text = pytesseract.image_to_string(image, config=custom_config)
        
        return raw_text
        
    def parse_dates_from_text(self, text: str) -> List[str]:
        """Extract test dates from the OCR text"""
        # Look for date patterns
        date_patterns = [
            r'\b\d{2}/\d{2}/\d{2}\b',  # DD/MM/YY
            r'\b\d{2}/\d{2}/\d{4}\b',  # DD/MM/YYYY
            r'\b\d{1,2}/\d{1,2}/\d{2,4}\b'  # More flexible
        ]
        
        all_dates = []
        for pattern in date_patterns:
            dates = re.findall(pattern, text)
            all_dates.extend(dates)
        
        # Remove duplicates while preserving order
        unique_dates = []
        for date in all_dates:
            if date not in unique_dates:
                unique_dates.append(date)
                
        return unique_dates[:6]  # Max 6 date columns as per template
    
    def normalize_test_name(self, raw_name: str) -> str:
        """Convert OCR test name to standard template name"""
        if not raw_name:
            return None
            
        # Clean and normalize
        cleaned = raw_name.lower().strip()
        cleaned = re.sub(r'[^\w\s]', '', cleaned)  # Remove punctuation
        cleaned = re.sub(r'\s+', ' ', cleaned)     # Normalize spaces
        
        # Try exact match first
        if cleaned in self.test_name_mapping:
            return self.test_name_mapping[cleaned]
            
        # Try partial matches
        for key, value in self.test_name_mapping.items():
            if key in cleaned or cleaned in key:
                return value
                
        return None
    
    def extract_test_results(self, text: str, dates: List[str]) -> Dict[str, Dict]:
        """Extract test results from OCR text"""
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        
        results = {}  # {test_name: {date1: value1, date2: value2, ...}}
        
        # Find the data section (after header with dates)
        data_start_idx = 0
        for i, line in enumerate(lines):
            # Look for a line that contains multiple dates
            date_count = sum(1 for date in dates if date in line)
            if date_count >= 2:
                data_start_idx = i + 3  # Skip header lines
                break
        
        i = data_start_idx
        while i < len(lines):
            line = lines[i].strip()
            
            # Stop conditions
            if (line.lower().startswith('comment') or 
                line.lower().startswith('end') or
                'www.' in line.lower()):
                break
                
            # Try to parse this line as a test result
            test_data = self.parse_result_line(line, dates)
            
            if test_data:
                test_name = test_data['name']
                if test_name:
                    results[test_name] = test_data['values']
            
            i += 1
                
        return results
    
    def parse_result_line(self, line: str, dates: List[str]) -> Dict:
        """Parse a single result line"""
        if not line or len(line) < 5:
            return None
            
        # Split line into parts
        parts = line.split()
        if len(parts) < 2:
            return None
            
        # Find where numeric values start
        value_start_idx = None
        for i, part in enumerate(parts):
            # Look for numeric patterns (including H/L prefixes)
            if re.match(r'^[HL]?\d+\.?\d*$|^[<>]\d+\.?\d*$', part):
                value_start_idx = i
                break
        
        if value_start_idx is None:
            return None
            
        # Extract test name (everything before values)
        test_name_parts = parts[:value_start_idx]
        test_name_raw = ' '.join(test_name_parts)
        test_name = self.normalize_test_name(test_name_raw)
        
        if not test_name:
            return None
            
        # Extract values
        value_parts = parts[value_start_idx:]
        values = {}
        
        # We expect up to len(dates) values
        for i, part in enumerate(value_parts):
            if i >= len(dates):
                break
                
            # Check if this looks like a test value
            if re.match(r'^[HL]?\d+\.?\d*$|^[<>]\d+\.?\d*$', part):
                if i < len(dates):
                    values[dates[i]] = part
            elif part in ['Unkn', 'Unknown', '']:
                # Handle unknown values
                if i < len(dates):
                    values[dates[i]] = ''
                    
        return {
            'name': test_name,
            'values': values
        }
    
    def load_template(self, template_path: str) -> pd.DataFrame:
        """Load the Excel template and return as DataFrame"""
        df = pd.read_excel(template_path, sheet_name='Blood Tests')
        return df
    
    def populate_template_with_results(self, template_df: pd.DataFrame, 
                                     results: Dict[str, Dict], 
                                     dates: List[str]) -> pd.DataFrame:
        """Populate the template with extracted results"""
        # Make a copy of the template
        df = template_df.copy()
        
        # Update the date headers (columns C onwards)
        date_cols = df.columns[2:2+len(dates)]  # Skip Test and Unit columns
        for i, date in enumerate(dates):
            if i < len(date_cols):
                df.columns.values[2+i] = date
        
        # Populate results
        for test_name, test_values in results.items():
            # Find the row for this test
            test_row_idx = df[df['Test'].str.contains(test_name, case=False, na=False)].index
            
            if len(test_row_idx) == 0:
                # Test not found in template, add it
                print(f"Warning: Test '{test_name}' not found in template")
                continue
                
            row_idx = test_row_idx[0]
            
            # Populate the values for each date
            for date, value in test_values.items():
                if date in dates:
                    date_col_idx = 2 + dates.index(date)
                    if date_col_idx < len(df.columns):
                        df.iloc[row_idx, date_col_idx] = value
        
        return df
    
    def process_blood_test(self, image_path: str, template_path: str, output_path: str = "blood_test_results.csv"):
        """Main processing function"""
        print("Starting blood test processing...")
        
        # Step 1: Extract text from image
        print("1. Extracting text from image...")
        ocr_text = self.extract_text_from_image(image_path)
        
        # Debug: Save OCR text for inspection
        with open("debug_ocr_output.txt", "w", encoding="utf-8") as f:
            f.write(ocr_text)
        print("   OCR text saved to debug_ocr_output.txt")
        
        # Step 2: Extract dates
        print("2. Extracting test dates...")
        dates = self.parse_dates_from_text(ocr_text)
        print(f"   Found dates: {dates}")
        
        if not dates:
            print("   Warning: No dates found in OCR text!")
            return None
            
        # Step 3: Extract test results
        print("3. Extracting test results...")
        results = self.extract_test_results(ocr_text, dates)
        print(f"   Extracted {len(results)} tests")
        
        # Debug: Print extracted results
        for test, values in results.items():
            print(f"   {test}: {values}")
        
        # Step 4: Load template
        print("4. Loading template...")
        template_df = self.load_template(template_path)
        print(f"   Template loaded with {len(template_df)} test rows")
        
        # Step 5: Populate template
        print("5. Populating template with results...")
        final_df = self.populate_template_with_results(template_df, results, dates)
        
        # Step 6: Save results
        print(f"6. Saving results to {output_path}...")
        final_df.to_csv(output_path, index=False)
        
        print(f"✅ Processing complete!")
        print(f"   CSV file: {output_path}")
        
        return final_df

# Usage
if __name__ == "__main__":
    parser = TemplateBloodTestParser()
    
    # Process the blood test
    try:
        result_df = parser.process_blood_test(
            image_path="your_blood_test_image.jpg",
            template_path="PathologyPro_Template.xlsx",
            output_path="populated_blood_test_results.csv"
        )
        
        if result_df is not None:
            print("\n📊 Sample of populated data:")
            print(result_df.head(10))
        
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()