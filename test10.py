import pandas as pd
import numpy as np
from decimal import Decimal, getcontext
from datetime import datetime
import os
from pathlib import Path
import warnings

warnings.filterwarnings('ignore')
getcontext().prec = 28


class FinancialAnalyzer:
    def __init__(self, input_folder_path):
        self.input_folder = Path(input_folder_path)
        self.current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.output_dir = self.input_folder / "Financial_Reports"
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)

        # Updated variables mapping with alternative text variations
        self.variables_mapping = {
            "موجودی نقد": ["موجودی نقد", "وجه نقد"],
            "دارایی‌های جاری": ["دارایی‌های جاری", "دارایی\u200cهای جاری", "جمع دارایی‌های جاری"],
            "موجودی مواد و کالا": ["موجودی مواد و کالا", "موجودی کالا"],
            "بدهی‌های جاری": ["بدهی‌های جاری", "بدهی\u200cهای جاری", "جمع بدهی‌های جاری"],
            "سود خالص": ["سود خالص", "سود (زیان) خالص"],
            "جمع دارایی‌ها": ["جمع دارایی‌ها", "جمع کل دارایی\u200cها"],
            "جمع حقوق مالکانه": ["جمع حقوق مالکانه", "جمع حقوق صاحبان سهام"],
            "فروش": ["درآمدهای عملیاتی", "فروش خالص", "درآمد عملیاتی"],
            "سود عملیاتی": ["سود عملیاتی", "سود (زیان) عملیاتی"],
            "سود ناخالص": ["سود ناخالص", "سود (زیان) ناخالص"],
            "دریافتنی‌های تجاری و سایر دریافتنی‌ها": ["دریافتنی‌های تجاری", "حساب‌های دریافتنی تجاری"],
            "بهای تمام شده کالای فروش رفته": ["بهای تمام‌شده درآمدهای عملیاتی", "بهای تمام شده کالای فروش رفته"],
            "جمع بدهی‌ها": ["جمع بدهی‌ها", "جمع کل بدهی\u200cها"]
        }

    def get_value_by_row(self, df, search_terms):
        """Enhanced value extraction with multiple search terms and special character handling"""
        try:
            if isinstance(search_terms, str):
                search_terms = [search_terms]

            # Convert DataFrame to string type for better text matching
            df = df.astype(str)

            for search_term in search_terms:
                # Handle special characters in Persian text
                normalized_search = search_term.replace('‌', ' ').replace('\u200c', ' ')

                # Try exact match first
                for col in range(df.shape[1]):
                    matches = df[
                        df.iloc[:, col].str.replace('‌', ' ').str.replace('\u200c', ' ').str.contains(normalized_search,
                                                                                                      regex=False,
                                                                                                      na=False)]
                    if not matches.empty:
                        row_idx = matches.index[0]
                        # Look for numeric values in the same row
                        for value_col in range(df.shape[1]):
                            if value_col != col:  # Don't check the column with the label
                                value = df.iloc[row_idx, value_col]
                                if pd.notna(value) and str(value).strip() not in ['', '-']:
                                    try:
                                        # Clean the value
                                        cleaned_value = str(value).strip()
                                        cleaned_value = cleaned_value.replace(',', '').replace('٫', '.')
                                        cleaned_value = cleaned_value.replace('−', '-').replace('(', '-').replace(')',
                                                                                                                  '')

                                        # Convert Persian numbers
                                        for persian, english in zip('۰۱۲۳۴۵۶۷۸۹', '0123456789'):
                                            cleaned_value = cleaned_value.replace(persian, english)

                                        # Convert to Decimal
                                        decimal_value = Decimal(cleaned_value)
                                        if decimal_value != 0:
                                            print(f"Found value for {search_term}: {float(decimal_value):,.2f}")
                                            return decimal_value
                                    except (ValueError, TypeError):
                                        continue

            print(f"No valid value found for {search_terms[0]}")
            return Decimal('0')

        except Exception as e:
            print(f"Error processing {search_terms[0]}: {str(e)}")
            return Decimal('0')

    def safe_divide(self, numerator, denominator):
        """Precise division with comprehensive error checking"""
        try:
            if numerator is None or denominator is None:
                return Decimal('0')

            # Ensure we're working with Decimal objects
            if not isinstance(numerator, Decimal):
                numerator = Decimal(str(numerator))
            if not isinstance(denominator, Decimal):
                denominator = Decimal(str(denominator))

            if abs(denominator) < Decimal('1E-10'):
                return Decimal('0')

            result = numerator / denominator

            if not result.is_finite():
                return Decimal('0')

            return result

        except Exception as e:
            print(f"Division error: {str(e)}")
            return Decimal('0')

    def process_files(self):
        try:
            all_years_data = {
                'variables': {},
                'ratios': {}
            }

            excel_files = sorted([f for f in self.input_folder.glob('*.xlsx')
                                  if not f.name.startswith('~$')])

            for file_path in excel_files:
                try:
                    year = file_path.stem
                    print(f"\nProcessing year {year}...")

                    # Read Excel file without encoding parameter
                    df = pd.read_excel(
                        file_path,
                        engine='openpyxl',
                        header=None,
                        dtype=str  # Read all values as strings initially
                    )

                    # Calculate variables
                    variables = {}
                    for var_key, search_terms in self.variables_mapping.items():
                        raw_value = self.get_value_by_row(df, search_terms)
                        if raw_value == Decimal('0'):
                            print(f"Warning: Zero value found for {var_key} in {year}")
                            continue
                        variables[var_key] = raw_value / Decimal('1000000.0')  # Convert to millions
                        print(f"{var_key}: {float(variables[var_key]):,.10f}")

                    # Calculate ratios with proper error handling
                    try:
                        ratios = {
                            "نسبت جاری": self.safe_divide(variables["دارایی‌های جاری"], variables["بدهی‌های جاری"]),
                            "نسبت آنی": self.safe_divide(
                                variables["دارایی‌های جاری"] - variables["موجودی مواد و کالا"],
                                variables["بدهی‌های جاری"]
                            ),
                            "نسبت نقدی": self.safe_divide(variables["موجودی نقد"], variables["بدهی‌های جاری"]),
                            "بازده دارایی‌ها": self.safe_divide(variables["سود خالص"], variables["جمع دارایی‌ها"]),
                            "بازده حقوق صاحبان سهام": self.safe_divide(variables["سود خالص"],
                                                                       variables["جمع حقوق مالکانه"]),
                            "حاشیه سود خالص": self.safe_divide(variables["سود خالص"], variables["فروش"]),
                            "حاشیه سود عملیاتی": self.safe_divide(variables["سود عملیاتی"], variables["فروش"]),
                            "حاشیه سود ناخالص": self.safe_divide(variables["سود ناخالص"], variables["فروش"]),
                            "دوره وصول مطالبات": self.safe_divide(
                                variables["دریافتنی‌های تجاری و سایر دریافتنی‌ها"] * Decimal('365'),
                                variables["فروش"]
                            ),
                            "گردش مطالبات": self.safe_divide(
                                variables["فروش"],
                                variables["دریافتنی‌های تجاری و سایر دریافتنی‌ها"]
                            ),
                            "گردش موجودی کالا": self.safe_divide(
                                variables["بهای تمام شده کالای فروش رفته"],
                                variables["موجودی مواد و کالا"]
                            ),
                            "نسبت بدهی به دارایی": self.safe_divide(variables["جمع بدهی‌ها"],
                                                                    variables["جمع دارایی‌ها"])
                        }

                        # Store data with high precision
                        all_years_data['variables'][year] = {k: float(v) for k, v in variables.items()}
                        all_years_data['ratios'][year] = {k: float(v) for k, v in ratios.items()}

                    except Exception as e:
                        print(f"Error calculating ratios for {year}: {str(e)}")
                        import traceback
                        print(traceback.format_exc())
                        continue

                except Exception as e:
                    print(f"Error processing file {file_path}: {str(e)}")
                    import traceback
                    print(traceback.format_exc())
                    continue

            return self.create_consolidated_report(all_years_data)

        except Exception as e:
            print(f"Error in process_files: {str(e)}")
            import traceback
            print(traceback.format_exc())
            return None

    def create_consolidated_report(self, all_years_data):
        try:
            filename = self.output_dir / f"Consolidated_Financial_Analysis_{self.current_time}.xlsx"

            with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
                workbook = writer.book

                # Create formats
                header_format = workbook.add_format({
                    'bold': True,
                    'font_color': 'white',
                    'bg_color': '#0066cc',
                    'border': 1,
                    'align': 'right',
                    'font_name': 'B Nazanin',
                    'num_format': '#,##0.0000000000'
                })

                number_format = workbook.add_format({
                    'num_format': '#,##0.0000000000',  # 10 decimal places
                    'border': 1,
                    'align': 'right',
                    'font_name': 'B Nazanin'
                })

                # Create DataFrames with high precision
                variables_df = pd.DataFrame(all_years_data['variables']).round(10)
                ratios_df = pd.DataFrame(all_years_data['ratios']).round(10)

                # Write sheets
                variables_df.to_excel(writer, sheet_name='متغیرهای پایه', index=True)
                ratios_df.to_excel(writer, sheet_name='نسبت‌های مالی', index=True)

                # Format sheets
                for sheet_name in writer.sheets:
                    worksheet = writer.sheets[sheet_name]
                    worksheet.set_column('A:A', 40)
                    worksheet.set_column('B:Z', 20, number_format)
                    worksheet.set_row(0, None, header_format)
                    worksheet.right_to_left()

            print(f"\nConsolidated report created at:\n{filename}")
            return filename

        except Exception as e:
            print(f"Error creating consolidated report: {str(e)}")
            import traceback
            print(traceback.format_exc())
            return None


def main():
    print("Financial Analysis Tool")
    print("-" * 50)

    input_folder = input("Please enter the folder path containing the yearly files: ").strip()

    if not os.path.exists(input_folder):
        print("Invalid path!")
        return

    analyzer = FinancialAnalyzer(input_folder)
    output_file = analyzer.process_files()

    if output_file:
        print("\nProcessing completed successfully!")
        print(f"Output file location: {output_file}")


if __name__ == "__main__":
    main()