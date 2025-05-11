
import pandas as pd
import numpy as np
from decimal import Decimal, getcontext, ROUND_HALF_UP, DivisionByZero, InvalidOperation
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
            "موجودی نقد": ["موجودی نقد", "وجه نقد", "موجودی نقد و معادل نقد", "نقد و معادل نقد"],
            "دارایی‌های جاری": [
                "دارایی‌های جاری",
                "دارایی های جاری",
                "جمع دارایی‌های جاری",
                "جمع دارایی های جاری",
                "جمع دارایی جاری"
            ],
            "موجودی مواد و کالا": [
                "موجودی مواد و کالا",
                "موجودی کالا",
                "موجودی‌ها",
                "موجودی مواد، کالا و قطعات"
            ],
            "بدهی‌های جاری": [
                "بدهی‌های جاری",
                "بدهی های جاری",
                "جمع بدهی‌های جاری",
                "جمع بدهی های جاری",
                "جمع بدهی جاری"
            ],
            "سود خالص": [
                "سود خالص",
                "سود (زیان) خالص",
                "سود و زیان خالص",
                "سود (زیان) خالص دوره"
            ],
            "جمع دارایی‌ها": [
                "جمع دارایی‌ها",
                "جمع کل دارایی‌ها",
                "جمع دارایی ها",
                "جمع کل دارایی ها"
            ],
            "جمع حقوق مالکانه": [
                "جمع حقوق مالکانه",
                "جمع حقوق صاحبان سهام",
                "حقوق صاحبان سهام"
            ],
            "فروش": [
                "درآمدهای عملیاتی",
                "فروش خالص",
                "درآمد عملیاتی",
                "جمع درآمدهای عملیاتی",
                "فروش"
            ],
            "سود عملیاتی": [
                "سود عملیاتی",
                "سود (زیان) عملیاتی",
                "سود و زیان عملیاتی"
            ],
            "سود ناخالص": [
                "سود ناخالص",
                "سود (زیان) ناخالص",
                "سود و زیان ناخالص"
            ],
            "دریافتنی‌های تجاری و سایر دریافتنی‌ها": [
                "دریافتنی‌های تجاری",
                "حساب‌های دریافتنی تجاری",
                "دریافتنی های تجاری",
                "حساب های دریافتنی"
            ],
            "بهای تمام شده کالای فروش رفته": [
                "بهای تمام‌شده درآمدهای عملیاتی",
                "بهای تمام شده کالای فروش رفته",
                "بهای تمام شده درآمدهای عملیاتی",
                "بهای تمام شده فروش"
            ],
            "جمع بدهی‌ها": [
                "جمع بدهی‌ها",
                "جمع کل بدهی‌ها",
                "جمع بدهی ها",
                "جمع کل بدهی ها"
            ],
            "موجودی نقد": [
                "موجودی نقد",
                "وجه نقد",
                "موجودی نقد و معادل نقد",
                "نقد",
                "وجوه نقد"
            ]
        }

    def get_value_by_row(self, df, search_terms):
        """Enhanced value extraction with better pattern matching for Persian financial statements"""
        try:
            if isinstance(search_terms, str):
                search_terms = [search_terms]

            # Convert DataFrame to string type and clean it
            df = df.astype(str)

            for search_term in search_terms:
                # Clean search term
                clean_search = search_term.replace('‌', ' ').replace('\u200c', ' ').strip()

                # Search through all columns and rows
                for col in range(df.shape[1]):
                    for idx, row in df.iterrows():
                        cell_value = str(row.iloc[col]).replace('‌', ' ').replace('\u200c', ' ').strip()

                        # Check if search term is in the cell
                        if clean_search in cell_value:
                            # Look for numeric values in all columns of this row
                            for value_col in range(df.shape[1]):
                                value = str(row.iloc[value_col]).strip()

                                # Skip empty or non-numeric cells
                                if not value or value in ['-', 'nan', 'None']:
                                    continue

                                try:
                                    # Clean and convert the value
                                    cleaned_value = value.replace(',', '')
                                    cleaned_value = cleaned_value.replace('٫', '.')
                                    cleaned_value = cleaned_value.replace('−', '-')
                                    cleaned_value = cleaned_value.replace('(', '-')
                                    cleaned_value = cleaned_value.replace(')', '')

                                    # Convert Persian numbers to English
                                    persian_numbers = '۰۱۲۳۴۵۶۷۸۹'
                                    english_numbers = '0123456789'
                                    for persian, english in zip(persian_numbers, english_numbers):
                                        cleaned_value = cleaned_value.replace(persian, english)

                                    # Remove any remaining non-numeric characters except decimal point and minus
                                    cleaned_value = ''.join(c for c in cleaned_value
                                                            if c.isdigit() or c in '.-')

                                    if cleaned_value:
                                        decimal_value = Decimal(cleaned_value)
                                        if decimal_value != 0:
                                            print(f"Found value for {search_term}: {float(decimal_value):,.2f}")
                                            return decimal_value
                                except (ValueError, TypeError, InvalidOperation):
                                    continue

                print(f"No valid value found for {search_terms[0]}")
                return Decimal('0')

            return Decimal('0')

        except Exception as e:
            print(f"Error processing {search_terms[0]}: {str(e)}")
            return Decimal('0')



    def safe_divide(self, numerator, denominator):
        """Precise division for financial calculations"""
        try:
            if numerator is None or denominator is None or pd.isna(numerator) or pd.isna(denominator):
                return Decimal('0')

            # Convert to Decimal
            if not isinstance(numerator, Decimal):
                numerator = Decimal(str(numerator))
            if not isinstance(denominator, Decimal):
                denominator = Decimal(str(denominator))

            # Check for zero denominator
            if abs(denominator) < Decimal('1E-28'):
                return Decimal('0')

            result = numerator / denominator
            if not result.is_finite():
                return Decimal('0')

            return result.quantize(Decimal('0.0000000000'), rounding=ROUND_HALF_UP)

        except Exception as e:
            print(f"Division error: {str(e)}")
            return Decimal('0')

    # [Previous process_files and create_consolidated_report methods remain the same]


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

                    # Read Excel file
                    # In the process_files method, update the Excel reading part:

                    df = pd.read_excel(
                        file_path,
                        engine='openpyxl',
                        header=None,
                        dtype=str,
                        na_filter=False  # This prevents pandas from converting empty cells to NaN
                    )

                    # Remove any completely empty rows and columns
                    df = df.dropna(how='all').dropna(axis=1, how='all')

                    # Initialize variables dictionary with zeros for all keys
                    variables = {key: Decimal('0') for key in self.variables_mapping.keys()}

                    # Calculate variables
                    # In process_files method, update the variables calculation part:

                    # Calculate variables with multiple attempts
                    for var_key, search_terms in self.variables_mapping.items():
                        raw_value = Decimal('0')

                        # Try each search term
                        for term in search_terms:
                            raw_value = self.get_value_by_row(df, term)
                            if raw_value != Decimal('0'):
                                break

                        if raw_value != Decimal('0'):
                            # Convert to millions and store
                            variables[var_key] = raw_value / Decimal('1000000.0')
                            print(f"{var_key}: {float(variables[var_key]):,.10f}")
                        else:
                            print(f"Warning: No value found for {var_key} in {year}")

                    # Initialize ratios dictionary
                    ratios = {}

                    # Calculate ratios only if we have valid values
                    try:
                        if variables["دارایی‌های جاری"] != 0 and variables["بدهی‌های جاری"] != 0:
                            ratios["نسبت جاری"] = self.safe_divide(variables["دارایی‌های جاری"],
                                                                   variables["بدهی‌های جاری"])

                            ratios["نسبت آنی"] = self.safe_divide(
                                variables["دارایی‌های جاری"] - variables["موجودی مواد و کالا"],
                                variables["بدهی‌های جاری"]
                            )

                            ratios["نسبت نقدی"] = self.safe_divide(variables["موجودی نقد"],
                                                                   variables["بدهی‌های جاری"])

                        if variables["سود خالص"] != 0:
                            ratios["بازده دارایی‌ها"] = self.safe_divide(variables["سود خالص"],
                                                                         variables["جمع دارایی‌ها"])
                            ratios["بازده حقوق صاحبان سهام"] = self.safe_divide(variables["سود خالص"],
                                                                                variables["جمع حقوق مالکانه"])
                            ratios["حاشیه سود خالص"] = self.safe_divide(variables["سود خالص"],
                                                                        variables["فروش"])

                        if variables["سود عملیاتی"] != 0:
                            ratios["حاشیه سود عملیاتی"] = self.safe_divide(variables["سود عملیاتی"],
                                                                           variables["فروش"])

                        if variables["سود ناخالص"] != 0:
                            ratios["حاشیه سود ناخالص"] = self.safe_divide(variables["سود ناخالص"],
                                                                          variables["فروش"])

                        if variables["فروش"] != 0:
                            ratios["دوره وصول مطالبات"] = self.safe_divide(
                                variables["دریافتنی‌های تجاری و سایر دریافتنی‌ها"] * Decimal('365'),
                                variables["فروش"]
                            )

                            ratios["گردش مطالبات"] = self.safe_divide(
                                variables["فروش"],
                                variables["دریافتنی‌های تجاری و سایر دریافتنی‌ها"]
                            )

                        if variables["موجودی مواد و کالا"] != 0:
                            ratios["گردش موجودی کالا"] = self.safe_divide(
                                variables["بهای تمام شده کالای فروش رفته"],
                                variables["موجودی مواد و کالا"]
                            )

                        if variables["جمع دارایی‌ها"] != 0:
                            ratios["نسبت بدهی به دارایی"] = self.safe_divide(variables["جمع بدهی‌ها"],
                                                                             variables["جمع دارایی‌ها"])

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

                # In create_consolidated_report method:
                number_format = workbook.add_format({
                    'num_format': '#,##0.0000000000',
                    'border': 1,
                    'align': 'right',
                    'font_name': 'B Nazanin'
                })

                # Create DataFrames with exact precision
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