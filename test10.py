import pandas as pd
import numpy as np
from datetime import datetime
import os
from pathlib import Path
import warnings

warnings.filterwarnings('ignore')


class FinancialAnalyzer:
    def __init__(self, input_folder_path):
        self.input_folder = Path(input_folder_path)
        self.current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.output_dir = self.input_folder / "Financial_Reports"
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)

    def get_value_by_row(self, df, var_name):
        """Enhanced value extraction with maximum precision"""
        try:
            # Store all potential matches
            matches = []

            # First pass: Look for exact matches
            for idx, row in df.iterrows():
                if pd.notna(row[0]) and str(row[0]).strip() == var_name:
                    for col in range(1, len(df.columns)):
                        value = df.iloc[idx, col]
                        if pd.notna(value):
                            try:
                                # Handle different numeric formats
                                if isinstance(value, (int, float)):
                                    matches.append(float(value))
                                elif isinstance(value, str):
                                    # Remove grouping separators and handle different decimal symbols
                                    cleaned = value.replace(',', '').replace('٫', '.').strip()
                                    if cleaned and cleaned != '-':
                                        matches.append(float(cleaned))
                            except ValueError:
                                continue

            # Second pass: Look for partial matches if no exact match found
            if not matches:
                for idx, row in df.iterrows():
                    if pd.notna(row[0]) and var_name in str(row[0]):
                        for col in range(1, len(df.columns)):
                            value = df.iloc[idx, col]
                            if pd.notna(value):
                                try:
                                    if isinstance(value, (int, float)):
                                        matches.append(float(value))
                                    elif isinstance(value, str):
                                        cleaned = value.replace(',', '').replace('٫', '.').strip()
                                        if cleaned and cleaned != '-':
                                            matches.append(float(cleaned))
                                except ValueError:
                                    continue

            # If we found matches, return the most appropriate one
            if matches:
                # Filter out obvious outliers
                filtered_matches = [m for m in matches if m != 0]
                if filtered_matches:
                    # Return the largest absolute value (assuming it's the most significant)
                    value = max(filtered_matches, key=abs)
                    print(f"Found value for {var_name}: {value:,.2f}")
                    return value

            print(f"No valid value found for {var_name}")
            return 0.0

        except Exception as e:
            print(f"Error processing {var_name}: {str(e)}")
            return 0.0

    def safe_divide(numerator, denominator):
        """Enhanced division with maximum precision"""
        try:
            if denominator is None or numerator is None:
                return 0.0
            if abs(denominator) < np.finfo(np.float64).eps:
                return 0.0

            # Convert to high precision floating point
            num = np.float64(numerator)
            den = np.float64(denominator)

            result = num / den

            # Check for invalid results
            if np.isinf(result) or np.isnan(result):
                return 0.0

            return result

        except Exception as e:
            print(f"Division error: {str(e)}")
            return 0.0

    # In the process_files method, update the variables calculation:
    variables = {}
    for var_key, var_name in variable_mappings.items():
        raw_value = self.get_value_by_row(all_data, var_name)
        # Use high precision conversion
        variables[var_key] = np.float64(raw_value) / np.float64(1_000_000.0)
        print(f"{var_key}: {variables[var_key]:,.6f}")

    # And in create_consolidated_report, update the DataFrame creation:
    variables_df = pd.DataFrame(all_years_data['variables']).round(6)  # Increased precision
    ratios_df = pd.DataFrame(all_years_data['ratios']).round(6)  # Increased precision

    # Update the number format in the Excel output:
    number_format = workbook.add_format({
        'num_format': '#,##0.000000',  # Increased decimal places
        'border': 1,
        'align': 'right',
        'font_name': 'B Nazanin'
    })

    def process_files(self):
        try:
            all_years_data = {
                'variables': {},
                'ratios': {}
            }

            excel_files = sorted([f for f in self.input_folder.glob('*.xlsx')
                                  if not f.name.startswith('~$')])

            # Correct row numbers based on your Excel structure
            variables_mapping = {
                "موجودی نقد": 92,
                "دارایی‌های جاری": 87,
                "موجودی مواد و کالا": 89,
                "بدهی‌های جاری": 116,
                "سود خالص": 59,
                "جمع دارایی‌ها": 96,
                "جمع حقوق مالکانه": 109,
                "فروش": 101,
                "سود عملیاتی": 20,
                "سود ناخالص": 15,
                "دریافتنی‌های تجاری و سایر دریافتنی‌ها": 98,
                "بهای تمام شده کالای فروش رفته": 14,
                "جمع بدهی‌ها": 135,
                "سرمایه‌گذاری‌های کوتاه‌مدت": 99
            }

            for file_path in excel_files:
                try:
                    year = file_path.stem
                    print(f"\nProcessing year {year}...")

                    df = pd.read_excel(
                        file_path,
                        engine='openpyxl',
                        header=None,
                        dtype=object
                    )

                    # Calculate variables
                    variables = {}
                    for var_name, row_num in variables_mapping.items():
                        value = self.get_value_by_row(df, row_num)
                        variables[var_name] = value / 1_000_000.0  # Convert to millions
                        print(f"{var_name}: {value:,.2f} -> {variables[var_name]:,.6f}")

                    def safe_divide(numerator, denominator):
                        """Precise division with error checking"""
                        try:
                            if denominator is None or numerator is None:
                                return 0.0
                            if abs(denominator) < 1e-10:
                                return 0.0
                            result = float(numerator) / float(denominator)
                            if np.isinf(result) or np.isnan(result):
                                return 0.0
                            return result
                        except:
                            return 0.0

                    # Calculate ratios
                    ratios = {
                        "نسبت جاری": safe_divide(variables["دارایی‌های جاری"], variables["بدهی‌های جاری"]),
                        "نسبت آنی": safe_divide(
                            variables["دارایی‌های جاری"] - variables["موجودی مواد و کالا"],
                            variables["بدهی‌های جاری"]
                        ),
                        "نسبت نقدی": safe_divide(variables["موجودی نقد"], variables["بدهی‌های جاری"]),
                        "بازده دارایی‌ها": safe_divide(variables["سود خالص"], variables["جمع دارایی‌ها"]),
                        "بازده حقوق صاحبان سهام": safe_divide(variables["سود خالص"], variables["جمع حقوق مالکانه"]),
                        "حاشیه سود خالص": safe_divide(variables["سود خالص"], variables["فروش"]),
                        "حاشیه سود عملیاتی": safe_divide(variables["سود عملیاتی"], variables["فروش"]),
                        "حاشیه سود ناخالص": safe_divide(variables["سود ناخالص"], variables["فروش"]),
                        "دوره وصول مطالبات": safe_divide(
                            variables["دریافتنی‌های تجاری و سایر دریافتنی‌ها"] * 365,
                            variables["فروش"]
                        ),
                        "گردش مطالبات": safe_divide(
                            variables["فروش"],
                            variables["دریافتنی‌های تجاری و سایر دریافتنی‌ها"]
                        ),
                        "گردش موجودی کالا": safe_divide(
                            variables["بهای تمام شده کالای فروش رفته"],
                            variables["موجودی مواد و کالا"]
                        ),
                        "نسبت بدهی به دارایی": safe_divide(variables["جمع بدهی‌ها"], variables["جمع دارایی‌ها"])
                    }

                    # Store data
                    all_years_data['variables'][year] = variables
                    all_years_data['ratios'][year] = ratios

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
                    'num_format': '#,##0.00'
                })

                number_format = workbook.add_format({
                    'num_format': '#,##0.00',
                    'border': 1,
                    'align': 'right',
                    'font_name': 'B Nazanin'
                })

                # Create DataFrames
                variables_df = pd.DataFrame(all_years_data['variables']).round(6)
                ratios_df = pd.DataFrame(all_years_data['ratios']).round(6)

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