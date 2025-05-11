import os
import pandas as pd

# تعریف نسبت‌ها و متغیرها
ratios = [
    "نسبت جاری", "نسبت آنی", "نسبت وجه نقد", "بازده دارایی ها", "بازده حقوق صاحبان سهام",
    "حاشیه سود خالص", "حاشیه سود عملیاتی", "حاشیه سود ناخالص", "دوره وصول مطالبات",
    "نسبت بدهی به دارایی", "گردش حساب دریافتنی", "گردش موجودی کالا"
]

variables = [
    "وجه نقد", "حساب دریافتنی", "موجودی کالا", "سرمایه گذاری کوتاه مدت", "دارایی جاری",
    "کل دارایی ها", "بدهی جاری", "کل بدهی ها", "حقوق صاحبان سهام", "فروش",
    "بهای تمام شده کالای فروش رفته", "سود ناخالص", "سود عملیاتی", "سود خالص",
    "هزینه مالیات", "متوسط فروش روزانه", "میانگین حساب دریافتنی", "میانگین موجودی کالا"
]


def combine_financial_data(folder_path, output_file):
    ratios_df = pd.DataFrame(columns=["سال"] + ratios)
    variables_df = pd.DataFrame(columns=["سال"] + variables)

    files = [os.path.join(folder_path, file) for file in os.listdir(folder_path) if file.endswith(".xls")]

    for file in files:
        year = os.path.splitext(os.path.basename(file))[0].split("_")[-1]
        data = pd.read_excel(file, engine='xlrd')

        # تمیز کردن ستون‌ها
        data.columns = data.columns.str.strip()

        # بررسی نسبت‌ها
        ratio_data = {ratio: data[ratio].iloc[0] if ratio in data.columns else None for ratio in ratios}
        for ratio in ratios:
            if ratio not in data.columns:
                print(f"ستون {ratio} در فایل {file} پیدا نشد.")

        # بررسی متغیرها
        variable_data = {variable: data[variable].iloc[0] if variable in data.columns else None for variable in variables}
        for variable in variables:
            if variable not in data.columns:
                print(f"ستون {variable} در فایل {file} پیدا نشد.")

        ratios_df = pd.concat([ratios_df, pd.DataFrame([{"سال": year, **ratio_data}])], ignore_index=True)
        variables_df = pd.concat([variables_df, pd.DataFrame([{"سال": year, **variable_data}])], ignore_index=True)

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        ratios_df.to_excel(writer, sheet_name="نسبت‌ها", index=False)
        variables_df.to_excel(writer, sheet_name="متغیرها", index=False)

    print(f"فایل ترکیب‌شده با موفقیت ذخیره شد: {output_file}")


if __name__ == "__main__":
    folder_path = r"C:\retina_env\BOT\New folder" # مسیر پوشه را اینجا وارد کنید
    output_file = "combined_financial_data.xlsx"
    combine_financial_data(folder_path, output_file)


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

                    # Read Excel file with all sheets
                    df = pd.read_excel(
                        file_path,
                        engine='openpyxl',
                        header=None,
                        dtype=str  # Read all values as strings initially
                    )

                    # Calculate variables
                    variables = {}
                    for var_key, search_term in self.variables_mapping.items():
                        raw_value = self.get_value_by_row(df, search_term)
                        variables[var_key] = raw_value / Decimal('1000000.0')  # Convert to millions
                        print(f"{var_key}: {float(variables[var_key]):,.10f}")

                    # Calculate ratios with proper error handling
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
                        "نسبت بدهی به دارایی": self.safe_divide(variables["جمع بدهی‌ها"], variables["جمع دارایی‌ها"])
                    }

                    # Store data
                    all_years_data['variables'][year] = {k: float(v) for k, v in variables.items()}
                    all_years_data['ratios'][year] = {k: float(v) for k, v in ratios.items()}

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
