import pandas as pd
import numpy as np
from decimal import Decimal, getcontext, ROUND_HALF_UP
from pathlib import Path
import os
import warnings
import matplotlib.pyplot as plt
import seaborn as sns
from typing import Dict, List, Union
import re
import pandas as pd
from decimal import Decimal

# تنظیمات اولیه
warnings.filterwarnings('ignore')
getcontext().prec = 28
sns.set_theme(style="whitegrid")  # تنظیم تم seaborn
plt.rcParams.update({
    'font.family': 'B Nazanin',
    'figure.facecolor': 'white',
    'axes.grid': True,
    'grid.color': '#CCCCCC',
    'grid.linestyle': '--',
    'grid.alpha': 0.3,
    'axes.labelsize': 12,
    'axes.titlesize': 14,
    'xtick.labelsize': 10,
    'ytick.labelsize': 10,
    'legend.fontsize': 10,
    'lines.linewidth': 2,
    'lines.markersize': 8
})


class FinancialAnalyzer:
    def __init__(self, input_folder: str):
        """تنظیمات اولیه کلاس تحلیلگر مالی"""
        self.input_folder = Path(input_folder)
        self.current_time = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
        self.output_dir = self.input_folder / "Financial_Reports"
        self.output_dir.mkdir(exist_ok=True)

        # تعریف نسبت‌ها
        self.ratios = {
            "نسبت جاری": "دارایی جاری / بدهی جاری",
            "نسبت آنی": "(دارایی جاری - موجودی کالا) / بدهی جاری",
            "نسبت وجه نقد": "وجه نقد / بدهی جاری",
            "بازده دارایی ها": "(سود خالص / کل دارایی ها) * 100",
            "بازده حقوق صاحبان سهام": "(سود خالص / حقوق صاحبان سهام) * 100",
            "حاشیه سود خالص": "(سود خالص / فروش) * 100",
            "حاشیه سود عملیاتی": "(سود عملیاتی / فروش) * 100",
            "حاشیه سود ناخالص": "(سود ناخالص / فروش) * 100",
            "دوره وصول مطالبات": "(حساب دریافتنی * 365) / فروش",
            "گردش حساب دریافتنی": "فروش / حساب دریافتنی",
            "گردش موجودی کالا": "بهای تمام شده کالای فروش رفته / موجودی کالا",
            "نسبت بدهی به دارایی": "(کل بدهی ها / کل دارایی ها) * 100"
        }

        # عبارات جستجو برای متغیرها
        self.search_terms = {
            "وجه نقد": [
                "موجودی نقد", "وجه نقد", "موجودی نقد و بانک",
                "نقد و معادل نقد", "موجودی نقد و معادل نقد"
            ],
            "حساب دریافتنی": [
                "حسابهای دریافتنی تجاری", "دریافتنی‌های تجاری",
                "حساب‌های دریافتنی", "حسابها و اسناد دریافتنی تجاری"
            ],
            "موجودی کالا": [
                "موجودی مواد و کالا", "موجودی کالا", "موجودی‌ها",
                "موجودی مواد، کالا و قطعات", "موجودی مواد اولیه و کالا"
            ],
            "دارایی جاری": [
                "جمع دارایی‌های جاری", "دارایی‌های جاری",
                "جمع دارایی های جاری", "دارایی های جاری"
            ],
            "کل دارایی ها": [
                "جمع دارایی‌ها", "جمع کل دارایی‌ها",
                "دارایی‌ها", "جمع دارایی ها"
            ],
            "بدهی جاری": [
                "جمع بدهی‌های جاری", "بدهی‌های جاری",
                "جمع بدهی های جاری", "بدهی های جاری"
            ],
            "کل بدهی ها": [
                "جمع بدهی‌ها", "جمع کل بدهی‌ها",
                "بدهی‌ها", "جمع بدهی ها"
            ],
            "حقوق صاحبان سهام": [
                "جمع حقوق صاحبان سهام", "حقوق صاحبان سهام",
                "جمع حقوق مالکانه", "حقوق مالکانه"
            ],
            "فروش": [
                "درآمدهای عملیاتی", "فروش", "فروش خالص",
                "درآمد عملیاتی", "جمع درآمدهای عملیاتی"
            ],
            "بهای تمام شده کالای فروش رفته": [
                "بهای تمام‌شده درآمدهای عملیاتی",
                "بهای تمام شده کالای فروش رفته",
                "بهای تمام شده فروش"
            ],
            "سود ناخالص": [
                "سود ناخالص", "سود (زیان) ناخالص",
                "سود و زیان ناخالص"
            ],
            "سود عملیاتی": [
                "سود عملیاتی", "سود (زیان) عملیاتی",
                "سود و زیان عملیاتی"
            ],
            "سود خالص": [
                "سود خالص", "سود (زیان) خالص",
                "سود (زیان) خالص دوره"
            ]
        }

    def clean_persian_text(self, text: str) -> str:
        """پاکسازی متن فارسی"""
        if not isinstance(text, str):
            return ''
        # حذف فاصله‌های اضافی و یکسان‌سازی کاراکترها
        text = re.sub(r'\s+', ' ', text)
        text = text.replace('ي', 'ی').replace('ك', 'ک')
        text = text.replace('٠', '۰').replace('١', '۱').replace('٢', '۲')
        text = text.replace('٣', '۳').replace('٤', '۴').replace('٥', '۵')
        text = text.replace('٦', '۶').replace('٧', '۷').replace('٨', '۸')
        text = text.replace('٩', '۹')
        return text.strip()

    def convert_to_number(self, value: str) -> Decimal:
        """تبدیل متن به عدد"""
        try:
            if pd.isna(value) or not str(value).strip():
                return Decimal('0')

            value = str(value)
            # تبدیل اعداد فارسی به انگلیسی
            persian_nums = '۰۱۲۳۴۵۶۷۸۹'
            english_nums = '0123456789'
            for persian, english in zip(persian_nums, english_nums):
                value = value.replace(persian, english)

            # پاکسازی متن
            value = value.replace(',', '').replace('٬', '')
            value = value.replace('(', '-').replace(')', '')
            value = value.replace('−', '-').replace('–', '-')

            # حذف همه کاراکترها بجز اعداد و علامت‌های خاص
            value = ''.join(c for c in value if c.isdigit() or c in '.-')

            if value:
                return Decimal(value)
            return Decimal('0')

        except Exception as e:
            print(f"خطا در تبدیل مقدار {value} به عدد: {str(e)}")
            return Decimal('0')

    def find_value_in_df(self, df: pd.DataFrame, search_terms: List[str]) -> Decimal:
        """یافتن مقدار در دیتافریم"""
        try:
            for term in search_terms:
                term = self.clean_persian_text(term)
                for idx, row in df.iterrows():
                    for col in df.columns:
                        cell_value = self.clean_persian_text(str(row[col]))
                        if term in cell_value:
                            # جستجو در همان ردیف برای یافتن عدد
                            for value_col in df.columns:
                                value = str(row[value_col])
                                number = self.convert_to_number(value)
                                if number != 0:
                                    return number
            return Decimal('0')

        except Exception as e:
            print(f"خطا در جستجوی مقدار: {str(e)}")
            return Decimal('0')

    def calculate_financial_ratios(self, variables: Dict[str, Decimal]) -> Dict[str, Decimal]:
        """محاسبه نسبت‌های مالی"""
        try:
            ratios = {}

            # نسبت‌های نقدینگی
            if variables["بدهی جاری"] != 0:
                ratios["نسبت جاری"] = (variables["دارایی جاری"] /
                                       variables["بدهی جاری"]).quantize(Decimal('0.01'))

                ratios["نسبت آنی"] = ((variables["دارایی جاری"] - variables["موجودی کالا"]) /
                                      variables["بدهی جاری"]).quantize(Decimal('0.01'))

                ratios["نسبت وجه نقد"] = (variables["وجه نقد"] /
                                          variables["بدهی جاری"]).quantize(Decimal('0.01'))

            # نسبت‌های سودآوری
            if variables["کل دارایی ها"] != 0:
                ratios["بازده دارایی ها"] = ((variables["سود خالص"] / variables["کل دارایی ها"]) *
                                             Decimal('100')).quantize(Decimal('0.01'))

            if variables["حقوق صاحبان سهام"] != 0:
                ratios["بازده حقوق صاحبان سهام"] = ((variables["سود خالص"] /
                                                     variables["حقوق صاحبان سهام"]) *
                                                    Decimal('100')).quantize(Decimal('0.01'))

            if variables["فروش"] != 0:
                ratios["حاشیه سود خالص"] = ((variables["سود خالص"] / variables["فروش"]) *
                                            Decimal('100')).quantize(Decimal('0.01'))

                ratios["حاشیه سود عملیاتی"] = ((variables["سود عملیاتی"] / variables["فروش"]) *
                                               Decimal('100')).quantize(Decimal('0.01'))

                ratios["حاشیه سود ناخالص"] = ((variables["سود ناخالص"] / variables["فروش"]) *
                                              Decimal('100')).quantize(Decimal('0.01'))

            # نسبت‌های فعالیت
            if variables["فروش"] != 0:
                ratios["دوره وصول مطالبات"] = ((variables["حساب دریافتنی"] * Decimal('365')) /
                                               variables["فروش"]).quantize(Decimal('0.01'))

            if variables["حساب دریافتنی"] != 0:
                ratios["گردش حساب دریافتنی"] = (variables["فروش"] /
                                                variables["حساب دریافتنی"]).quantize(Decimal('0.01'))

            if variables["موجودی کالا"] != 0:
                ratios["گردش موجودی کالا"] = (variables["بهای تمام شده کالای فروش رفته"] /
                                              variables["موجودی کالا"]).quantize(Decimal('0.01'))

            # نسبت‌های اهرمی
            if variables["کل دارایی ها"] != 0:
                ratios["نسبت بدهی به دارایی"] = ((variables["کل بدهی ها"] /
                                                  variables["کل دارایی ها"]) *
                                                 Decimal('100')).quantize(Decimal('0.01'))

            """محاسبه نسبت‌های مالی"""
            return {
                "نسبت جاری": safe_divide(
                    variable_data["دارایی جاری"],
                    variable_data["بدهی جاری"]
                ),
                "نسبت آنی": safe_divide(
                    variable_data["دارایی جاری"] - variable_data["موجودی کالا"],
                    variable_data["بدهی جاری"]
                ),
                "نسبت وجه نقد": safe_divide(
                    variable_data["وجه نقد"],
                    variable_data["بدهی جاری"]
                ),
                "بازده دارایی ها": safe_divide(
                    variable_data["سود خالص"],
                    variable_data["کل دارایی ها"]
                ) * 100,
                "بازده حقوق صاحبان سهام": safe_divide(
                    variable_data["سود خالص"],
                    variable_data["حقوق صاحبان سهام"]
                ) * 100,
                "حاشیه سود خالص": safe_divide(
                    variable_data["سود خالص"],
                    variable_data["فروش"]
                ) * 100,
                "حاشیه سود عملیاتی": safe_divide(
                    variable_data["سود عملیاتی"],
                    variable_data["فروش"]
                ) * 100,
                "حاشیه سود ناخالص": safe_divide(
                    variable_data["سود ناخالص"],
                    variable_data["فروش"]
                ) * 100,
                "دوره وصول مطالبات": safe_divide(
                    variable_data["حساب دریافتنی"] * 365,
                    variable_data["فروش"]
                ),
                "گردش حساب دریافتنی": safe_divide(
                    variable_data["فروش"],
                    variable_data["حساب دریافتنی"]
                ),
                "گردش موجودی کالا": safe_divide(
                    variable_data["بهای تمام شده کالای فروش رفته"],
                    variable_data["موجودی کالا"]
                ),
                "نسبت بدهی به دارایی": safe_divide(
                    variable_data["کل بدهی ها"],
                    variable_data["کل دارایی ها"]
                ) * 100
            }

        except Exception as e:
            print(f"خطا در محاسبه نسبت‌های مالی: {str(e)}")
            return {}

        def calculate_ratios(variable_data):
            """محاسبه نسبت‌های مالی"""
            try:
                sales = variable_data.get("فروش", 0)
                if sales == 0:
                    sales = 1  # برای جلوگیری از تقسیم بر صفر

                # محاسبه بهای تمام شده اگر صفر باشد
                if variable_data["بهای تمام شده کالای فروش رفته"] == 0:
                    gross_profit = variable_data.get("سود ناخالص", 0)
                    if sales > 0 and gross_profit > 0:
                        cogs = sales - gross_profit
                        if cogs > 0:
                            variable_data["بهای تمام شده کالای فروش رفته"] = cogs
                            print(f"بهای تمام شده محاسبه شده: {cogs:,.0f}")

                ratios_data = {
                    "نسبت جاری": safe_divide(variable_data["دارایی جاری"], variable_data["بدهی جاری"]),
                    "نسبت آنی": safe_divide(
                        variable_data["دارایی جاری"] - variable_data["موجودی کالا"],
                        variable_data["بدهی جاری"]
                    ),
                    "نسبت وجه نقد": safe_divide(variable_data["وجه نقد"], variable_data["بدهی جاری"]),
                    "بازده دارایی ها": safe_divide(variable_data["سود خالص"], variable_data["کل دارایی ها"]) * 100,
                    "بازده حقوق صاحبان سهام": safe_divide(variable_data["سود خالص"],
                                                          variable_data["حقوق صاحبان سهام"]) * 100,
                    "حاشیه سود خالص": safe_divide(variable_data["سود خالص"], sales) * 100,
                    "حاشیه سود عملیاتی": safe_divide(variable_data["سود عملیاتی"], sales) * 100,
                    "حاشیه سود ناخالص": safe_divide(variable_data["سود ناخالص"], sales) * 100,
                    "دوره وصول مطالبات": safe_divide(variable_data["حساب دریافتنی"] * 365, sales),
                    "گردش حساب دریافتنی": safe_divide(sales, variable_data["حساب دریافتنی"]),
                    "گردش موجودی کالا": safe_divide(variable_data["بهای تمام شده کالای فروش رفته"],
                                                    variable_data["موجودی کالا"]),
                    "نسبت بدهی به دارایی": safe_divide(variable_data["کل بدهی ها"], variable_data["کل دارایی ها"]) * 100
                }
                return ratios_data
            except Exception as e:
                print(f"خطا در محاسبه نسبت‌های مالی: {str(e)}")
                return {}

        def safe_convert_to_number(value):
            """تبدیل مقادیر به عدد با در نظر گرفتن حالت‌های مختلف"""
            if pd.isna(value):
                return 0
            if isinstance(value, (int, float)):
                return value
            try:
                # حذف کاراکترهای غیر عددی و تبدیل به عدد
                cleaned_value = str(value).replace(',', '').replace(' ', '')
                if cleaned_value.strip() in ('', '-', '--'):
                    return 0
                return float(cleaned_value)
            except:
                print(f"خطا در تبدیل مقدار {value} به عدد")
                return 0


        def safe_divide(self, numerator, denominator):

            try:
                if numerator is None or denominator is None:
                    return Decimal('0')

                numerator = Decimal(str(numerator))
                denominator = Decimal(str(denominator))

                # بررسی تقسیم بر صفر
                if denominator == 0:
                    return Decimal('0')

                result = numerator / denominator
                return result.quantize(Decimal('0.0000000000'), rounding=ROUND_HALF_UP)

            except Exception as e:
                print(f"خطا در تقسیم: {str(e)}")
                return Decimal('0')

    def process_file(self, file_path: Path) -> tuple:
        """پردازش فایل اکسل و استخراج داده‌ها"""
        try:
            print(f"\nدر حال پردازش فایل: {file_path.name}")

            # خواندن فایل اکسل
            df = pd.read_excel(
                file_path,
                header=None,
                dtype=str,
                na_filter=False
            )

            # استخراج متغیرها
            variables = {}
            for var_name, search_terms in self.search_terms.items():
                value = self.find_value_in_df(df, search_terms)
                variables[var_name] = value
                print(f"{var_name}: {float(value):,.0f}")

            # محاسبه نسبت‌ها
            ratios = self.calculate_financial_ratios(variables)

            # نمایش نسبت‌ها
            print("\nنسبت‌های مالی محاسبه شده:")
            for ratio_name, ratio_value in ratios.items():
                print(f"{ratio_name}: {float(ratio_value):.2f}")

            return variables, ratios

        except Exception as e:
            print(f"خطا در پردازش فایل {file_path.name}: {str(e)}")
            return {}, {}

    def create_excel_report(self, all_data: Dict, company_name: str):
        """ایجاد گزارش اکسل"""
        try:
            output_file = self.output_dir / f"{company_name}_Financial_Analysis_{self.current_time}.xlsx"

            with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                workbook = writer.book

                # تعریف فرمت‌ها
                header_format = workbook.add_format({
                    'bold': True,
                    'font_color': 'white',
                    'bg_color': '#0066cc',
                    'border': 1,
                    'align': 'center',
                    'valign': 'vcenter',
                    'font_name': 'B Nazanin',
                    'font_size': 12
                })

                number_format = workbook.add_format({
                    'num_format': '#,##0.00',
                    'border': 1,
                    'align': 'center',
                    'valign': 'vcenter',
                    'font_name': 'B Nazanin',
                    'font_size': 11
                })

                # تبدیل داده‌ها به دیتافریم
                variables_df = pd.DataFrame(all_data['variables']).T
                ratios_df = pd.DataFrame(all_data['ratios']).T

                # نوشتن در شیت‌ها
                variables_df.to_excel(writer, sheet_name='متغیرها')
                ratios_df.to_excel(writer, sheet_name='نسبت‌ها')

                # فرمت‌بندی شیت‌ها
                for sheet_name in writer.sheets:
                    worksheet = writer.sheets[sheet_name]
                    worksheet.set_column('A:Z', 15)
                    worksheet.right_to_left()

            print(f"\nگزارش اکسل در مسیر زیر ایجاد شد:\n{output_file}")
            return output_file

        except Exception as e:
            print(f"خطا در ایجاد گزارش اکسل: {str(e)}")
            return None

    def process_variables(data, variables):
        """پردازش متغیرهای مالی"""
        variable_data = {}
        for variable in variables:
            if variable in data.columns:
                value = safe_convert_to_number(data[variable].iloc[0])
                variable_data[variable] = value
                print(f"{variable}: {value:,.0f}")
            else:
                variable_data[variable] = 0

        # محاسبه بهای تمام شده اگر صفر باشد
        if variable_data["بهای تمام شده کالای فروش رفته"] == 0:
            cogs = calculate_cogs(
                variable_data["فروش"],
                variable_data["سود ناخالص"]
            )
            if cogs > 0:
                variable_data["بهای تمام شده کالای فروش رفته"] = cogs
                print(f"بهای تمام شده محاسبه شده: {cogs:,.0f}")

        return variable_data

    def create_charts(self, all_data: Dict, company_name: str):
        """ایجاد نمودارهای تحلیلی"""
        try:
            years = sorted(all_data['ratios'].keys())

            # تنظیمات نمودار
            plt.figure(figsize=(15, 10))

            # نمودار روند نسبت‌های مهم
            important_ratios = [
                "نسبت جاری",
                "حاشیه سود خالص",
                "بازده حقوق صاحبان سهام",
                "نسبت بدهی به دارایی"
            ]

            for ratio in important_ratios:
                values = [float(all_data['ratios'][year].get(ratio, 0)) for year in years]
                plt.plot(years, values, marker='o', linewidth=2, markersize=8, label=ratio)

            plt.title(f'روند نسبت‌های مالی کلیدی - {company_name}', fontsize=14, pad=20)
            plt.xlabel('سال', fontsize=12)
            plt.ylabel('درصد', fontsize=12)
            plt.grid(True, linestyle='--', alpha=0.7)
            plt.legend(loc='center left', bbox_to_anchor=(1, 0.5))

            # ذخیره نمودار
            chart_path = self.output_dir / f"{company_name}_Financial_Ratios_{self.current_time}.png"
            plt.savefig(chart_path, bbox_inches='tight', dpi=300)
            plt.close()

            print(f"\nنمودار در مسیر زیر ذخیره شد:\n{chart_path}")
            return chart_path

        except Exception as e:
            print(f"خطا در ایجاد نمودار: {str(e)}")
            return None

    def analyze_company(self, company_name: str):
        """تحلیل کامل یک شرکت"""
        try:
            all_data = {'variables': {}, 'ratios': {}}

            # یافتن فایل‌های شرکت
            files = sorted(self.input_folder.glob(f'*{company_name}*.xlsx'))

            if not files:
                print(f"هیچ فایلی برای شرکت {company_name} یافت نشد!")
                return None

            # پردازش هر فایل
            for file_path in files:
                try:
                    year = re.search(r'\d{4}', file_path.stem)
                    if year:
                        year = year.group()
                        variables, ratios = self.process_file(file_path)
                        if variables and ratios:
                            all_data['variables'][year] = variables
                            all_data['ratios'][year] = ratios

                except Exception as e:
                    print(f"خطا در پردازش سال {file_path.stem}: {str(e)}")
                    continue

            if all_data['variables']:
                # ایجاد گزارش‌ها
                self.create_excel_report(all_data, company_name)
                self.create_charts(all_data, company_name)
                return all_data

            return None

        except Exception as e:
            print(f"خطا در تحلیل شرکت {company_name}")
            print(f"علت خطا: {str(e)}")
            return None


def combine_financial_data(folder_path, output_file):
    """تابع اصلی برای ترکیب داده‌های مالی"""
    ratios_df = pd.DataFrame(columns=["سال"] + ratios)
    variables_df = pd.DataFrame(columns=["سال"] + variables)

    files = [os.path.join(folder_path, file) for file in os.listdir(folder_path)
             if file.endswith((".xls", ".xlsx"))]

    for file in files:
        try:
            year = os.path.splitext(os.path.basename(file))[0].split("_")[-1]
            print(f"\nدر حال پردازش فایل: {os.path.basename(file)}")

            engine = 'openpyxl' if file.endswith('.xlsx') else 'xlrd'
            data = pd.read_excel(file, engine=engine)
            data.columns = data.columns.str.strip()

            # پردازش متغیرها
            variable_data = {}
            for variable in variables:
                if variable in data.columns:
                    value = convert_to_number(data[variable].iloc[0])
                    variable_data[variable] = value
                    print(f"{variable}: {value:,.0f}")
                else:
                    variable_data[variable] = 0

            ratio_data = {ratio: data[ratio].iloc[0] if ratio in data.columns else None for ratio in ratios}
            variable_data = {variable: data[variable].iloc[0] if variable in data.columns else None for variable in variables}
            ratio_data = calculate_ratios(variable_data)

            # محاسبه نسبت‌ها
            ratio_data["نسبت جاری"] = self.safe_divide(variable_data["دارایی جاری"], variable_data["بدهی جاری"])
            ratio_data["نسبت آنی"] = self.safe_divide(
                variable_data["دارایی جاری"] - variable_data["موجودی کالا"],
                variable_data["بدهی جاری"]
            )

            print("\nنسبت‌های مالی محاسبه شده:")
            for ratio_name, ratio_value in ratio_data.items():
                print(f"{ratio_name}: {ratio_value:.2f}")

            # اضافه کردن به دیتافریم‌ها
            ratios_df = pd.concat([ratios_df, pd.DataFrame([{"سال": year, **ratio_data}])], ignore_index=True)
            variables_df = pd.concat([variables_df, pd.DataFrame([{"سال": year, **variable_data}])], ignore_index=True)


        except Exception as e:

            print(f"خطا در پردازش فایل {file}: {str(e)}")

            continue

        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:

            ratios_df.to_excel(writer, sheet_name="نسبت‌ها", index=False)

            variables_df.to_excel(writer, sheet_name="متغیرها", index=False)

        print(f"فایل ترکیب‌شده با موفقیت ذخیره شد: {output_file}")


def main():
    # دریافت مسیر پوشه
    input_folder = input("لطفا مسیر پوشه حاوی فایل‌های اکسل را وارد کنید: ").strip()

    if not os.path.exists(input_folder):
        print("مسیر وارد شده معتبر نیست!")
        return

    # ایجاد آنالایزر
    analyzer = FinancialAnalyzer(input_folder)

    # تحلیل شرکت
    company_name = input("نام شرکت را وارد کنید: ").strip()
    results = analyzer.analyze_company(company_name)

    if results:
        print("\nتحلیل با موفقیت انجام شد!")
        print(f"گزارش‌ها در پوشه {analyzer.output_dir} ذخیره شده‌اند.")
    else:
        print("\nخطا در انجام تحلیل!")


if __name__ == "__main__":
    main()


