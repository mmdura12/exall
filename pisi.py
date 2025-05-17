# Standard library imports
import os
import re
import warnings
from datetime import datetime
from decimal import Decimal, getcontext, ROUND_HALF_UP
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Union

# Third party imports
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# تنظیمات اولیه
warnings.filterwarnings('ignore')
getcontext().prec = 28
plt.style.use('default')  # استفاده از استایل پیش‌فرض به جای seaborn
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
        self.input_folder = Path(input_folder)
        self.current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.output_dir = self.input_folder / "Financial_Reports"
        self.output_dir.mkdir(exist_ok=True)

        # تعریف عبارات جستجو برای متغیرها
        self.search_terms = {
            "وجه نقد": ["موجودی نقد", "وجه نقد", "موجودی نقد و بانک",
                        "نقد و معادل نقد", "موجودی نقد و معادل نقد"],
            "حساب دریافتنی": ["حسابهای دریافتنی تجاری", "دریافتنی‌های تجاری",
                              "حساب‌های دریافتنی", "حسابها و اسناد دریافتنی تجاری"],
            "موجودی کالا": ["موجودی مواد و کالا", "موجودی کالا", "موجودی‌ها",
                            "موجودی مواد، کالا و قطعات", "موجودی مواد اولیه و کالا"],
            "دارایی جاری": ["جمع دارایی‌های جاری", "دارایی‌های جاری",
                            "جمع دارایی های جاری", "دارایی های جاری"],
            "کل دارایی ها": ["جمع دارایی‌ها", "جمع کل دارایی‌ها",
                             "دارایی‌ها", "جمع دارایی ها"],
            "بدهی جاری": ["جمع بدهی‌های جاری", "بدهی‌های جاری",
                          "جمع بدهی های جاری", "بدهی های جاری"],
            "کل بدهی ها": ["جمع بدهی‌ها", "جمع کل بدهی‌ها",
                           "بدهی‌ها", "جمع بدهی ها"],
            "حقوق صاحبان سهام": ["جمع حقوق صاحبان سهام", "حقوق صاحبان سهام",
                                 "جمع حقوق مالکانه", "حقوق مالکانه"],
            "فروش": ["درآمدهای عملیاتی", "فروش", "فروش خالص",
                     "درآمد عملیاتی", "جمع درآمدهای عملیاتی"],
            "بهای تمام شده کالای فروش رفته": ["بهای تمام‌شده درآمدهای عملیاتی",
                                              "بهای تمام شده کالای فروش رفته",
                                              "بهای تمام شده فروش"],
            "سود ناخالص": ["سود ناخالص", "سود (زیان) ناخالص",
                           "سود و زیان ناخالص"],
            "سود عملیاتی": ["سود عملیاتی", "سود (زیان) عملیاتی",
                            "سود و زیان عملیاتی"],
            "سود خالص": ["سود خالص", "سود (زیان) خالص",
                         "سود (زیان) خالص دوره"]
        }

    def clean_persian_text(self, text: str) -> str:
        """پاکسازی متن فارسی"""
        if not isinstance(text, str):
            return ''
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

            value = str(value).strip()

            # حذف کاراکترهای خاص و فرمت‌بندی
            value = value.replace(',', '').replace('٬', '')

            # تبدیل اعداد منفی در پرانتز
            if '(' in value and ')' in value:
                value = '-' + value.replace('(', '').replace(')', '')

            # تبدیل انواع خط تیره به منفی
            value = value.replace('−', '-').replace('–', '-')

            # تبدیل اعداد فارسی به انگلیسی
            persian_nums = '۰۱۲۳۴۵۶۷۸۹'
            english_nums = '0123456789'
            for persian, english in zip(persian_nums, english_nums):
                value = value.replace(persian, english)

            # حذف همه کاراکترها بجز اعداد، نقطه و منفی
            value = ''.join(c for c in value if c.isdigit() or c in '.-')

            if not value or value in ('-', '--'):
                return Decimal('0')

            return Decimal(value)

        except Exception as e:
            print(f"خطا در تبدیل '{value}' به عدد: {e}")
            return Decimal('0')

    def find_value_in_df(self, df: pd.DataFrame, search_terms: List[str]) -> Decimal:
        """یافتن مقدار در دیتافریم"""
        try:
            for term in search_terms:
                cleaned_term = self.clean_persian_text(term)
                for idx, row in df.iterrows():
                    found_in_row = False
                    # اول در ستون اول (معمولاً نام حساب) جستجو می‌کنیم
                    if cleaned_term in self.clean_persian_text(str(row.iloc[0])):
                        found_in_row = True
                    if found_in_row:
                        # در بقیه ستون‌ها دنبال عدد می‌گردیم
                        for i in range(1, len(row)):
                            try:
                                value = str(row.iloc[i]).strip()
                                if value and value not in ['-', '--']:
                                    number = self.convert_to_number(value)
                                    if number != 0:  # اگر عدد معتبر پیدا شد
                                        return number
                            except:
                                continue
            return Decimal('0')
        except Exception as e:
            print(f"خطا در جستجوی {search_terms}: {e}")
            return Decimal('0')

    @staticmethod
    def safe_divide(numerator: Decimal, denominator: Decimal) -> Decimal:
        """تقسیم ایمن دو عدد"""
        try:
            if numerator is None or denominator is None:
                return Decimal('0')

            numerator = Decimal(str(numerator))
            denominator = Decimal(str(denominator))

            if denominator == 0:
                return Decimal('0')

            result = numerator / denominator
            return result.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)

        except Exception as e:
            print(f"خطا در تقسیم {numerator} بر {denominator}: {str(e)}")
            return Decimal('0')

        def calculate_financial_ratios(self, variables: Dict[str, Decimal]) -> Dict[str, Decimal]:
            """محاسبه نسبت‌های مالی"""
            try:
                ratios = {}

                # محاسبه بهای تمام شده اگر صفر است
                if variables["بهای تمام شده کالای فروش رفته"] == 0 and variables["فروش"] > 0 and variables[
                    "سود ناخالص"] > 0:
                    variables["بهای تمام شده کالای فروش رفته"] = variables["فروش"] - variables["سود ناخالص"]

                # نسبت‌های نقدینگی
                if variables["بدهی جاری"] > 0:
                    ratios["نسبت جاری"] = self.safe_divide(variables["دارایی جاری"], variables["بدهی جاری"])
                    ratios["نسبت آنی"] = self.safe_divide(variables["دارایی جاری"] - variables["موجودی کالا"],
                                                          variables["بدهی جاری"])
                    ratios["نسبت وجه نقد"] = self.safe_divide(variables["وجه نقد"], variables["بدهی جاری"])

                # نسبت‌های سودآوری
                if variables["فروش"] > 0:
                    ratios["حاشیه سود ناخالص"] = self.safe_divide(variables["سود ناخالص"], variables["فروش"]) * 100
                    ratios["حاشیه سود عملیاتی"] = self.safe_divide(variables["سود عملیاتی"], variables["فروش"]) * 100
                    ratios["حاشیه سود خالص"] = self.safe_divide(variables["سود خالص"], variables["فروش"]) * 100

                # نسبت‌های فعالیت
                if variables["فروش"] > 0 and variables["حساب دریافتنی"] > 0:
                    ratios["دوره وصول مطالبات"] = self.safe_divide(variables["حساب دریافتنی"] * 365, variables["فروش"])
                    ratios["گردش حساب دریافتنی"] = self.safe_divide(variables["فروش"], variables["حساب دریافتنی"])

                if variables["موجودی کالا"] > 0 and variables["بهای تمام شده کالای فروش رفته"] > 0:
                    ratios["گردش موجودی کالا"] = self.safe_divide(variables["بهای تمام شده کالای فروش رفته"],
                                                                  variables["موجودی کالا"])

                # نسبت‌های اهرمی و بازده
                if variables["کل دارایی ها"] > 0:
                    ratios["بازده دارایی ها"] = self.safe_divide(variables["سود خالص"], variables["کل دارایی ها"]) * 100
                    ratios["نسبت بدهی به دارایی"] = self.safe_divide(variables["کل بدهی ها"],
                                                                     variables["کل دارایی ها"]) * 100

                if variables["حقوق صاحبان سهام"] > 0:
                    ratios["بازده حقوق صاحبان سهام"] = self.safe_divide(variables["سود خالص"],
                                                                        variables["حقوق صاحبان سهام"]) * 100

                return ratios
            except Exception as e:
                print(f"خطا در محاسبه نسبت‌های مالی: {e}")
                return {}

        def convert_to_number(self, value: str) -> Decimal:
            """تبدیل متن به عدد با پشتیبانی کامل از فرمت‌های مختلف"""
            try:
                if pd.isna(value) or not str(value).strip():
                    return Decimal('0')

                value = str(value)

                # تمیز کردن متن
                value = value.strip()
                value = value.replace(',', '')  # حذف کاما
                value = value.replace('٬', '')  # حذف کاما فارسی

                # تبدیل اعداد منفی
                if '(' in value and ')' in value:  # اعداد منفی در پرانتز
                    value = value.replace('(', '-').replace(')', '')
                value = value.replace('−', '-').replace('–', '-')  # خط تیره‌های مختلف

                # تبدیل اعداد فارسی به انگلیسی
                persian_nums = '۰۱۲۳۴۵۶۷۸۹'
                english_nums = '0123456789'
                for persian, english in zip(persian_nums, english_nums):
                    value = value.replace(persian, english)

                # حذف همه کاراکترها بجز اعداد، نقطه و منفی
                value = ''.join(c for c in value if c.isdigit() or c in '.-')

                # اگر فقط علامت منفی یا خالی بود
                if value in ('', '-', '--'):
                    return Decimal('0')

                return Decimal(value)

            except Exception as e:
                print(f"خطا در تبدیل مقدار {value} به عدد: {str(e)}")
                return Decimal('0')

    def process_file(self, file_path: Path) -> tuple:
        """پردازش فایل اکسل و استخراج داده‌ها"""
        try:
            print(f"\nدر حال پردازش فایل: {file_path.name}")
            df = pd.read_excel(file_path, header=None, dtype=str)

            variables = {}
            for var_name, search_terms in self.search_terms.items():
                value = self.find_value_in_df(df, search_terms)
                variables[var_name] = value
                print(f"{var_name}: {float(value):,.0f}")

            ratios = self.calculate_financial_ratios(variables)
            return variables, ratios

        except Exception as e:
            print(f"خطا در پردازش فایل {file_path.name}: {e}")
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
        """ایجاد نمودارهای تحلیلی با جزئیات بیشتر"""
        try:
            years = sorted(all_data['ratios'].keys())

            # ایجاد نمودارهای جداگانه برای هر گروه نسبت
            ratio_groups = {
                'نسبت‌های نقدینگی': ['نسبت جاری', 'نسبت آنی', 'نسبت وجه نقد'],
                'نسبت‌های سودآوری': ['حاشیه سود ناخالص', 'حاشیه سود عملیاتی', 'حاشیه سود خالص'],
                'نسبت‌های بازده': ['بازده دارایی ها', 'بازده حقوق صاحبان سهام'],
                'نسبت‌های فعالیت': ['گردش حساب دریافتنی', 'گردش موجودی کالا', 'دوره وصول مطالبات']
            }

            for group_name, ratios in ratio_groups.items():
                plt.figure(figsize=(12, 6))
                for ratio in ratios:
                    values = [float(all_data['ratios'][year].get(ratio, 0)) for year in years]
                    plt.plot(years, values, marker='o', linewidth=2, markersize=8, label=ratio)

                plt.title(f'{group_name} - {company_name}', fontsize=14, pad=20)
                plt.xlabel('سال', fontsize=12)
                plt.ylabel('مقدار', fontsize=12)
                plt.grid(True, linestyle='--', alpha=0.7)
                plt.legend(loc='center left', bbox_to_anchor=(1, 0.5))

                # تنظیمات ظاهری
                plt.xticks(rotation=45)
                plt.tight_layout()

                # ذخیره نمودار
                chart_path = self.output_dir / f"{company_name}_{group_name}_{self.current_time}.png"
                plt.savefig(chart_path, bbox_inches='tight', dpi=300)
                plt.close()

            print("\nنمودارها با موفقیت ایجاد شدند.")
            return True

        except Exception as e:
            print(f"خطا در ایجاد نمودار: {str(e)}")
            return False

    def analyze_company(self, company_name: str) -> Optional[Dict]:
        """تحلیل کامل یک شرکت"""
        try:
            results = {'variables': {}, 'ratios': {}}

            # یافتن فایل‌های شرکت
            files = sorted(self.input_folder.glob(f'*{company_name}*.xlsx'))

            if not files:
                print(f"هیچ فایلی برای شرکت {company_name} یافت نشد!")
                return None

            print(f"\nتعداد {len(files)} فایل برای پردازش یافت شد.")

            # پردازش هر فایل
            for file_path in files:
                try:
                    year = re.search(r'\d{4}', file_path.stem)
                    if year:
                        year = year.group()
                        print(f"\nپردازش فایل سال {year}...")

                        variables, ratios = self.process_file(file_path)

                        if variables and ratios:
                            results['variables'][year] = variables
                            results['ratios'][year] = ratios
                            print(f"پردازش سال {year} با موفقیت انجام شد.")

                except Exception as e:
                    print(f"خطا در پردازش فایل {file_path.name}: {str(e)}")
                    continue

            if results['variables']:
                # ایجاد گزارش‌ها
                print("\nدر حال ایجاد گزارش‌ها...")
                self.create_excel_report(results, company_name)
                self.create_charts(results, company_name)
                print("\nتحلیل با موفقیت انجام شد!")
                return results
            else:
                print("\nهیچ داده‌ای برای تحلیل یافت نشد!")
                return None

        except Exception as e:
            print(f"\nخطا در تحلیل شرکت {company_name}")
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
    """تابع اصلی برنامه برای تحلیل و مقایسه مالی شرکت‌ها"""
    try:
        # نمایش اطلاعات شروع برنامه
        current_time = datetime.now()
        print("\n" + "=" * 50)
        print("=== سیستم تحلیل و مقایسه مالی شرکت‌ها ===")
        print("=" * 50)
        print(f"تاریخ و زمان: {current_time.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"کاربر: {os.getenv('USERNAME', 'mmdura12')}")
        print("-" * 50 + "\n")

        # دریافت و اعتبارسنجی مسیر پوشه
        while True:
            input_folder = input("لطفا مسیر پوشه حاوی فایل‌های اکسل را وارد کنید: ").strip()

            if not input_folder:
                print("خطا: مسیر نمی‌تواند خالی باشد!")
                continue

            if not os.path.exists(input_folder):
                print("خطا: مسیر وارد شده وجود ندارد!")
                if input("آیا می‌خواهید دوباره تلاش کنید؟ (بله/خیر) ").lower() != 'بله':
                    return
                continue

            excel_files = [f for f in os.listdir(input_folder) if f.endswith(('.xlsx', '.xls'))]
            if not excel_files:
                print("خطا: هیچ فایل اکسلی در این مسیر یافت نشد!")
                if input("آیا می‌خواهید مسیر دیگری وارد کنید؟ (بله/خیر) ").lower() != 'بله':
                    return
                continue

            break

        # ایجاد آنالایزر
        print("\nدر حال آماده‌سازی آنالایزر...")
        analyzer = FinancialAnalyzer(input_folder)
        print("آنالایزر با موفقیت ایجاد شد.")

        # لیست برای ذخیره نتایج همه شرکت‌ها
        all_companies_results = {}

        while True:
            # دریافت نام شرکت‌ها
            company_names = []
            print("\nلطفا نام شرکت‌ها را وارد کنید (برای پایان 'پایان' را وارد کنید):")
            while True:
                company_name = input(f"نام شرکت {len(company_names) + 1}: ").strip()
                if company_name.lower() == 'پایان':
                    break
                if not company_name:
                    print("نام شرکت نمی‌تواند خالی باشد!")
                    continue

                # بررسی وجود فایل‌های مربوط به شرکت
                matching_files = [f for f in excel_files if company_name in f]
                if not matching_files:
                    print(f"هیچ فایلی برای شرکت '{company_name}' یافت نشد!")
                    continue

                company_names.append(company_name)
                print(f"تعداد {len(matching_files)} فایل برای شرکت {company_name} یافت شد:")
                for file in matching_files:
                    print(f"- {file}")

                if len(company_names) >= 5:  # حداکثر 5 شرکت
                    print("حداکثر تعداد شرکت‌ها (5) وارد شده است.")
                    break

            if not company_names:
                print("هیچ شرکتی برای تحلیل وارد نشده است!")
                if input("آیا می‌خواهید دوباره تلاش کنید؟ (بله/خیر) ").lower() != 'بله':
                    break
                continue

            print("\n=== شروع تحلیل شرکت‌ها ===")

            # تحلیل هر شرکت
            for company_name in company_names:
                print(f"\nتحلیل شرکت {company_name}...")
                results = analyzer.analyze_company(company_name)
                if results:
                    all_companies_results[company_name] = results
                    print(f"تحلیل شرکت {company_name} با موفقیت انجام شد.")
                else:
                    print(f"خطا در تحلیل شرکت {company_name}")

            # مقایسه نتایج اگر بیش از یک شرکت وجود دارد
            if len(all_companies_results) > 1:
                print("\n=== مقایسه شرکت‌ها ===")

                # ایجاد فایل مقایسه‌ای
                comparison_file = os.path.join(analyzer.output_dir,
                                               f"Companies_Comparison_{current_time.strftime('%Y%m%d_%H%M%S')}.xlsx")

                with pd.ExcelWriter(comparison_file) as writer:
                    # مقایسه نسبت‌ها
                    comparison_ratios = {}
                    for company, data in all_companies_results.items():
                        latest_year = max(data['ratios'].keys())
                        comparison_ratios[company] = data['ratios'][latest_year]

                    pd.DataFrame(comparison_ratios).to_excel(writer, sheet_name='مقایسه نسبت‌ها')

                    # مقایسه متغیرها
                    comparison_vars = {}
                    for company, data in all_companies_results.items():
                        latest_year = max(data['variables'].keys())
                        comparison_vars[company] = data['variables'][latest_year]

                    pd.DataFrame(comparison_vars).to_excel(writer, sheet_name='مقایسه متغیرها')

                print(f"\nفایل مقایسه در مسیر زیر ذخیره شد:\n{comparison_file}")

            # پرسش برای ادامه
            if input("\nآیا می‌خواهید شرکت‌های دیگری را تحلیل کنید؟ (بله/خیر) ").lower() != 'بله':
                break

        print("\nتحلیل با موفقیت به پایان رسید!")
        print(f"همه گزارش‌ها در پوشه {analyzer.output_dir} ذخیره شده‌اند.")

    except KeyboardInterrupt:
        print("\n\nبرنامه توسط کاربر متوقف شد.")
    except Exception as e:
        print(f"\nخطای غیرمنتظره: {str(e)}")
        print("لطفا با پشتیبانی تماس بگیرید.")
    finally:
        print("\n" + "=" * 50)
        print("پایان برنامه")
        print("=" * 50)


if __name__ == "__main__":
    main()