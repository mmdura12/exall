# ابتدا همه import های مورد نیاز
import os
import re
import warnings
from datetime import datetime
from decimal import (Decimal, getcontext, ROUND_HALF_UP,
                    InvalidOperation, DivisionByZero, localcontext)
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Union
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
# تنظیمات اولیه
warnings.filterwarnings('ignore')
getcontext().prec = 28

class FinancialAnalyzer:
    def __init__(self, input_folder: str):
        """مقداردهی اولیه کلاس تحلیلگر مالی"""
        self.input_folder = Path(input_folder)
        self.current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.output_dir = self.input_folder / "Financial_Reports"
        self.output_dir.mkdir(exist_ok=True)

        # تعریف عبارات جستجوی دقیق برای متغیرها
        self.search_terms = {
            "وجه نقد": [
                "موجودی نقد و بانک",
                "نقد و معادل نقد",
                "موجودی نقد و معادل نقد",
                "موجودی نقد",
                "وجه نقد"
            ],
            "حساب دریافتنی": [
                "حسابهای دریافتنی تجاری",
                "دریافتنی های تجاری",
                "حسابها و اسناد دریافتنی تجاری",
                "حساب‌های دریافتنی تجاری",
                "دریافتنی‌های تجاری",
                "حساب های دریافتنی"
            ],
            "موجودی کالا": [
                "موجودی مواد و کالا",
                "موجودی کالا",
                "موجودی مواد، کالا و قطعات",
                "موجودی مواد اولیه و کالا",
                "موجودی‌ها"
            ],
            "دارایی جاری": [
                "جمع دارایی‌های جاری",
                "دارایی‌های جاری",
                "جمع دارایی های جاری",
                "دارایی های جاری",
                "جمع داراییهای جاری"
            ],
            "کل دارایی ها": [
                "جمع کل دارایی‌ها",
                "جمع دارایی‌ها",
                "جمع دارایی ها",
                "دارایی‌ها",
                "کل دارایی‌ها"
            ],
            "بدهی جاری": [
                "جمع بدهی‌های جاری",
                "بدهی‌های جاری",
                "جمع بدهی های جاری",
                "بدهی های جاری",
                "جمع بدهیهای جاری"
            ],
            "کل بدهی ها": [
                "جمع کل بدهی‌ها",
                "جمع بدهی‌ها",
                "جمع بدهی ها",
                "بدهی‌ها",
                "کل بدهی‌ها"
            ],
            "حقوق صاحبان سهام": [
                "جمع حقوق صاحبان سهام",
                "حقوق صاحبان سهام",
                "جمع حقوق مالکانه",
                "حقوق مالکانه",
                "جمع حقوق سهامداران"
            ],
            "فروش": [
                "فروش و درآمد ارائه خدمات",
                "درآمدهای عملیاتی",
                "فروش خالص",
                "درآمد عملیاتی",
                "جمع درآمدهای عملیاتی",
                "فروش"
            ],
            "بهای تمام شده کالای فروش رفته": [
                "بهای تمام‌شده درآمدهای عملیاتی",
                "بهای تمام شده کالای فروش رفته",
                "بهای تمام شده فروش",
                "بهای تمام‌شده کالای فروش رفته"
            ],
            "سود ناخالص": [
                "سود (زیان) ناخالص",
                "سود و زیان ناخالص",
                "سود ناخالص",
                "سود (زیان) ناخالص فروش"
            ],
            "سود عملیاتی": [
                "سود (زیان) عملیاتی",
                "سود و زیان عملیاتی",
                "سود عملیاتی"
            ],
            "سود خالص": [
                "سود (زیان) خالص",
                "سود (زیان) خالص دوره",
                "سود خالص",
                "سود خالص دوره"
            ]
        }

        # تنظیمات نمودار
        plt.style.use('default')
        plt.rcParams.update({
            'font.family': 'B Nazanin',
            'figure.figsize': (12, 8),
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

    @staticmethod
    def safe_divide(numerator: Decimal, denominator: Decimal) -> Decimal:
        """
        تقسیم ایمن دو عدد با دقت بالا و کنترل خطای پیشرفته

        Args:
            numerator (Decimal): صورت کسر
            denominator (Decimal): مخرج کسر

        Returns:
            Decimal: حاصل تقسیم با دقت دو رقم اعشار
        """
        try:
            # تبدیل ورودی‌ها به Decimal و اعتبارسنجی
            if numerator is None or denominator is None:
                return Decimal('0.00')

            # تبدیل به Decimal با حفظ دقت
            try:
                num = Decimal(str(numerator)).normalize()
                den = Decimal(str(denominator)).normalize()
            except (ValueError, TypeError, InvalidOperation):
                return Decimal('0.00')

            # کنترل صفر بودن مخرج
            if den.is_zero():
                return Decimal('0.00')

            # کنترل اعداد بسیار کوچک
            if abs(den) < Decimal('1E-10'):
                return Decimal('0.00')

            # انجام تقسیم با تنظیمات دقیق
            with localcontext() as ctx:
                ctx.prec = 10
                ctx.rounding = ROUND_HALF_UP
                result = (num / den).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)

                # کنترل محدوده مجاز
                if abs(result) > Decimal('1E+15'):
                    return Decimal('0.00')

                return result

        except Exception:
            return Decimal('0.00')

    def clean_persian_text(self, text: str) -> str:
        """پاکسازی و استانداردسازی متن فارسی"""
        if not isinstance(text, str):
            return ''

        # حذف فاصله‌های اضافی و یکسان‌سازی کاراکترها
        text = re.sub(r'\s+', ' ', text)

        # تبدیل کاراکترهای عربی/فارسی به فرم استاندارد
        replacements = {
            'ي': 'ی', 'ك': 'ک',
            '٠': '۰', '١': '۱', '٢': '۲', '٣': '۳', '٤': '۴',
            '٥': '۵', '٦': '۶', '٧': '۷', '٨': '۸', '٩': '۹',
            'ة': 'ه', 'آ': 'ا'
        }
        for arab, pers in replacements.items():
            text = text.replace(arab, pers)

        return text.strip()

    def convert_to_number(self, value: str) -> Decimal:
        """تبدیل متن به عدد با پشتیبانی از فرمت‌های مختلف"""
        try:
            if pd.isna(value) or not str(value).strip():
                return Decimal('0')

            value = str(value).strip()

            # پاکسازی متن از کاراکترهای اضافی
            value = value.replace(',', '').replace('٬', '')
            value = value.replace('−', '-').replace('–', '-')

            # تشخیص اعداد منفی در پرانتز
            if '(' in value and ')' in value:
                value = '-' + value.replace('(', '').replace(')', '')

            # تبدیل اعداد فارسی به انگلیسی
            persian_nums = '۰۱۲۳۴۵۶۷۸۹'
            english_nums = '0123456789'
            for persian, english in zip(persian_nums, english_nums):
                value = value.replace(persian, english)

            # حذف همه کاراکترها بجز اعداد، نقطه و منفی
            value = ''.join(c for c in value if c.isdigit() or c in '.-')

            # بررسی اعتبار عدد
            if not value or value in ('-', '--'):
                return Decimal('0')

            return Decimal(value)

        except Exception as e:
            print(f"خطا در تبدیل '{value}' به عدد: {e}")
            return Decimal('0')

    def find_value_in_df(self, df: pd.DataFrame, search_terms: List[str]) -> Decimal:
        """یافتن مقدار در دیتافریم با دقت بالا"""
        try:
            max_value = None
            for term in search_terms:
                term = self.clean_persian_text(term)
                for idx, row in df.iterrows():
                    for col in df.columns:
                        cell_value = self.clean_persian_text(str(row[col]))
                        if term in cell_value:
                            # جستجو در همان سطر برای یافتن عدد معتبر
                            row_values = [str(row[c]) for c in df.columns]
                            for value in row_values:
                                try:
                                    number = self.convert_to_number(value)
                                    if number != 0:  # اگر عدد معتبر یافت شد
                                        if max_value is None or abs(number) > abs(max_value):
                                            max_value = number
                                except:
                                    continue

            return max_value if max_value is not None else Decimal('0')

        except Exception as e:
            print(f"خطا در جستجوی مقادیر برای {search_terms}: {e}")
            return Decimal('0')

    def calculate_financial_ratios(self, variables: Dict[str, Decimal]) -> Dict[str, Decimal]:
        """محاسبه نسبت‌های مالی با دقت بالا"""
        try:
            ratios = {}

            # محاسبه بهای تمام شده اگر صفر است
            if variables["بهای تمام شده کالای فروش رفته"] == 0 and variables["فروش"] > 0 and variables[
                "سود ناخالص"] > 0:
                variables["بهای تمام شده کالای فروش رفته"] = variables["فروش"] - variables["سود ناخالص"]
                print(f"بهای تمام شده محاسبه شده: {float(variables['بهای تمام شده کالای فروش رفته']):,.0f}")

            # نسبت‌های نقدینگی
            if variables["بدهی جاری"] > 0:
                ratios["نسبت جاری"] = self.safe_divide(variables["دارایی جاری"], variables["بدهی جاری"])
                ratios["نسبت آنی"] = self.safe_divide(
                    variables["دارایی جاری"] - variables["موجودی کالا"],
                    variables["بدهی جاری"]
                )
                ratios["نسبت وجه نقد"] = self.safe_divide(variables["وجه نقد"], variables["بدهی جاری"])

            # نسبت‌های سودآوری
            if variables["فروش"] > 0:
                ratios["حاشیه سود ناخالص"] = self.safe_divide(variables["سود ناخالص"], variables["فروش"]) * Decimal(
                    '100')
                ratios["حاشیه سود عملیاتی"] = self.safe_divide(variables["سود عملیاتی"], variables["فروش"]) * Decimal(
                    '100')
                ratios["حاشیه سود خالص"] = self.safe_divide(variables["سود خالص"], variables["فروش"]) * Decimal('100')

            # نسبت‌های فعالیت
            if variables["فروش"] > 0 and variables["حساب دریافتنی"] > 0:
                ratios["دوره وصول مطالبات"] = self.safe_divide(
                    variables["حساب دریافتنی"] * Decimal('365'),
                    variables["فروش"]
                )
                ratios["گردش حساب دریافتنی"] = self.safe_divide(variables["فروش"], variables["حساب دریافتنی"])

            if variables["موجودی کالا"] > 0 and variables["بهای تمام شده کالای فروش رفته"] > 0:
                ratios["گردش موجودی کالا"] = self.safe_divide(
                    variables["بهای تمام شده کالای فروش رفته"],
                    variables["موجودی کالا"]
                )

            # نسبت‌های بازده و اهرمی
            if variables["کل دارایی ها"] > 0:
                ratios["بازده دارایی ها"] = self.safe_divide(variables["سود خالص"],
                                                             variables["کل دارایی ها"]) * Decimal('100')
                ratios["نسبت بدهی به دارایی"] = self.safe_divide(variables["کل بدهی ها"],
                                                                 variables["کل دارایی ها"]) * Decimal('100')

            if variables["حقوق صاحبان سهام"] > 0:
                ratios["بازده حقوق صاحبان سهام"] = self.safe_divide(
                    variables["سود خالص"],
                    variables["حقوق صاحبان سهام"]
                ) * Decimal('100')

            return ratios

        except Exception as e:
            print(f"خطا در محاسبه نسبت‌های مالی: {e}")
            return {}



    def process_file(self, file_path: Path) -> Tuple[Dict[str, Decimal], Dict[str, Decimal]]:
        """پردازش فایل اکسل و استخراج داده‌ها با دقت بالا"""
        try:
            print(f"\nدر حال پردازش فایل: {file_path.name}")

            # خواندن فایل اکسل با تنظیمات مناسب
            df = pd.read_excel(
                file_path,
                header=None,
                dtype=str,
                na_filter=False,
                engine='openpyxl'
            )

            variables = {}
            print("\nاستخراج متغیرهای مالی:")

            # استخراج و پردازش متغیرها
            for var_name, search_terms in self.search_terms.items():
                value = self.find_value_in_df(df, search_terms)
                variables[var_name] = value
                print(f"{var_name}: {float(value):,.0f}")

            # محاسبه بهای تمام شده اگر صفر باشد
            if variables["بهای تمام شده کالای فروش رفته"] == 0 and variables["فروش"] > 0 and variables[
                "سود ناخالص"] > 0:
                variables["بهای تمام شده کالای فروش رفته"] = variables["فروش"] - variables["سود ناخالص"]
                print(f"\nبهای تمام شده محاسبه شده: {float(variables['بهای تمام شده کالای فروش رفته']):,.0f}")

            # محاسبه نسبت‌های مالی
            print("\nمحاسبه نسبت‌های مالی:")
            ratios = self.calculate_financial_ratios(variables)

            # نمایش نسبت‌های محاسبه شده
            for ratio_name, ratio_value in ratios.items():
                print(f"{ratio_name}: {float(ratio_value):.2f}")

            return variables, ratios

        except Exception as e:
            print(f"خطا در پردازش فایل {file_path.name}: {str(e)}")
            return {}, {}

    def create_excel_report(self, all_data: Dict, company_name: str) -> Optional[Path]:
        """ایجاد گزارش اکسل با فرمت‌بندی پیشرفته"""
        try:
            output_file = self.output_dir / f"{company_name}_Financial_Analysis_{self.current_time}.xlsx"

            with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                workbook = writer.book

                # تعریف فرمت‌های سلول
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

                percentage_format = workbook.add_format({
                    'num_format': '0.00%',
                    'border': 1,
                    'align': 'center',
                    'valign': 'vcenter',
                    'font_name': 'B Nazanin',
                    'font_size': 11
                })

                # تبدیل داده‌ها به دیتافریم
                variables_df = pd.DataFrame(all_data['variables']).T
                ratios_df = pd.DataFrame(all_data['ratios']).T

                # نوشتن متغیرها
                variables_sheet = writer.book.add_worksheet('متغیرها')
                variables_sheet.right_to_left()

                # نوشتن هدر متغیرها
                for col, value in enumerate(['سال'] + list(variables_df.columns)):
                    variables_sheet.write(0, col, value, header_format)

                # نوشتن داده‌های متغیرها
                for row, (year, data) in enumerate(variables_df.iterrows()):
                    variables_sheet.write(row + 1, 0, year)
                    for col, value in enumerate(data):
                        variables_sheet.write(row + 1, col + 1, float(value), number_format)

                # نوشتن نسبت‌ها
                ratios_sheet = writer.book.add_worksheet('نسبت‌ها')
                ratios_sheet.right_to_left()

                # نوشتن هدر نسبت‌ها
                for col, value in enumerate(['سال'] + list(ratios_df.columns)):
                    ratios_sheet.write(0, col, value, header_format)

                # نوشتن داده‌های نسبت‌ها
                for row, (year, data) in enumerate(ratios_df.iterrows()):
                    ratios_sheet.write(row + 1, 0, year)
                    for col, value in enumerate(data):
                        if 'درصد' in ratios_df.columns[col] or '%' in ratios_df.columns[col]:
                            ratios_sheet.write(row + 1, col + 1, float(value) / 100, percentage_format)
                        else:
                            ratios_sheet.write(row + 1, col + 1, float(value), number_format)

                # تنظیم عرض ستون‌ها
                for worksheet in [variables_sheet, ratios_sheet]:
                    worksheet.set_column('A:Z', 15)

                # اضافه کردن فیلتر به هر دو شیت
                variables_sheet.autofilter(0, 0, len(variables_df), len(variables_df.columns))
                ratios_sheet.autofilter(0, 0, len(ratios_df), len(ratios_df.columns))

            print(f"\nگزارش اکسل در مسیر زیر ایجاد شد:\n{output_file}")
            return output_file

        except Exception as e:
            print(f"خطا در ایجاد گزارش اکسل: {str(e)}")
            return None

    def create_charts(self, all_data: Dict, company_name: str) -> bool:
        """ایجاد نمودارهای تحلیلی با جزئیات و کیفیت بالا"""
        try:
            years = sorted(all_data['ratios'].keys())

            # تعریف گروه‌های نسبت‌ها
            ratio_groups = {
                'نسبت‌های نقدینگی': {
                    'ratios': ['نسبت جاری', 'نسبت آنی', 'نسبت وجه نقد'],
                    'ylabel': 'مرتبه'
                },
                'نسبت‌های سودآوری': {
                    'ratios': ['حاشیه سود ناخالص', 'حاشیه سود عملیاتی', 'حاشیه سود خالص'],
                    'ylabel': 'درصد'
                },
                'نسبت‌های بازده': {
                    'ratios': ['بازده دارایی ها', 'بازده حقوق صاحبان سهام'],
                    'ylabel': 'درصد'
                },
                'نسبت‌های فعالیت': {
                    'ratios': ['گردش حساب دریافتنی', 'گردش موجودی کالا', 'دوره وصول مطالبات'],
                    'ylabel': 'مرتبه/روز'
                }
            }

            for group_name, group_info in ratio_groups.items():
                plt.figure(figsize=(12, 6))

                for ratio in group_info['ratios']:
                    values = [float(all_data['ratios'][year].get(ratio, 0)) for year in years]
                    plt.plot(years, values, marker='o', linewidth=2, markersize=8, label=ratio)

                    # اضافه کردن برچسب مقادیر روی نقاط
                    for x, y in zip(years, values):
                        plt.annotate(f'{y:.2f}',
                                     (x, y),
                                     textcoords="offset points",
                                     xytext=(0, 10),
                                     ha='center',
                                     fontsize=8)

                plt.title(f'{group_name} - {company_name}', fontsize=14, pad=20)
                plt.xlabel('سال', fontsize=12)
                plt.ylabel(group_info['ylabel'], fontsize=12)
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

            try:
                # ساختار نتایج
                results = {
                    'variables': {},
                    'ratios': {},
                    'metadata': {
                        'analysis_time': datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S'),
                        'company_name': company_name,
                        'analyst': os.getenv('USERNAME', 'mmdura12')
                    }
                }

                # یافتن و فیلتر کردن فایل‌های اکسل
                try:
                    files = sorted([
                        f for f in self.input_folder.glob(f'*{company_name}*.xlsx')
                        if not f.name.startswith('~$')
                    ])
                except Exception as e:
                    print(f"\nخطا در جستجوی فایل‌ها: {str(e)}")
                    return None

                if not files:
                    print(f"\nهیچ فایلی برای شرکت {company_name} یافت نشد!")
                    return None

                print(f"\nتعداد {len(files)} فایل برای پردازش یافت شد.")

                # پردازش هر فایل
                processed_files = 0
                for file_path in files:
                    try:
                        year_match = re.search(r'13([0-9]{2})', file_path.stem)
                        if not year_match:
                            print(f"نام فایل {file_path.name} فاقد سال معتبر است.")
                            continue

                        year = '14' + year_match.group(1)
                        print(f"\nپردازش فایل سال {year}...")

                        variables, ratios = self.process_file(file_path)

                        if not variables or not ratios:
                            print(f"خطا: داده‌های سال {year} خالی است!")
                            continue

                        if not self.validate_financial_data(variables):
                            print(f"هشدار: داده‌های سال {year} نامعتبر است!")
                            continue

                        results['variables'][year] = variables
                        results['ratios'][year] = ratios
                        processed_files += 1
                        print(f"پردازش سال {year} با موفقیت انجام شد.")

                    except Exception as e:
                        print(f"خطا در پردازش فایل {file_path.name}: {str(e)}")
                        continue

                if not results['variables']:
                    print("\nهیچ داده معتبری برای تحلیل یافت نشد!")
                    return None

                print(f"\nتعداد {processed_files} فایل با موفقیت پردازش شد.")

                try:
                    print("\nدر حال ایجاد گزارش‌ها...")
                    excel_file = self.create_excel_report(results, company_name)

                    if not excel_file:
                        print("خطا در ایجاد گزارش اکسل!")
                        return results

                    print("\nدر حال ایجاد نمودارها...")
                    if self.create_charts(results, company_name):
                        print("نمودارها با موفقیت ایجاد شدند.")
                    else:
                        print("هشدار: خطا در ایجاد نمودارها")

                    self.show_analysis_summary(results, company_name)

                    print(f"\nتحلیل شرکت {company_name} با موفقیت انجام شد.")
                    return results

                except Exception as e:
                    print(f"\nخطا در ایجاد گزارش‌ها و نمودارها: {str(e)}")
                    return results

            except Exception as e:
                print(f"\nخطای غیرمنتظره در تحلیل شرکت {company_name}")
                print(f"علت خطا: {str(e)}")
                return None

            finally:
                print(f"\nپایان پردازش شرکت {company_name}")
                print(f"زمان پایان (UTC): {datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')}")

        def analyze_company(self, company_name: str) -> Optional[Dict]:
            """تحلیل کامل اطلاعات مالی شرکت"""
            try:
                results = {
                    'variables': {},
                    'ratios': {},
                    'metadata': {
                        'analysis_time': datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S'),
                        'company_name': company_name,
                        'analyst': os.getenv('USERNAME', 'mmdura12')
                    }
                }

                # بقیه کد متد...

            except Exception as e:
                print(f"\nخطای غیرمنتظره در تحلیل شرکت {company_name}")
                print(f"علت خطا: {str(e)}")
                return None

        def validate_financial_data(self, variables: Dict[str, Decimal]) -> bool:
            """اعتبارسنجی داده‌های مالی"""
            try:
                key_variables = [
                    "دارایی جاری", "کل دارایی ها", "بدهی جاری",
                    "کل بدهی ها", "فروش", "سود ناخالص"
                ]

                # بقیه کد متد...

            except Exception as e:
                print(f"خطا در اعتبارسنجی داده‌ها: {str(e)}")
                return False

        def show_analysis_summary(self, results: Dict, company_name: str):
            """نمایش خلاصه تحلیلی از روند شرکت"""
            try:
                # کد متد...
                pass
            except Exception as e:
                print(f"خطا در نمایش خلاصه تحلیلی: {str(e)}")

        # خارج از کلاس - تابع main را اینجا قرار می‌دهیم
        def main():
            """تابع اصلی برنامه"""
            try:
                current_time = datetime.utcnow()
                formatted_time = current_time.strftime('%Y-%m-%d %H:%M:%S')

                print("\n" + "=" * 60)
                print("=== سیستم تحلیل و مقایسه مالی شرکت‌ها ===")
                print("=" * 60)
                print(f"Current Date and Time (UTC): {formatted_time}")
                print(f"Current User's Login: {os.getenv('USERNAME', 'mmdura12')}")
                print("-" * 60 + "\n")

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

                    excel_files = [f for f in os.listdir(input_folder)
                                   if f.endswith(('.xlsx', '.xls')) and not f.startswith('~$')]
                    if not excel_files:
                        print("خطا: هیچ فایل اکسل معتبری در این مسیر یافت نشد!")
                        if input("آیا می‌خواهید مسیر دیگری وارد کنید؟ (بله/خیر) ").lower() != 'بله':
                            return
                        continue

                    break

                print("\nدر حال آماده‌سازی آنالایزر...")
                analyzer = FinancialAnalyzer(input_folder)
                print("آنالایزر با موفقیت ایجاد شد.")

                # درخواست نام شرکت
                company_name = input("\nلطفا نام شرکت را وارد کنید: ").strip()
                if company_name:
                    results = analyzer.analyze_company(company_name)
                    if results:
                        print("تحلیل با موفقیت انجام شد.")

            except Exception as e:
                print(f"\nخطای غیرمنتظره: {str(e)}")
            finally:
                print("\nپایان برنامه")

        # در نهایت، اجرای برنامه
        if __name__ == "__main__":
            main()