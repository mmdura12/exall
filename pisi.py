import os
import pandas as pd
import numpy as np
import warnings
from datetime import datetime
from pathlib import Path
import glob

# غیرفعال کردن هشدارها
warnings.filterwarnings('ignore')


class FinancialAnalyzer:
    def __init__(self, base_folder):
        """مقداردهی اولیه"""
        self.base_folder = Path(base_folder)
        self.output_folder = self.base_folder / 'reports'
        self.output_folder.mkdir(exist_ok=True)

        # الگوهای جستجو برای یافتن مقادیر
        self.search_patterns = {
            'دارایی جاری': [
                'جمع دارایی‌های جاری', 'جمع داراییهای جاری', 'دارایی‌های جاری',
                'داراییهای جاری', 'دارایی های جاری', 'جمع دارایی های جاری',
                'دارایی جاری', 'داراییهای جاری', 'دارائیهای جاری', 'دارائی های جاری',
                'جمع داراییهای جاری', 'جمع دارائیهای جاری', 'مجموع داراییهای جاری',
                'جمع حسابهای دارایی جاری', 'کل دارایی های جاری', 'مجموع دارایی‌های جاری',
                'جمع کل دارایی‌های جاری', 'جمع کل داراییهای جاری', 'دارایی‌های جاری - جمع',
                'جمع داراییهای جاری و غیر جاری'
            ],
            'کل دارایی ها': [
                'جمع دارایی‌ها', 'جمع داراییها', 'جمع کل دارایی‌ها',
                'کل دارایی‌ها', 'دارایی ها', 'جمع دارایی ها',
                'جمع کل داراییها', 'جمع کل دارائیها', 'کل داراییها',
                'جمع داراییها', 'جمع دارائیها', 'مجموع کل داراییها',
                'جمع حسابهای دارایی', 'مجموع داراییها', 'دارایی های کل',
                'جمع کل حسابهای دارایی', 'جمع دارایی های شرکت',
                'مجموع دارایی‌ها', 'جمع داراییهای جاری و غیر جاری',
                'دارایی‌ها - جمع کل', 'جمع کل'
            ],
            'بدهی جاری': [
                'جمع بدهی‌های جاری', 'جمع بدهیهای جاری', 'بدهی‌های جاری',
                'بدهیهای جاری', 'بدهی های جاری', 'جمع بدهی های جاری',
                'بدهی جاری', 'بدهیهای جاری', 'بدهی‌های جاری', 'جمع کل بدهی های جاری',
                'مجموع بدهیهای جاری', 'کل بدهیهای جاری', 'جمع حسابهای بدهی جاری',
                'مجموع بدهی‌های جاری', 'جمع کل بدهی‌های جاری',
                'بدهی‌های جاری - جمع', 'جمع بدهی‌های کوتاه مدت'
            ],
            'کل بدهی ها': [
                'جمع بدهی‌ها', 'جمع بدهیها', 'جمع کل بدهی‌ها',
                'کل بدهی‌ها', 'بدهی ها', 'جمع بدهی ها',
                'جمع کل بدهیها', 'مجموع بدهیها', 'کل بدهیها',
                'جمع بدهی‌های جاری و غیرجاری', 'مجموع کل بدهی ها',
                'جمع حسابهای بدهی', 'بدهی های کل', 'جمع کل حسابهای بدهی',
                'مجموع بدهی‌ها', 'بدهی‌ها - جمع کل', 'جمع بدهی های شرکت'
            ],
            'فروش': [
                'درآمدهای عملیاتی', 'درآمد عملیاتی', 'فروش خالص',
                'فروش', 'درآمد حاصل از فروش', 'فروش و درآمد ارائه خدمات',
                'جمع درآمدهای عملیاتی', 'درآمد حاصل از فروش کالا',
                'فروش کالا و خدمات', 'درآمد عملیاتی - خالص',
                'فروش خالص و درآمد ارائه خدمات', 'درآمد خالص',
                'درآمد عملیاتی خالص', 'فروش و درآمد خالص',
                'درآمد حاصل از فروش و ارائه خدمات'
            ],
            'سود ناخالص': [
                'سود ناخالص', 'سود (زیان) ناخالص', 'سود و زیان ناخالص',
                'سود(زیان)ناخالص', 'سود/زیان ناخالص', 'سود یا زیان ناخالص',
                'سود (زیان) ناخالص فروش', 'سود ناخالص عملیاتی',
                'سود و زیان ناخالص عملیاتی', 'ناخالص سود و زیان'
            ],
            'سود عملیاتی': [
                'سود عملیاتی', 'سود (زیان) عملیاتی', 'سود و زیان عملیاتی',
                'سود(زیان)عملیاتی', 'سود/زیان عملیاتی', 'سود یا زیان عملیاتی',
                'سود و زیان خالص عملیات', 'سود خالص عملیاتی',
                'سود عملیاتی خالص', 'سود و زیان عملیاتی خالص'
            ],
            'سود خالص': [
                'سود خالص', 'سود (زیان) خالص', 'سود خالص دوره',
                'سود(زیان)خالص', 'سود/زیان خالص', 'سود یا زیان خالص',
                'سود خالص پس از کسر مالیات', 'سود و زیان خالص',
                'سود (زیان) خالص دوره', 'سود خالص سال',
                'سود و زیان خالص دوره', 'سود دوره خالص'
            ],
            'موجودی کالا': [
                'موجودی مواد و کالا', 'موجودی کالا', 'موجودی‌های مواد و کالا',
                'موجودی کالا و مواد', 'موجودیهای مواد و کالا',
                'موجودی مواد، کالا و قطعات', 'موجودی مواد',
                'موجودی کالای ساخته شده', 'موجودی کالای در جریان ساخت',
                'موجودی مواد اولیه', 'موجودی قطعات و ملزومات',
                'موجودی کالای در راه', 'موجودی مواد و کالای ساخته شده'
            ],
            'حساب های دریافتنی': [
                'حساب‌های دریافتنی تجاری', 'حسابهای دریافتنی', 'دریافتنی‌های تجاری',
                'حساب های دریافتنی تجاری', 'دریافتنی های تجاری',
                'حسابها و اسناد دریافتنی تجاری', 'حساب‌های دریافتنی',
                'حسابهای دریافتنی عملیاتی', 'دریافتنی های عملیاتی',
                'حساب و اسناد دریافتنی تجاری', 'مطالبات تجاری',
                'حسابهای دریافتنی - خالص', 'دریافتنی های تجاری و غیرتجاری'
            ]
        }
    def clean_number(self, value):
        """تمیز کردن و تبدیل مقادیر عددی"""
        try:
            if isinstance(value, (int, float)):
                return float(value)

            value = str(value).strip()
            value = value.replace(',', '').replace('٬', '')
            value = value.replace('(', '-').replace(')', '')
            value = value.replace('−', '-').replace('–', '-')

            # حذف کاراکترهای غیر عددی
            value = ''.join(c for c in value if c.isdigit() or c in '.-')

            if value and value not in ['-', '.']:
                return float(value)
            return 0
        except:
            return 0

    def find_value_in_df(self, df, keywords):
        """جستجوی دقیق‌تر مقدار در دیتافریم با حذف مقادیر صفر"""
        try:
            def normalize_text(text):
                """نرمال‌سازی پیشرفته متن فارسی"""
                text = str(text).strip()
                replacements = {
                    '‌': ' ',  # حذف نیم‌فاصله
                    'ي': 'ی',  # یکسان‌سازی ی
                    'ك': 'ک',  # یکسان‌سازی ک
                    '\u200c': ' ',  # حذف نیم‌فاصله یونیکد
                    '\n': ' ',  # حذف خط جدید
                    '\r': ' ',  # حذف کریج ریترن
                    '\t': ' ',  # حذف تب
                    '  ': ' ',  # حذف فاصله‌های اضافی
                }
                for old, new in replacements.items():
                    text = text.replace(old, new)
                return ' '.join(text.split())

            # پاکسازی و آماده‌سازی دیتافریم
            df = df.fillna('')
            df = df.replace('[\$,)]', '', regex=True)
            df = df.replace('[(]', '-', regex=True)
            df = df.astype(str)

            found_values = []  # لیست همه مقادیر یافت شده

            for keyword in keywords:
                normalized_keyword = normalize_text(keyword)

                for idx, row in df.iterrows():
                    row_values = []  # مقادیر یافت شده در یک سطر

                    # بررسی هر سلول در سطر
                    for col in df.columns:
                        try:
                            cell_value = normalize_text(row[col])

                            if normalized_keyword in cell_value:
                                col_idx = df.columns.get_loc(col)

                                # بررسی ستون‌های مجاور
                                nearby_values = []

                                # بررسی ستون‌های بعدی (تا 5 ستون)
                                for i in range(1, 6):
                                    if col_idx + i < len(df.columns):
                                        val = self.clean_number(row[df.columns[col_idx + i]])
                                        if val > 0:  # فقط مقادیر مثبت
                                            nearby_values.append(val)

                                # بررسی ستون‌های قبلی (تا 3 ستون)
                                for i in range(1, 4):
                                    if col_idx - i >= 0:
                                        val = self.clean_number(row[df.columns[col_idx - i]])
                                        if val > 0:  # فقط مقادیر مثبت
                                            nearby_values.append(val)

                                # بررسی ستون فعلی
                                val = self.clean_number(row[col])
                                if val > 0:
                                    nearby_values.append(val)

                                if nearby_values:
                                    row_values.extend(nearby_values)

                        except Exception as e:
                            continue

                    # اضافه کردن مقادیر معتبر به لیست نهایی
                    if row_values:
                        # حذف مقادیر تکراری
                        row_values = list(set(row_values))
                        found_values.extend(row_values)

            # انتخاب مقدار نهایی
            if found_values:
                # حذف مقادیر پرت
                found_values = sorted(found_values)  # مرتب‌سازی مقادیر
                if len(found_values) > 4:  # اگر بیش از 4 مقدار یافت شد
                    q1 = np.percentile(found_values, 25)
                    q3 = np.percentile(found_values, 75)
                    iqr = q3 - q1
                    lower_bound = q1 - (1.5 * iqr)
                    upper_bound = q3 + (1.5 * iqr)
                    valid_values = [x for x in found_values if lower_bound <= x <= upper_bound]
                else:
                    valid_values = found_values

                if valid_values:
                    # بررسی پراکندگی مقادیر
                    if max(valid_values) / min(valid_values) > 1000:
                        # اگر اختلاف زیاد است، میانه را برگردان
                        return np.median(valid_values)
                    else:
                        # در غیر این صورت، بیشترین مقدار را برگردان
                        return max(valid_values)

            return None  # اگر هیچ مقدار معتبری پیدا نشد

        except Exception as e:
            print(f"خطا در جستجوی مقدار: {str(e)}")
            return None

    def clean_number(self, value):
        """تمیز کردن و تبدیل مقادیر عددی با دقت بیشتر"""
        try:
            if isinstance(value, (int, float)):
                return float(value)

            if not value or value.isspace():
                return 0

            # حذف کاراکترهای خاص
            value = str(value).strip()
            value = value.replace(',', '').replace('٬', '')
            value = value.replace('(', '-').replace(')', '')
            value = value.replace('−', '-').replace('–', '-')
            value = value.replace('/', '').replace('\\', '')

            # حذف همه کاراکترها به جز اعداد، نقطه و منفی
            value = ''.join(c for c in value if c.isdigit() or c in '.-')

            if value and value not in ['-', '.']:
                num = float(value)
                if abs(num) < 1e-10:  # حذف مقادیر خیلی کوچک
                    return 0
                return num
            return 0

        except:
            return 0

    def read_financial_data(self, file_path):
        """خواندن داده‌های مالی از فایل اکسل با دقت بیشتر"""
        try:
            # خواندن فایل با روش‌های مختلف برای اطمینان از خواندن صحیح داده‌ها
            try:
                df = pd.read_excel(file_path, header=None)
                if df.empty:
                    df = pd.read_excel(file_path, header=0)
            except:
                try:
                    df = pd.read_excel(file_path, header=0)
                except Exception as e:
                    print(f"خطا در خواندن فایل {file_path}: {str(e)}")
                    return None

            # حذف ستون‌های خالی
            df = df.dropna(axis=1, how='all')

            # پر کردن مقادیر NaN با مقدار خالی
            df = df.fillna('')

            # استخراج سال از نام فایل
            try:
                year = str(file_path).split('_')[0].split('\\')[-1]
            except:
                year = str(file_path).split('/')[-1].split('_')[0]

            data = {'سال': year}

            # جستجوی مقادیر با چند روش مختلف
            for metric, patterns in self.search_patterns.items():
                found_value = False
                max_value = 0

                # روش اول: جستجوی مستقیم
                value = self.find_value_in_df(df, patterns)
                if value > 0:
                    max_value = value
                    found_value = True

                # روش دوم: جستجو در ترانسپوز dataframe
                if not found_value:
                    df_t = df.transpose()
                    value = self.find_value_in_df(df_t, patterns)
                    if value > 0:
                        max_value = value
                        found_value = True

                # روش سوم: جستجو با تبدیل همه سلول‌ها به رشته
                if not found_value:
                    df_str = df.astype(str)
                    value = self.find_value_in_df(df_str, patterns)
                    if value > 0:
                        max_value = value
                        found_value = True

                # ذخیره مقدار پیدا شده
                data[metric] = max_value

                # نمایش مقدار پیدا شده با فرمت مناسب
                if max_value > 0:
                    print(f"{metric}: {max_value:,.0f}")
                else:
                    print(f"{metric}: 0")

            # اطمینان از وجود همه فیلدها
            for metric in self.search_patterns.keys():
                if metric not in data:
                    data[metric] = 0

            return data

        except Exception as e:
            print(f"خطا در خواندن فایل {file_path}: {str(e)}")
            return None

    def calculate_ratios(self, data):
        """محاسبه نسبت‌های مالی با دقت بالا"""
        ratios = {}
        try:
            # تابع کمکی برای محاسبه نسبت‌های مالی با کنترل خطا
            def safe_divide(numerator, denominator, multiplier=1):
                try:
                    if denominator and isinstance(denominator, (int, float)) and denominator != 0:
                        ratio = (float(numerator) / float(denominator)) * multiplier
                        if abs(ratio) > 0.000001:  # حذف مقادیر بسیار کوچک
                            return round(ratio, 6)
                except (TypeError, ValueError, ZeroDivisionError):
                    pass
                return None

            # دریافت مقادیر پایه با مقدار پیش‌فرض صفر
            current_assets = float(data.get('دارایی جاری', 0))
            current_liab = float(data.get('بدهی جاری', 0))
            inventory = float(data.get('موجودی کالا', 0))
            sales = float(data.get('فروش', 0))
            gross_profit = float(data.get('سود ناخالص', 0))
            operating_profit = float(data.get('سود عملیاتی', 0))
            net_profit = float(data.get('سود خالص', 0))
            total_assets = float(data.get('کل دارایی ها', 0))
            total_liab = float(data.get('کل بدهی ها', 0))

            # محاسبه نسبت‌های نقدینگی
            if current_liab > 0:
                # نسبت جاری
                current_ratio = safe_divide(current_assets, current_liab)
                if current_ratio is not None:
                    ratios['نسبت جاری'] = current_ratio

                # نسبت آنی
                quick_ratio = safe_divide(current_assets - inventory, current_liab)
                if quick_ratio is not None:
                    ratios['نسبت آنی'] = quick_ratio

            # محاسبه نسبت‌های سودآوری
            if sales > 0:
                # حاشیه سود ناخالص
                if gross_profit != 0:
                    gross_margin = safe_divide(gross_profit, sales, 100)
                    if gross_margin is not None:
                        ratios['حاشیه سود ناخالص'] = gross_margin

                # حاشیه سود عملیاتی
                if operating_profit != 0:
                    operating_margin = safe_divide(operating_profit, sales, 100)
                    if operating_margin is not None:
                        ratios['حاشیه سود عملیاتی'] = operating_margin

                # حاشیه سود خالص
                if net_profit != 0:
                    net_margin = safe_divide(net_profit, sales, 100)
                    if net_margin is not None:
                        ratios['حاشیه سود خالص'] = net_margin

            # محاسبه نسبت بدهی
            if total_assets > 0 and total_liab != 0:
                debt_ratio = safe_divide(total_liab, total_assets, 100)
                if debt_ratio is not None:
                    ratios['نسبت بدهی'] = debt_ratio

            # چاپ نسبت‌های محاسبه شده برای اطمینان
            for ratio_name, ratio_value in ratios.items():
                print(f"{ratio_name}: {ratio_value:,.6f}")

            return ratios

        except Exception as e:
            print(f"خطا در محاسبه نسبت‌ها: {str(e)}")
            print(f"داده‌های ورودی: {data}")
            return {}

    def save_results(self, results, output_path):
        """ذخیره نتایج در اکسل با فرمت عمودی و شرکت‌ها در هدر"""
        try:
            with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
                workbook = writer.book

                # تعریف فرمت‌ها
                header_format = workbook.add_format({
                    'bold': True,
                    'align': 'center',
                    'valign': 'vcenter',
                    'bg_color': '#D8E4BC',
                    'border': 1,
                    'text_wrap': True
                })

                number_format = workbook.add_format({
                    'num_format': '#,##0.000000',
                    'align': 'center',
                    'border': 1
                })

                # تعریف متغیرها و نسبت‌ها
                ratios = [
                    'نسبت جاری',
                    'نسبت آنی',
                    'حاشیه سود ناخالص',
                    'حاشیه سود عملیاتی',
                    'حاشیه سود خالص',
                    'نسبت بدهی'
                ]

                years = ['1398', '1399', '1400', '1401', '1402']
                companies = sorted(results.keys())

                # ایجاد شیت نسبت‌های مالی
                worksheet = workbook.add_worksheet('نسبت‌های مالی')

                # تنظیم عرض ستون‌ها
                worksheet.set_column(0, 0, 25)  # ستون شاخص
                worksheet.set_column(1, 30, 15)  # ستون‌های داده

                # نوشتن هدر شرکت‌ها
                current_col = 1
                for company in companies:
                    # ادغام سلول‌ها برای نام شرکت
                    worksheet.merge_range(0, current_col, 0, current_col + 4, company, header_format)

                    # نوشتن سال‌ها
                    for i, year in enumerate(years):
                        worksheet.write(1, current_col + i, year, header_format)

                    current_col += 5

                # نوشتن نام نسبت‌ها
                for i, ratio in enumerate(ratios):
                    worksheet.write(i + 2, 0, ratio, header_format)

                # نوشتن داده‌ها
                # در تابع save_results، قسمت نوشتن داده‌ها را اینطور تغییر دهید:

                # نوشتن داده‌ها
                current_col = 1
                for company in companies:
                    for year_idx, year in enumerate(years):
                        if year in results[company]:
                            ratios_data = results[company][year].get('نسبت‌ها', {})
                            for ratio_idx, ratio in enumerate(ratios):
                                value = ratios_data.get(ratio)
                                if value is not None and value != 0:
                                    worksheet.write_number(ratio_idx + 2, current_col + year_idx, value, number_format)
                                else:
                                    worksheet.write_blank(ratio_idx + 2, current_col + year_idx, None, number_format)

                    current_col += 5

                # تنظیم فریز پنل
                worksheet.freeze_panes(2, 1)

                print(f"\nنتایج با موفقیت در فایل زیر ذخیره شد:\n{output_path}")

                # چاپ مقادیر برای اطمینان از صحت داده‌ها
                print("\nمقادیر ذخیره شده:")
                for company in companies:
                    print(f"\nشرکت {company}:")
                    for year in years:
                        if year in results[company]:
                            ratios_data = results[company][year].get('نسبت‌ها', {})
                            print(f"\nسال {year}:")
                            for ratio in ratios:
                                value = ratios_data.get(ratio, 0)
                                print(f"{ratio}: {value:.6f}")

                return True

        except Exception as e:
            print(f"خطا در ذخیره نتایج: {str(e)}")
            return False

def main():
    print("\n=== سیستم تحلیل مالی ===")
    print(f"زمان اجرا: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"کاربر: {os.getenv('USERNAME', 'unknown')}")

    try:
        # دریافت مسیر
        folder_path = input("\nلطفاً مسیر پوشه حاوی فایل‌های اکسل را وارد کنید: ").strip()
        if not os.path.exists(folder_path):
            print("خطا: مسیر وارد شده وجود ندارد!")
            return

        # ایجاد آنالایزر
        analyzer = FinancialAnalyzer(folder_path)

        # دریافت نام شرکت‌ها
        companies = []
        print("\nلطفاً نام شرکت‌ها را وارد کنید (برای پایان، Enter خالی بزنید):")
        while len(companies) < 5:
            company = input(f"نام شرکت {len(companies) + 1}: ").strip()
            if not company:
                break
            companies.append(company)

        if not companies:
            print("هیچ شرکتی برای تحلیل وارد نشده است!")
            return

        # پردازش هر شرکت
        all_results = {}
        for company in companies:
            print(f"\nپردازش شرکت {company}:")
            files = list(Path(folder_path).glob(f'*{company}*.xlsx'))

            if not files:
                print(f"هیچ فایلی برای شرکت {company} یافت نشد!")
                continue

            company_data = {}
            for file in sorted(files):
                print(f"\nپردازش فایل: {file.name}")
                data = analyzer.read_financial_data(file)

                if data:
                    year = data['سال']
                    ratios = analyzer.calculate_ratios(data)
                    company_data[year] = {
                        'متغیرها': data,
                        'نسبت‌ها': ratios
                    }

            if company_data:
                all_results[company] = company_data

        # ذخیره نتایج
        if all_results:
            output_file = Path(folder_path) / f"نتایج_مالی_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            analyzer.save_results(all_results, output_file)
        else:
            print("\nهیچ داده‌ای برای ذخیره‌سازی یافت نشد!")

    except Exception as e:
        print(f"\nخطای غیرمنتظره: {str(e)}")
    finally:
        print("\nپایان برنامه")


if __name__ == "__main__":
    main()