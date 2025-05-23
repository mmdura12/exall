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
        """جستجوی دقیق‌تر مقدار در دیتافریم با کنترل کیفیت داده‌ها"""
        try:
            def normalize_text(text):
                """نرمال‌سازی پیشرفته متن فارسی با کنترل بیشتر"""
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
                    '،': ' ',  # تبدیل کاما فارسی
                    '؛': ' ',  # تبدیل نقطه‌ویرگول فارسی
                    '۰': '0', '۱': '1', '۲': '2', '۳': '3', '۴': '4',
                    '۵': '5', '۶': '6', '۷': '7', '۸': '8', '۹': '9'
                }
                for old, new in replacements.items():
                    text = text.replace(old, new)
                return ' '.join(text.split())

            def validate_value(value, metric_name=""):
                """اعتبارسنجی مقدار عددی"""
                try:
                    val = float(value)
                    if val <= 0:
                        return None
                    if val > 1e12:  # مقادیر خیلی بزرگ
                        print(f"هشدار: مقدار خیلی بزرگ برای {metric_name}: {val:,.0f}")
                        return None
                    return val
                except (ValueError, TypeError):
                    return None

            # پاکسازی و آماده‌سازی دیتافریم
            df = df.fillna('')
            df = df.replace('[\$,)]', '', regex=True)
            df = df.replace('[(]', '-', regex=True)
            df = df.astype(str)

            found_values = []
            value_sources = {}  # برای ردیابی منبع هر مقدار

            for keyword in keywords:
                normalized_keyword = normalize_text(keyword)
                print(f"\nجستجو برای کلیدواژه: {keyword}")

                for idx, row in df.iterrows():
                    row_values = []

                    for col in df.columns:
                        try:
                            cell_value = normalize_text(row[col])

                            if normalized_keyword in cell_value:
                                col_idx = df.columns.get_loc(col)
                                cell_location = f"سطر {idx + 1}, ستون {col}"

                                # بررسی ستون‌های مجاور با الگوریتم بهبود یافته
                                search_ranges = [
                                    # (شروع نسبت به ستون فعلی, پایان, اولویت)
                                    (1, 6, 3),  # ستون‌های بعدی با اولویت بالا
                                    (-3, 0, 2),  # ستون‌های قبلی با اولویت متوسط
                                    (0, 1, 1),  # ستون فعلی با اولویت پایین
                                ]

                                for start, end, priority in search_ranges:
                                    for i in range(start, end):
                                        target_idx = col_idx + i
                                        if 0 <= target_idx < len(df.columns):
                                            val = self.clean_number(row[df.columns[target_idx]])
                                            if val > 0:
                                                val_validated = validate_value(val, keyword)
                                                if val_validated:
                                                    row_values.append(val_validated)
                                                    value_sources[val_validated] = {
                                                        'location': cell_location,
                                                        'priority': priority,
                                                        'value': f"{val_validated:,.0f}"
                                                    }

                        except Exception as e:
                            print(f"خطا در پردازش سلول: {str(e)}")
                            continue

                    # اضافه کردن مقادیر معتبر به لیست نهایی
                    if row_values:
                        row_values = list(set(row_values))  # حذف تکرار
                        found_values.extend(row_values)

            # پردازش و انتخاب مقدار نهایی
            if found_values:
                print(f"\nتعداد مقادیر یافت شده: {len(found_values)}")

                # حذف مقادیر پرت
                found_values = sorted(found_values)

                if len(found_values) > 4:
                    q1 = np.percentile(found_values, 25)
                    q3 = np.percentile(found_values, 75)
                    iqr = q3 - q1
                    lower_bound = q1 - (1.5 * iqr)
                    upper_bound = q3 + (1.5 * iqr)
                    valid_values = [x for x in found_values if lower_bound <= x <= upper_bound]

                    print(f"مقادیر معتبر پس از حذف مقادیر پرت: {len(valid_values)}")
                    for val in valid_values:
                        source = value_sources.get(val, {})
                        print(f"مقدار: {source.get('value', val):>15} | محل: {source.get('location', 'نامشخص')}")
                else:
                    valid_values = found_values
                    print("تعداد مقادیر کم است، همه مقادیر حفظ می‌شوند")

                if valid_values:
                    if len(valid_values) > 1:
                        max_val = max(valid_values)
                        median_val = np.median(valid_values)
                        ratio = max_val / min(valid_values)

                        if ratio > 1000:
                            print(f"پراکندگی زیاد در داده‌ها (نسبت {ratio:.2f})، استفاده از میانه")
                            return median_val
                        else:
                            print(f"پراکندگی معقول در داده‌ها (نسبت {ratio:.2f})، استفاده از مقدار حداکثر")
                            return max_val
                    else:
                        print("تنها یک مقدار معتبر یافت شد")
                        return valid_values[0]

            print("هیچ مقدار معتبری یافت نشد")
            return None

        except Exception as e:
            print(f"خطای کلی در جستجوی مقدار: {str(e)}")
            import traceback
            print(traceback.format_exc())
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
            # خواندن فایل اکسل با روش‌های مختلف
            df = None
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

            if df is None or df.empty:
                print(f"فایل {file_path} خالی است یا قابل خواندن نیست.")
                return None

            # پاکسازی و آماده‌سازی داده‌ها
            df = df.dropna(axis=1, how='all')  # حذف ستون‌های خالی
            df = df.fillna('')  # پر کردن مقادیر NaN با مقدار خالی

            # استخراج سال از نام فایل
            try:
                file_name = str(file_path)
                if '\\' in file_name:
                    year = file_name.split('_')[0].split('\\')[-1]
                else:
                    year = file_name.split('/')[-1].split('_')[0]
            except:
                print("خطا در استخراج سال از نام فایل")
                return None

            # دیکشنری برای ذخیره داده‌ها
            data = {'سال': year}
            found_data = False

            # جستجوی مقادیر با روش‌های مختلف
            for metric, patterns in self.search_patterns.items():
                max_value = 0
                value_found = False

                # روش‌های مختلف جستجو
                search_attempts = [
                    (df, "اصلی"),
                    (df.transpose(), "ترانسپوز"),
                    (df.astype(str), "متنی")
                ]

                for search_df, method in search_attempts:
                    if not value_found:
                        value = self.find_value_in_df(search_df, patterns)
                        if value > 0:
                            max_value = max(max_value, value)
                            value_found = True
                            found_data = True
                            print(f"{metric} (روش {method}): {max_value:,.0f}")

                # ذخیره مقدار نهایی
                data[metric] = max_value

            # بررسی صحت داده‌ها
            required_fields = [
                'دارایی جاری', 'کل دارایی ها', 'بدهی جاری',
                'کل بدهی ها', 'فروش', 'سود ناخالص',
                'سود عملیاتی', 'سود خالص', 'موجودی کالا'
            ]

            missing_fields = [field for field in required_fields if data.get(field, 0) == 0]

            if missing_fields:
                print("\nفیلدهای یافت نشده:")
                for field in missing_fields:
                    print(f"- {field}")

            if not found_data:
                print("هیچ داده معتبری در فایل یافت نشد!")
                return None

            # چاپ خلاصه داده‌های یافت شده
            print("\nخلاصه داده‌های یافت شده:")
            for metric, value in data.items():
                if metric != 'سال' and value > 0:
                    print(f"{metric}: {value:,.0f}")

            return data

        except Exception as e:
            print(f"خطای کلی در پردازش فایل {file_path}:")
            print(str(e))
            import traceback
            print(traceback.format_exc())
            return None

    def calculate_ratios(self, data):
        """محاسبه نسبت‌های مالی با دقت بالا و کنترل خطا"""
        ratios = {}
        try:
            def safe_divide(numerator, denominator, multiplier=1, metric_name=""):
                """محاسبه نسبت با کنترل خطا و اعتبارسنجی"""
                try:
                    num = float(numerator)
                    den = float(denominator)

                    if den == 0:
                        print(f"هشدار: مخرج صفر در محاسبه {metric_name}")
                        return None

                    ratio = (num / den) * multiplier

                    # حذف مقادیر غیرمنطقی
                    if abs(ratio) < 0.000001:  # خیلی کوچک
                        print(f"هشدار: مقدار محاسبه شده برای {metric_name} خیلی کوچک است")
                        return None
                    if abs(ratio) > 1000000:  # خیلی بزرگ
                        print(f"هشدار: مقدار محاسبه شده برای {metric_name} خیلی بزرگ است")
                        return None

                    return round(ratio, 6)
                except (TypeError, ValueError, ZeroDivisionError) as e:
                    print(f"خطا در محاسبه {metric_name}: {str(e)}")
                    return None

            # بررسی و تبدیل داده‌های ورودی
            financial_metrics = {
                'دارایی جاری': 'current_assets',
                'کل دارایی ها': 'total_assets',
                'بدهی جاری': 'current_liab',
                'کل بدهی ها': 'total_liab',
                'فروش': 'sales',
                'سود ناخالص': 'gross_profit',
                'سود عملیاتی': 'operating_profit',
                'سود خالص': 'net_profit',
                'موجودی کالا': 'inventory'
            }

            # تبدیل و اعتبارسنجی داده‌ها
            metrics = {}
            print("\nمقادیر ورودی برای محاسبات:")
            for fa_name, en_name in financial_metrics.items():
                value = data.get(fa_name, 0)
                try:
                    value = float(value)
                    if value < 0:
                        print(f"هشدار: مقدار منفی برای {fa_name}: {value:,.0f}")
                    metrics[en_name] = value
                    print(f"{fa_name}: {value:,.0f}")
                except (TypeError, ValueError) as e:
                    print(f"خطا در تبدیل {fa_name}: {str(e)}")
                    metrics[en_name] = 0

            print("\nمحاسبه نسبت‌های مالی:")

            # محاسبه نسبت جاری
            if metrics['current_liab'] > 0:
                current_ratio = safe_divide(
                    metrics['current_assets'],
                    metrics['current_liab'],
                    1,
                    "نسبت جاری"
                )
                if current_ratio is not None:
                    ratios['نسبت جاری'] = current_ratio
                    print(f"نسبت جاری: {current_ratio:.6f}")

            # محاسبه نسبت آنی
            if metrics['current_liab'] > 0:
                quick_ratio = safe_divide(
                    metrics['current_assets'] - metrics['inventory'],
                    metrics['current_liab'],
                    1,
                    "نسبت آنی"
                )
                if quick_ratio is not None:
                    ratios['نسبت آنی'] = quick_ratio
                    print(f"نسبت آنی: {quick_ratio:.6f}")

            # محاسبه نسبت‌های سودآوری
            if metrics['sales'] > 0:
                # حاشیه سود ناخالص
                if metrics['gross_profit'] != 0:
                    gross_margin = safe_divide(
                        metrics['gross_profit'],
                        metrics['sales'],
                        100,
                        "حاشیه سود ناخالص"
                    )
                    if gross_margin is not None:
                        ratios['حاشیه سود ناخالص'] = gross_margin
                        print(f"حاشیه سود ناخالص: {gross_margin:.6f}%")

                # حاشیه سود عملیاتی
                if metrics['operating_profit'] != 0:
                    operating_margin = safe_divide(
                        metrics['operating_profit'],
                        metrics['sales'],
                        100,
                        "حاشیه سود عملیاتی"
                    )
                    if operating_margin is not None:
                        ratios['حاشیه سود عملیاتی'] = operating_margin
                        print(f"حاشیه سود عملیاتی: {operating_margin:.6f}%")

                # حاشیه سود خالص
                if metrics['net_profit'] != 0:
                    net_margin = safe_divide(
                        metrics['net_profit'],
                        metrics['sales'],
                        100,
                        "حاشیه سود خالص"
                    )
                    if net_margin is not None:
                        ratios['حاشیه سود خالص'] = net_margin
                        print(f"حاشیه سود خالص: {net_margin:.6f}%")

            # محاسبه نسبت بدهی
            if metrics['total_assets'] > 0:
                debt_ratio = safe_divide(
                    metrics['total_liab'],
                    metrics['total_assets'],
                    100,
                    "نسبت بدهی"
                )
                if debt_ratio is not None:
                    ratios['نسبت بدهی'] = debt_ratio
                    print(f"نسبت بدهی: {debt_ratio:.6f}%")

            # خلاصه نتایج
            print(f"\nتعداد نسبت‌های محاسبه شده: {len(ratios)}")
            if not ratios:
                print("هشدار: هیچ نسبتی محاسبه نشد!")

            return ratios

        except Exception as e:
            print(f"\nخطای کلی در محاسبه نسبت‌ها: {str(e)}")
            import traceback
            print(traceback.format_exc())
            print("\nداده‌های ورودی:")
            for key, value in data.items():
                print(f"{key}: {value}")
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
        # تنظیم مسیر خروجی
        output_base_path = Path("C:\\retina_env\\BOT\\sorat\\reports")
        output_base_path.mkdir(parents=True, exist_ok=True)

        # دریافت مسیر ورودی
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
            files.sort()  # مرتب‌سازی فایل‌ها بر اساس نام

            if not files:
                print(f"هیچ فایلی برای شرکت {company} یافت نشد!")
                continue

            company_data = {}
            for file in files:
                try:
                    print(f"\nپردازش فایل: {file.name}")

                    # خواندن داده‌های مالی
                    data = analyzer.read_financial_data(file)
                    if data and isinstance(data, dict):
                        year = data.get('سال')
                        if year:
                            # محاسبه نسبت‌ها
                            ratios = analyzer.calculate_ratios(data)
                            if ratios:  # اگر نسبت‌ها محاسبه شدند
                                print(f"\nنسبت‌های محاسبه شده برای سال {year}:")
                                for ratio_name, ratio_value in ratios.items():
                                    print(f"{ratio_name}: {ratio_value:.6f}")

                                company_data[year] = {
                                    'متغیرها': data,
                                    'نسبت‌ها': ratios
                                }
                            else:
                                print(f"خطا: نسبت‌ها برای سال {year} محاسبه نشدند")
                        else:
                            print("خطا: سال در داده‌ها یافت نشد")
                    else:
                        print("خطا: داده‌های معتبر خوانده نشد")

                except Exception as e:
                    print(f"خطا در پردازش فایل {file.name}: {str(e)}")
                    continue

            if company_data:
                all_results[company] = company_data
                print(f"\nداده‌های شرکت {company} با موفقیت پردازش شد.")
            else:
                print(f"\nهیچ داده معتبری برای شرکت {company} یافت نشد.")

        # ذخیره نتایج
        if all_results:
            # ایجاد پوشه با تاریخ امروز
            today_folder = output_base_path / datetime.now().strftime('%Y-%m-%d')
            today_folder.mkdir(exist_ok=True)

            # ایجاد نام فایل
            timestamp = datetime.now().strftime('%H%M%S')
            file_name = f"نتایج_مالی_{timestamp}.xlsx"
            output_file = today_folder / file_name

            # ذخیره فایل
            success = analyzer.save_results(all_results, output_file)

            if success:
                print(f"\nفایل با موفقیت در مسیر زیر ذخیره شد:")
                print(f"{output_file}")
            else:
                print("\nخطا در ذخیره فایل!")
        else:
            print("\nهیچ داده‌ای برای ذخیره‌سازی یافت نشد!")

    except Exception as e:
        print(f"\nخطای غیرمنتظره: {str(e)}")
        import traceback
        print(traceback.format_exc())
    finally:
        print("\nپایان برنامه")


if __name__ == "__main__":
    main()