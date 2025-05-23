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
        """جستجوی مقادیر در دیتافریم با دقت بیشتر"""
        try:
            def normalize_text(text):
                """نرمال‌سازی متن فارسی"""
                text = str(text).strip()
                persians = {
                    '‌': ' ', 'ي': 'ی', 'ك': 'ک', '\u200c': ' ',
                    '،': ' ', '؛': ' ', '\n': ' ', '\r': ' ', '\t': ' ',
                    '۰': '0', '۱': '1', '۲': '2', '۳': '3', '۴': '4',
                    '۵': '5', '۶': '6', '۷': '7', '۸': '8', '۹': '9'
                }
                for old, new in persians.items():
                    text = text.replace(old, new)
                return ' '.join(text.split())

            def find_number_in_row(row, col_idx, keyword):
                """یافتن عدد معتبر در سطر"""
                # الگوی جستجو در ستون‌های مجاور
                check_order = [
                    1,  # ستون بعدی
                    2,  # دو ستون بعد
                    -1,  # ستون قبلی
                    0,  # ستون فعلی
                    3,  # سه ستون بعد
                    -2,  # دو ستون قبل
                    4  # چهار ستون بعد
                ]

                found_values = []
                for offset in check_order:
                    target_idx = col_idx + offset
                    if 0 <= target_idx < len(row):
                        value = self.clean_number(row[target_idx])
                        if value > 0 and value < 1e12:  # محدوده معقول
                            found_values.append({
                                'value': value,
                                'distance': abs(offset),
                                'position': target_idx
                            })

                if found_values:
                    # اولویت با نزدیک‌ترین مقدار معتبر
                    found_values.sort(key=lambda x: (x['distance'], -x['value']))
                    return found_values[0]['value']
                return None

            # آماده‌سازی دیتافریم
            df = df.fillna('')
            df = df.replace('[\$,)]', '', regex=True)
            df = df.replace('[(]', '-', regex=True)
            df = df.astype(str)

            best_value = None
            best_location = None
            keyword_matches = []

            # جستجو برای هر کلیدواژه
            for keyword in keywords:
                normalized_keyword = normalize_text(keyword)

                # بررسی هر سطر
                for idx, row in df.iterrows():
                    for col in df.columns:
                        cell_value = normalize_text(row[col])

                        # اگر کلیدواژه در سلول پیدا شد
                        if normalized_keyword in cell_value:
                            col_idx = df.columns.get_loc(col)
                            value = find_number_in_row(row, col_idx, keyword)

                            if value is not None:
                                keyword_matches.append({
                                    'value': value,
                                    'location': f"سطر {idx + 1}, ستون {col}",
                                    'keyword': keyword
                                })

            if keyword_matches:
                # حذف مقادیر تکراری
                unique_values = []
                seen = set()
                for match in keyword_matches:
                    if match['value'] not in seen:
                        unique_values.append(match)
                        seen.add(match['value'])

                if len(unique_values) == 1:
                    best_match = unique_values[0]
                    print(f"\nمقدار یافت شده برای '{best_match['keyword']}': "
                          f"{best_match['value']:,.0f} در {best_match['location']}")
                    return best_match['value']

                elif len(unique_values) > 1:
                    # مرتب‌سازی بر اساس مقدار
                    values = [match['value'] for match in unique_values]
                    values.sort()

                    # بررسی پراکندگی مقادیر
                    if len(values) >= 3:
                        # حذف مقادیر پرت با IQR
                        q1 = np.percentile(values, 25)
                        q3 = np.percentile(values, 75)
                        iqr = q3 - q1
                        lower_bound = q1 - (1.5 * iqr)
                        upper_bound = q3 + (1.5 * iqr)
                        filtered_values = [v for v in values if lower_bound <= v <= upper_bound]
                    else:
                        filtered_values = values

                    if filtered_values:
                        # انتخاب مقدار مناسب
                        max_value = max(filtered_values)
                        min_value = min(filtered_values)
                        ratio = max_value / min_value if min_value > 0 else float('inf')

                        if ratio > 10:  # اختلاف زیاد
                            selected_value = np.median(filtered_values)
                            print(f"\nاستفاده از میانه به دلیل پراکندگی زیاد (نسبت: {ratio:.2f})")
                        else:
                            selected_value = max_value
                            print(f"\nاستفاده از مقدار حداکثر (نسبت: {ratio:.2f})")

                        # نمایش مقدار انتخاب شده
                        matching_location = next(
                            match['location'] for match in unique_values
                            if match['value'] == selected_value
                        )
                        print(f"مقدار نهایی: {selected_value:,.0f} در {matching_location}")
                        return selected_value

            print("هیچ مقدار معتبری یافت نشد")
            return 0

        except Exception as e:
            print(f"خطا در جستجوی مقدار: {str(e)}")
            return 0

    def clean_number(self, value):
        """تبدیل مقادیر به عدد با دقت بالا"""
        try:
            if isinstance(value, (int, float)):
                return float(value)

            # تبدیل به رشته و پاکسازی
            value = str(value).strip()

            # حذف کاراکترهای اضافی
            value = value.replace(',', '')
            value = value.replace('٬', '')
            value = value.replace('(', '-')
            value = value.replace(')', '')

            # تبدیل اعداد فارسی
            persian_nums = {'۰': '0', '۱': '1', '۲': '2', '۳': '3', '۴': '4',
                            '۵': '5', '۶': '6', '۷': '7', '۸': '8', '۹': '9'}
            for persian, latin in persian_nums.items():
                value = value.replace(persian, latin)

            # حذف همه کاراکترها به جز اعداد و علائم خاص
            value = ''.join(c for c in value if c.isdigit() or c in '.-')

            if value and value not in ['-', '.']:
                return float(value)
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
        """محاسبه نسبت‌های مالی با دقت بالا"""
        try:
            def safe_divide(num, denom, multiplier=1, metric_name=""):
                """محاسبه نسبت با کنترل دقیق خطا"""
                try:
                    if not isinstance(num, (int, float)) or not isinstance(denom, (int, float)):
                        print(f"خطا در {metric_name}: مقادیر ورودی باید عددی باشند")
                        return None

                    if denom == 0:
                        print(f"خطا در {metric_name}: مخرج صفر است")
                        return None

                    ratio = (float(num) / float(denom)) * multiplier

                    # کنترل محدوده منطقی
                    if abs(ratio) < 0.000001:
                        print(f"هشدار در {metric_name}: مقدار محاسبه شده بسیار کوچک است")
                        return None
                    if abs(ratio) > 1000:
                        print(f"هشدار در {metric_name}: مقدار محاسبه شده بسیار بزرگ است")
                        return None

                    return round(ratio, 6)
                except Exception as e:
                    print(f"خطا در محاسبه {metric_name}: {str(e)}")
                    return None

            # تبدیل داده‌های ورودی
            metrics = {
                'current_assets': float(data.get('دارایی جاری', 0)),
                'total_assets': float(data.get('کل دارایی ها', 0)),
                'current_liab': float(data.get('بدهی جاری', 0)),
                'total_liab': float(data.get('کل بدهی ها', 0)),
                'inventory': float(data.get('موجودی کالا', 0)),
                'sales': float(data.get('فروش', 0)),
                'gross_profit': float(data.get('سود ناخالص', 0)),
                'operating_profit': float(data.get('سود عملیاتی', 0)),
                'net_profit': float(data.get('سود خالص', 0))
            }

            # نمایش مقادیر ورودی
            print("\n=== مقادیر ورودی ===")
            for name, value in metrics.items():
                print(f"{name}: {value:,.0f}")

            ratios = {}

            # محاسبه نسبت‌های نقدینگی
            print("\n=== محاسبه نسبت‌های نقدینگی ===")
            if metrics['current_liab'] > 0:
                # نسبت جاری
                current_ratio = safe_divide(
                    metrics['current_assets'],
                    metrics['current_liab'],
                    1,
                    "نسبت جاری"
                )
                if current_ratio is not None:
                    ratios['نسبت جاری'] = current_ratio
                    print(f"نسبت جاری: {current_ratio:.6f}")
                    # تحلیل نسبت جاری
                    if current_ratio < 1:
                        print("هشدار: نسبت جاری کمتر از 1 است - نشان‌دهنده مشکل در نقدینگی")
                    elif current_ratio > 3:
                        print("هشدار: نسبت جاری بیش از حد بالاست - احتمال عدم استفاده بهینه از دارایی‌ها")

                # نسبت آنی
                quick_assets = metrics['current_assets'] - metrics['inventory']
                quick_ratio = safe_divide(
                    quick_assets,
                    metrics['current_liab'],
                    1,
                    "نسبت آنی"
                )
                if quick_ratio is not None:
                    ratios['نسبت آنی'] = quick_ratio
                    print(f"نسبت آنی: {quick_ratio:.6f}")
                    # تحلیل نسبت آنی
                    if quick_ratio < 0.5:
                        print("هشدار: نسبت آنی پایین است - ممکن است نشان‌دهنده مشکل نقدینگی باشد")

            # محاسبه نسبت‌های سودآوری
            print("\n=== محاسبه نسبت‌های سودآوری ===")
            if metrics['sales'] > 0:
                # حاشیه سود ناخالص
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
                net_margin = safe_divide(
                    metrics['net_profit'],
                    metrics['sales'],
                    100,
                    "حاشیه سود خالص"
                )
                if net_margin is not None:
                    ratios['حاشیه سود خالص'] = net_margin
                    print(f"حاشیه سود خالص: {net_margin:.6f}%")
                    # تحلیل حاشیه سود
                    if net_margin < 0:
                        print("هشدار: حاشیه سود خالص منفی است")
                    elif net_margin > 50:
                        print("توجه: حاشیه سود خالص بسیار بالاست")

            # محاسبه نسبت‌های اهرمی
            print("\n=== محاسبه نسبت‌های اهرمی ===")
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
                    # تحلیل نسبت بدهی
                    if debt_ratio > 70:
                        print("هشدار: نسبت بدهی بالاست - ریسک مالی زیاد")

            # خلاصه نتایج
            print(f"\n=== خلاصه نتایج ===")
            print(f"تعداد نسبت‌های محاسبه شده: {len(ratios)}")
            if len(ratios) == 0:
                print("هشدار: هیچ نسبتی محاسبه نشد!")
            elif len(ratios) < 4:
                print("هشدار: تعداد نسبت‌های محاسبه شده کمتر از حد انتظار است")

            return ratios

        except Exception as e:
            print(f"\nخطای کلی در محاسبه نسبت‌ها: {str(e)}")
            import traceback
            print(traceback.format_exc())
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