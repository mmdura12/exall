import os
import pandas as pd
import numpy as np
import warnings
from datetime import datetime
from pathlib import Path
import matplotlib.pyplot as plt
import seaborn as sns


warnings.filterwarnings('ignore')


class FinancialAnalyzer:
    def __init__(self, base_folder):
        self.base_folder = Path(base_folder)
        self.output_folder = self.base_folder / 'reports'
        self.output_folder.mkdir(exist_ok=True)

        # الگوهای جستجو برای متغیرهای مالی
        self.search_patterns = {
            'دارایی جاری': [
                'دارایی‌های جاری',
                'داراییهای جاری',
                'دارایی های جاری',
                'جمع دارایی‌های جاری',
                'جمع داراییهای جاری',
                'جمع دارایی های جاری',
                'جمع کل دارایی های جاری',
                'دارایی جاری',
                'دارایی‌جاری'
            ],
            'کل دارایی ها': [
                'جمع دارایی‌ها',
                'جمع داراییها',
                'جمع کل دارایی‌ها',
                'جمع کل داراییها',
                'کل دارایی‌ها',
                'کل داراییها',
                'دارایی‌ها',
                'داراییها',
                'جمع دارایی ها'
            ],
            'بدهی جاری': [
                'بدهی‌های جاری',
                'بدهیهای جاری',
                'بدهی های جاری',
                'جمع بدهی‌های جاری',
                'جمع بدهیهای جاری',
                'جمع بدهی های جاری',
                'بدهی جاری',
                'بدهی‌جاری'
            ],
            'کل بدهی ها': [
                'جمع بدهی‌ها',
                'جمع بدهیها',
                'جمع کل بدهی‌ها',
                'جمع کل بدهیها',
                'کل بدهی‌ها',
                'کل بدهیها',
                'بدهی‌ها',
                'بدهیها',
                'جمع بدهی ها'
            ],
            'فروش': [
                'درآمدهای عملیاتی',
                'درآمد عملیاتی',
                'فروش خالص',
                'فروش',
                'درآمد حاصل از فروش',
                'جمع فروش',
                'جمع درآمد عملیاتی'
            ],
            'سود ناخالص': [
                'سود ناخالص',
                'سود (زیان) ناخالص',
                'سود/زیان ناخالص',
                'سودناخالص'
            ],
            'سود عملیاتی': [
                'سود عملیاتی',
                'سود (زیان) عملیاتی',
                'سود/زیان عملیاتی',
                'سودعملیاتی'
            ],
            'سود خالص': [
                'سود خالص',
                'سود (زیان) خالص',
                'سود/زیان خالص',
                'سودخالص',
                'سود خالص دوره'
            ],
            'موجودی کالا': [
                'موجودی مواد و کالا',
                'موجودی کالا',
                'موجودی‌های مواد و کالا',
                'موجودی‌کالا',
                'موجودیهای مواد و کالا'
            ],
            'حساب های دریافتنی': [
                'حساب‌های دریافتنی تجاری',
                'حسابهای دریافتنی',
                'دریافتنی‌های تجاری',
                'حساب های دریافتنی تجاری',
                'حسابهای دریافتنی تجاری'
            ]
        }

    def find_value_in_df(self, df, patterns):
        """جستجوی پیشرفته مقادیر در دیتافریم"""
        try:
            # پیش‌پردازش داده‌ها
            df = df.fillna('')
            df = df.astype(str).applymap(lambda x: str(x).strip())
            results = []

            def extract_numbers_from_cell(cell_value):
                """استخراج تمام اعداد معتبر از یک سلول"""
                numbers = []
                parts = str(cell_value).split()
                for part in parts:
                    value = self.clean_number(part)
                    if value > 0:
                        numbers.append(value)
                return numbers

            def search_pattern_in_cell(cell_value, pattern):
                """بررسی تطابق الگو با محتوای سلول"""
                cell_value = str(cell_value).strip()
                pattern = str(pattern).strip()

                # حالت‌های مختلف نوشتاری
                cell_variations = [
                    cell_value,
                    cell_value.replace('‌', ' '),  # نیم‌فاصله
                    cell_value.replace(' ', ''),  # بدون فاصله
                    cell_value.replace('ي', 'ی'),  # ی عربی
                    cell_value.replace('ك', 'ک')  # ک عربی
                ]

                pattern_variations = [
                    pattern,
                    pattern.replace('‌', ' '),
                    pattern.replace(' ', ''),
                    pattern.replace('ي', 'ی'),
                    pattern.replace('ك', 'ک')
                ]

                return any(p in c for p in pattern_variations for c in cell_variations)

            # بررسی هر الگو
            for pattern in patterns:
                rows, cols = df.shape

                # جستجو در ماتریس
                for i in range(rows):
                    for j in range(cols):
                        current_cell = df.iloc[i, j]

                        if search_pattern_in_cell(current_cell, pattern):
                            # بررسی سلول‌های اطراف در محدوده بزرگتر
                            search_range = [-3, -2, -1, 0, 1, 2, 3]
                            for di in search_range:
                                for dj in search_range:
                                    if di == 0 and dj == 0:
                                        continue

                                    new_i, new_j = i + di, j + dj
                                    if 0 <= new_i < rows and 0 <= new_j < cols:
                                        target_cell = df.iloc[new_i, new_j]
                                        numbers = extract_numbers_from_cell(target_cell)

                                        for number in numbers:
                                            results.append({
                                                'value': number,
                                                'pattern': pattern,
                                                'distance': abs(di) + abs(dj),
                                                'position': (new_i, new_j),
                                                'original': (i, j)
                                            })

                            # بررسی خود سلول برای اعداد
                            numbers = extract_numbers_from_cell(current_cell)
                            for number in numbers:
                                results.append({
                                    'value': number,
                                    'pattern': pattern,
                                    'distance': 0,
                                    'position': (i, j),
                                    'original': (i, j)
                                })

            if results:
                # مرتب‌سازی نتایج بر اساس معیارهای مختلف
                results.sort(key=lambda x: (
                    x['distance'],  # اولویت اول: فاصله کمتر
                    -x['value'],  # اولویت دوم: مقدار بیشتر
                    x['position'][0]  # اولویت سوم: سطر کمتر
                ))

                best_match = results[0]
                print(f"یافتن مقدار برای '{best_match['pattern']}': {best_match['value']:,.0f} "
                      f"در موقعیت {best_match['position']}")
                return best_match['value']

            return None  # به جای 0، None برمی‌گردانیم

        except Exception as e:
            print(f"خطا در جستجوی مقدار: {str(e)}")
            return None

    def clean_number(self, value):
        """تمیز کردن و تبدیل مقادیر عددی با دقت بالا"""
        try:
            if pd.isna(value):
                return 0

            # تبدیل به رشته
            value = str(value).strip()

            # حذف کاراکترهای خاص
            replacements = {
                ',': '', '٬': '', '،': '',
                '(': '-', ')': '',
                '−': '-', '–': '-', '—': '-',
                'ـ': '', '_': '',
                '\u200c': '', '\u200b': '',
                'ر.ا': '', 'ريال': '', 'ریال': '',
                '%': '', '٪': ''
            }

            for old, new in replacements.items():
                value = value.replace(old, new)

            # تبدیل اعداد فارسی
            persian_nums = {
                '۰': '0', '۱': '1', '۲': '2', '۳': '3', '۴': '4',
                '۵': '5', '۶': '6', '۷': '7', '۸': '8', '۹': '9'
            }
            for persian, latin in persian_nums.items():
                value = value.replace(persian, latin)

            # استخراج عدد
            num_str = ''.join(c for c in value if c.isdigit() or c in '.-')
            if num_str and num_str not in ['-', '.']:
                try:
                    number = float(num_str)
                    # بررسی محدوده معقول
                    if 0 < abs(number) < 1e12:
                        return number
                except:
                    pass

            return 0

        except:
            return 0

    def read_financial_data(self, file_path):
        """خواندن داده‌های مالی با تکمیل مقادیر گمشده"""
        try:
            print(f"\nخواندن فایل: {file_path}")
            data = {}

            # خواندن تمام شیت‌ها
            xl = pd.ExcelFile(file_path)

            for sheet_name in xl.sheet_names:
                print(f"\nبررسی شیت {sheet_name}")

                # خواندن با تنظیمات مختلف
                df = pd.read_excel(xl, sheet_name=sheet_name, header=None)

                # جستجوی مقادیر
                for metric, patterns in self.search_patterns.items():
                    if metric not in data:
                        value = self.find_value_in_df(df, patterns)
                        if value is not None and value > 0:
                            data[metric] = value
                            print(f"یافتن {metric}: {value:,.0f}")

            # تکمیل مقادیر گمشده با تخمین‌های منطقی
            if data:
                estimated_data = self.estimate_missing_values(data)
                data.update(estimated_data)

                print("\nخلاصه نهایی مقادیر:")
                for metric, value in data.items():
                    print(f"{metric}: {value:,.0f}")
                return data

            return None

        except Exception as e:
            print(f"خطا در خواندن فایل: {str(e)}")
            return None

    def estimate_missing_values(self, data):
        """تخمین مقادیر گمشده با استفاده از روابط منطقی"""
        estimated = {}

        # تخمین دارایی‌های جاری
        if 'دارایی جاری' not in data and 'کل دارایی ها' in data:
            estimated['دارایی جاری'] = data['کل دارایی ها'] * 0.6

        # تخمین بدهی‌های جاری
        if 'بدهی جاری' not in data and 'کل بدهی ها' in data:
            estimated['بدهی جاری'] = data['کل بدهی ها'] * 0.7

        # تخمین موجودی کالا
        if 'موجودی کالا' not in data and 'دارایی جاری' in data:
            estimated['موجودی کالا'] = data['دارایی جاری'] * 0.3

        # تخمین سودها
        if 'فروش' in data:
            sales = data['فروش']
            if 'سود ناخالص' not in data:
                estimated['سود ناخالص'] = sales * 0.3
            if 'سود عملیاتی' not in data:
                estimated['سود عملیاتی'] = sales * 0.2
            if 'سود خالص' not in data:
                estimated['سود خالص'] = sales * 0.15

        # گزارش تخمین‌ها
        if estimated:
            print("\nمقادیر تخمین زده شده:")
            for metric, value in estimated.items():
                print(f"{metric} (تخمینی): {value:,.0f}")

        return estimated

    def calculate_ratios(self, data):
        """محاسبه نسبت‌های مالی با کنترل دقیق خطا"""
        ratios = {}
        try:
            def safe_divide(numerator, denominator, factor=1):
                """تقسیم ایمن با کنترل محدوده"""
                try:
                    if denominator and denominator != 0:
                        result = (numerator / denominator) * factor
                        # محدود کردن نتیجه به محدوده منطقی
                        if -10000 <= result <= 10000:
                            return result
                except:
                    pass
                return 0

            # نسبت‌های نقدینگی
            current_assets = data.get('دارایی جاری', 0)
            current_liab = data.get('بدهی جاری', 0)

            ratios['نسبت جاری'] = safe_divide(current_assets, current_liab)
            ratios['نسبت آنی'] = safe_divide(
                current_assets - data.get('موجودی کالا', 0),
                current_liab
            )

            # نسبت‌های سودآوری
            sales = data.get('فروش', 0)
            ratios['حاشیه سود ناخالص'] = safe_divide(data.get('سود ناخالص', 0), sales, 100)
            ratios['حاشیه سود عملیاتی'] = safe_divide(data.get('سود عملیاتی', 0), sales, 100)
            ratios['حاشیه سود خالص'] = safe_divide(data.get('سود خالص', 0), sales, 100)

            # نسبت‌های اهرمی
            ratios['نسبت بدهی'] = safe_divide(
                data.get('کل بدهی ها', 0),
                data.get('کل دارایی ها', 0),
                100
            )

            # نمایش نتایج
            print("\nنسبت‌های محاسبه شده:")
            for name, value in ratios.items():
                print(f"{name}: {value:.2f}%")

            return ratios

        except Exception as e:
            print(f"خطا در محاسبه نسبت‌ها: {str(e)}")
            return ratios
        import matplotlib.pyplot as plt
        import pandas as pd
        from pathlib import Path
        import seaborn as sns

    def plot_financial_metrics(self, results):
        """
        رسم نمودارهای خطی برای متغیرهای مالی هر شرکت در سال‌های مختلف
        """
        import matplotlib.pyplot as plt
        import seaborn as sns

        # تنظیم استایل نمودار - اصلاح شده
        plt.style.use('default')  # استفاده از استایل پیش‌فرض
        sns.set_theme()  # تنظیم تم seaborn

        # تنظیم فونت برای نمایش متن فارسی
        plt.rcParams['font.family'] = 'Arial'  # یا هر فونت دیگری که از فارسی پشتیبانی می‌کند

        # تعریف رنگ‌های مختلف برای شرکت‌های مختلف
        colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd']

        # تعریف متغیرهای مالی اصلی برای نمایش
        main_metrics = [
            'دارایی جاری',
            'کل دارایی ها',
            'بدهی جاری',
            'کل بدهی ها',
            'فروش',
            'سود ناخالص',
            'سود عملیاتی',
            'سود خالص'
        ]

        # ایجاد فولدر برای نمودارها
        charts_folder = self.output_folder / 'charts'
        charts_folder.mkdir(exist_ok=True)

        # تبدیل داده‌ها به دیتافریم
        all_data = []
        for company in results:
            for year in results[company]:
                data = results[company][year].get('متغیرها',
                                                  {}).copy()  # استفاده از copy برای جلوگیری از تغییر دیکشنری اصلی
                data['شرکت'] = company
                data['سال'] = year
                all_data.append(data)

        df = pd.DataFrame(all_data)

        # رسم نمودار برای هر متغیر مالی
        for metric in main_metrics:
            if metric not in df.columns:
                print(f"متغیر {metric} در داده‌ها یافت نشد.")
                continue

            plt.figure(figsize=(12, 6))

            for i, company in enumerate(df['شرکت'].unique()):
                company_data = df[df['شرکت'] == company]
                plt.plot(company_data['سال'],
                         company_data[metric],
                         marker='o',
                         label=company,
                         color=colors[i % len(colors)],
                         linewidth=2)

            plt.title(f'روند {metric} برای شرکت‌های مختلف', fontsize=14, pad=20)
            plt.xlabel('سال', fontsize=12)
            plt.ylabel('مقدار (ریال)', fontsize=12)
            plt.grid(True, linestyle='--', alpha=0.7)
            plt.legend(title='شرکت‌ها', bbox_to_anchor=(1.05, 1), loc='upper left')
            plt.tight_layout()

            # ذخیره نمودار
            chart_file = charts_folder / f"{metric.replace(' ', '_')}_trend.png"
            plt.savefig(chart_file, dpi=300, bbox_inches='tight')
            plt.close()

        # رسم نمودار نسبت‌های مالی
        financial_ratios = [
            'نسبت جاری',
            'نسبت آنی',
            'حاشیه سود ناخالص',
            'حاشیه سود عملیاتی',
            'حاشیه سود خالص',
            'نسبت بدهی'
        ]

        ratio_data = []
        for company in results:
            for year in results[company]:
                ratios = results[company][year].get('نسبت‌ها',
                                                    {}).copy()  # استفاده از copy برای جلوگیری از تغییر دیکشنری اصلی
                ratios['شرکت'] = company
                ratios['سال'] = year
                ratio_data.append(ratios)

        df_ratios = pd.DataFrame(ratio_data)

        # رسم نمودار برای هر نسبت مالی
        for ratio in financial_ratios:
            if ratio not in df_ratios.columns:
                print(f"نسبت {ratio} در داده‌ها یافت نشد.")
                continue

            plt.figure(figsize=(12, 6))

            for i, company in enumerate(df_ratios['شرکت'].unique()):
                company_data = df_ratios[df_ratios['شرکت'] == company]
                plt.plot(company_data['سال'],
                         company_data[ratio],
                         marker='o',
                         label=company,
                         color=colors[i % len(colors)],
                         linewidth=2)

            plt.title(f'روند {ratio} برای شرکت‌های مختلف', fontsize=14, pad=20)
            plt.xlabel('سال', fontsize=12)
            plt.ylabel('درصد', fontsize=12)
            plt.grid(True, linestyle='--', alpha=0.7)
            plt.legend(title='شرکت‌ها', bbox_to_anchor=(1.05, 1), loc='upper left')
            plt.tight_layout()

            # ذخیره نمودار
            chart_file = charts_folder / f"{ratio.replace(' ', '_')}_trend.png"
            plt.savefig(chart_file, dpi=300, bbox_inches='tight')
            plt.close()

        print("نمودارها با موفقیت رسم و ذخیره شدند.")

    def save_to_excel(self, results):
        """ذخیره نتایج در فایل اکسل"""
        try:
            # لیست تمام سال‌ها
            all_years = ['1398', '1399', '1400', '1401', '1402']

            metrics_data = []
            ratios_data = []

            # پردازش داده‌ها
            for company in results:
                for year in all_years:
                    # متغیرهای مالی
                    metrics_row = {
                        'شرکت': company,
                        'سال': year
                    }

                    if year in results[company]:
                        company_data = results[company][year]
                        metrics = company_data.get('متغیرها', {})
                        for metric in self.search_patterns.keys():
                            metrics_row[metric] = metrics.get(metric, 0)  # استفاده از صفر به جای None
                    else:
                        for metric in self.search_patterns.keys():
                            metrics_row[metric] = 0  # استفاده از صفر به جای None

                    metrics_data.append(metrics_row)

                    # نسبت‌های مالی
                    ratios_row = {
                        'شرکت': company,
                        'سال': year
                    }

                    if year in results[company]:
                        ratios = company_data.get('نسبت‌ها', {})
                        for ratio in ['نسبت جاری', 'نسبت آنی', 'حاشیه سود ناخالص',
                                      'حاشیه سود عملیاتی', 'حاشیه سود خالص', 'نسبت بدهی']:
                            ratios_row[ratio] = ratios.get(ratio, 0)  # استفاده از صفر به جای None
                    else:
                        for ratio in ['نسبت جاری', 'نسبت آنی', 'حاشیه سود ناخالص',
                                      'حاشیه سود عملیاتی', 'حاشیه سود خالص', 'نسبت بدهی']:
                            ratios_row[ratio] = 0  # استفاده از صفر به جای None

                    ratios_data.append(ratios_row)

            # تبدیل به DataFrame
            df_metrics = pd.DataFrame(metrics_data)
            df_ratios = pd.DataFrame(ratios_data)

            # مرتب‌سازی
            df_metrics = df_metrics.sort_values(['شرکت', 'سال'])
            df_ratios = df_ratios.sort_values(['شرکت', 'سال'])

            # پر کردن مقادیر NaN با صفر
            df_metrics = df_metrics.fillna(0)
            df_ratios = df_ratios.fillna(0)

            # ایجاد فایل خروجی
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_file = self.output_folder / f'financial_analysis_{timestamp}.xlsx'

            # تنظیمات Workbook
            workbook_options = {
                'nan_inf_to_errors': True,
                'strings_to_numbers': True
            }

            with pd.ExcelWriter(output_file, engine='xlsxwriter',
                                engine_kwargs={'options': workbook_options}) as writer:
                # نوشتن داده‌ها
                df_metrics.to_excel(writer, sheet_name='متغیرهای مالی', index=False)
                df_ratios.to_excel(writer, sheet_name='نسبت‌های مالی', index=False)

                workbook = writer.book

                # تعریف فرمت‌ها
                header_format = workbook.add_format({
                    'bold': True,
                    'align': 'center',
                    'valign': 'vcenter',
                    'bg_color': '#D8E4BC',
                    'border': 1
                })

                number_format = workbook.add_format({
                    'num_format': '#,##0',
                    'align': 'center',
                    'border': 1
                })

                percent_format = workbook.add_format({
                    'num_format': '0.00%',
                    'align': 'center',
                    'border': 1
                })

                # اعمال فرمت به شیت‌ها
                for sheet_name in writer.sheets:
                    worksheet = writer.sheets[sheet_name]

                    if sheet_name == 'متغیرهای مالی':
                        df = df_metrics
                        default_format = number_format
                    else:
                        df = df_ratios
                        default_format = percent_format

                    # تنظیم عرض ستون‌ها
                    for col_num, column in enumerate(df.columns):
                        # تنظیم عرض
                        max_length = max(
                            len(str(column)),
                            df[column].astype(str).str.len().max()
                        )
                        worksheet.set_column(col_num, col_num, max_length + 2)

                        # نوشتن هدر
                        worksheet.write(0, col_num, column, header_format)

                        # نوشتن داده‌ها با بررسی نوع داده
                        if column not in ['شرکت', 'سال']:
                            for row in range(len(df)):
                                value = df.iloc[row][column]
                                try:
                                    if pd.isna(value) or value == '':
                                        worksheet.write_blank(row + 1, col_num, None, default_format)
                                    else:
                                        worksheet.write_number(row + 1, col_num, float(value), default_format)
                                except:
                                    worksheet.write_string(row + 1, col_num, str(value), default_format)

            print(f"\nنتایج با موفقیت در فایل زیر ذخیره شد:\n{output_file}")
            print("\nخلاصه اطلاعات ذخیره شده:")
            print(f"تعداد شرکت‌ها: {len(results)}")
            print(f"سال‌های مورد بررسی: {', '.join(all_years)}")
            return True

        except Exception as e:
            print(f"خطا در ذخیره نتایج: {str(e)}")
            import traceback
            print(traceback.format_exc())
            return False


def main():
    print("\n=== سیستم تحلیل مالی ===")
    print(f"زمان اجرا: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"کاربر: {os.getenv('USERNAME', 'unknown')}")

    try:
        folder_path = input("\nلطفاً مسیر پوشه حاوی فایل‌های اکسل را وارد کنید: ").strip()
        if not os.path.exists(folder_path):
            print("خطا: مسیر وارد شده وجود ندارد!")
            return

        analyzer = FinancialAnalyzer(folder_path)

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

        results = {}
        for company in companies:
            print(f"\nپردازش شرکت {company}:")
            company_data = {}

            for year in range(1398, 1403):
                files = list(Path(folder_path).glob(f"{year}_{company}*.xlsx"))
                if files:
                    print(f"\nپردازش سال {year}:")
                    data = analyzer.read_financial_data(files[0])
                    if data:
                        ratios = analyzer.calculate_ratios(data)
                        company_data[str(year)] = {
                            'متغیرها': data,
                            'نسبت‌ها': ratios
                        }

            if company_data:
                results[company] = company_data
                print(f"\nداده‌های شرکت {company} با موفقیت پردازش شد.")
            else:
                print(f"\nهیچ داده‌ای برای شرکت {company} یافت نشد!")

        if results:
            print("\nدر حال رسم نمودارها...")
            try:
                analyzer.plot_financial_metrics(results)
                print("نمودارها با موفقیت در پوشه 'charts' ذخیره شدند.")
            except Exception as chart_error:
                print(f"خطا در رسم نمودارها: {str(chart_error)}")

            print("\nدر حال ذخیره نتایج در اکسل...")
            analyzer.save_to_excel(results)
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