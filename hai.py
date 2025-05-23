import os
import pandas as pd
import numpy as np
import warnings
from datetime import datetime
from pathlib import Path

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
        """جستجوی مقادیر در دیتافریم با دقت بیشتر"""
        try:
            # تبدیل تمام سلول‌ها به رشته و پر کردن مقادیر خالی
            df = df.astype(str).fillna('')

            # برای هر الگو در لیست الگوها
            for pattern in patterns:
                pattern = str(pattern).strip()

                # جستجو در تمام سطرها
                for i in range(len(df)):
                    row = df.iloc[i]

                    # جستجو در هر سلول سطر
                    for j in range(len(row)):
                        cell_value = str(row[j]).strip()

                        # اگر الگو در سلول پیدا شد
                        if pattern in cell_value:
                            # بررسی سلول بعدی در همان سطر
                            if j + 1 < len(row):
                                next_cell = str(row[j + 1]).strip()
                                value = self.clean_number(next_cell)
                                if value > 0:
                                    print(f"مقدار یافت شده برای '{pattern}': {value:,.0f} در سطر {i + 1}")
                                    return value

                            # بررسی سلول قبلی در همان سطر
                            if j - 1 >= 0:
                                prev_cell = str(row[j - 1]).strip()
                                value = self.clean_number(prev_cell)
                                if value > 0:
                                    print(f"مقدار یافت شده برای '{pattern}': {value:,.0f} در سطر {i + 1}")
                                    return value

                            # بررسی سلول‌های سمت راست تا 3 ستون
                            for k in range(2, 4):
                                if j + k < len(row):
                                    right_cell = str(row[j + k]).strip()
                                    value = self.clean_number(right_cell)
                                    if value > 0:
                                        print(f"مقدار یافت شده برای '{pattern}': {value:,.0f} در سطر {i + 1}")
                                        return value

            print(f"مقدار {patterns[0]} یافت نشد")
            return 0

        except Exception as e:
            print(f"خطا در جستجوی مقدار: {str(e)}")
            return 0

    def read_financial_data(self, file_path):
        """خواندن داده‌های مالی از فایل اکسل"""
        try:
            print(f"\nخواندن فایل: {file_path}")

            # خواندن تمام شیت‌های فایل اکسل
            all_sheets = pd.read_excel(file_path, sheet_name=None)

            data = {}
            found_any = False

            # بررسی هر شیت
            for sheet_name, df in all_sheets.items():
                print(f"بررسی شیت {sheet_name}")

                # استخراج مقادیر
                for metric, patterns in self.search_patterns.items():
                    if metric not in data:  # اگر قبلاً پیدا نشده
                        value = self.find_value_in_df(df, patterns)
                        if value > 0:
                            data[metric] = value
                            found_any = True

            if found_any:
                return data
            else:
                print("هیچ داده‌ای در فایل یافت نشد!")
                return None

        except Exception as e:
            print(f"خطا در خواندن فایل: {str(e)}")
            return None

    def clean_number(self, value):
        """تمیز کردن و تبدیل مقادیر عددی"""
        try:
            if isinstance(value, (int, float)):
                return float(value)

            if not isinstance(value, str):
                value = str(value)

            # حذف کاراکترهای غیر عددی
            value = value.replace(',', '').replace('٬', '')
            value = value.replace('(', '-').replace(')', '')
            value = value.replace('−', '-').replace('–', '-')
            value = ''.join(c for c in value if c.isdigit() or c in '.-')

            # تبدیل اعداد فارسی
            persian_nums = {
                '۰': '0', '۱': '1', '۲': '2', '۳': '3', '۴': '4',
                '۵': '5', '۶': '6', '۷': '7', '۸': '8', '۹': '9'
            }
            for persian, latin in persian_nums.items():
                value = value.replace(persian, latin)

            if value and value not in ['-', '.']:
                try:
                    return float(value)
                except:
                    return 0
            return 0

        except:
            return 0


    def calculate_ratios(self, data):
        """محاسبه نسبت‌های مالی"""
        try:
            ratios = {}

            # نسبت‌های نقدینگی
            if data.get('بدهی جاری', 0) != 0:
                ratios['نسبت جاری'] = data['دارایی جاری'] / data['بدهی جاری']
                ratios['نسبت آنی'] = (data['دارایی جاری'] - data.get('موجودی کالا', 0)) / data['بدهی جاری']

            # نسبت‌های سودآوری
            if data.get('فروش', 0) != 0:
                if data.get('سود ناخالص'):
                    ratios['حاشیه سود ناخالص'] = (data['سود ناخالص'] / data['فروش']) * 100
                if data.get('سود عملیاتی'):
                    ratios['حاشیه سود عملیاتی'] = (data['سود عملیاتی'] / data['فروش']) * 100
                if data.get('سود خالص'):
                    ratios['حاشیه سود خالص'] = (data['سود خالص'] / data['فروش']) * 100

            # نسبت‌های اهرمی
            if data.get('کل دارایی ها', 0) != 0:
                ratios['نسبت بدهی'] = (data.get('کل بدهی ها', 0) / data['کل دارایی ها']) * 100

            # چاپ نسبت‌های محاسبه شده
            for name, value in ratios.items():
                print(f"{name}: {value:.2f}")

            return ratios

        except Exception as e:
            print(f"خطا در محاسبه نسبت‌ها: {str(e)}")
            return {}

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