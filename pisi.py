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
        """جستجوی دقیق‌تر مقدار در دیتافریم"""
        try:
            # نرمال‌سازی کلمات کلیدی
            def normalize_text(text):
                """نرمال‌سازی متن فارسی"""
                text = str(text).strip()
                text = text.replace('‌', ' ')  # حذف نیم‌فاصله
                text = text.replace('ي', 'ی')  # یکسان‌سازی ی
                text = text.replace('ك', 'ک')  # یکسان‌سازی ک
                return ' '.join(text.split())  # حذف فاصله‌های اضافی

            # تبدیل DataFrame به string برای جستجوی بهتر
            df = df.fillna('')
            df = df.astype(str)

            for keyword in keywords:
                normalized_keyword = normalize_text(keyword)

                for idx, row in df.iterrows():
                    row_found = False

                    # بررسی هر سلول در سطر
                    for col in df.columns:
                        try:
                            cell_value = normalize_text(row[col])

                            # اگر کلمه کلیدی در سلول پیدا شد
                            if normalized_keyword in cell_value:
                                row_found = True
                                col_idx = df.columns.get_loc(col)

                                # جستجو در ستون‌های بعدی (تا 5 ستون)
                                for i in range(1, 6):
                                    if col_idx + i < len(df.columns):
                                        next_col = df.columns[col_idx + i]
                                        value = self.clean_number(row[next_col])
                                        if value != 0:
                                            return value

                                # جستجو در ستون‌های قبلی (تا 3 ستون)
                                for i in range(1, 4):
                                    if col_idx - i >= 0:
                                        prev_col = df.columns[col_idx - i]
                                        value = self.clean_number(row[prev_col])
                                        if value != 0:
                                            return value

                                # جستجو در همان ستون
                                value = self.clean_number(row[col])
                                if value != 0:
                                    return value

                        except Exception as e:
                            continue  # ادامه جستجو در صورت بروز خطا

                    # اگر کلمه کلیدی در سطر پیدا شد اما مقداری یافت نشد
                    if row_found:
                        # جستجوی عدد در کل سطر
                        for cell in row:
                            value = self.clean_number(cell)
                            if value != 0:
                                return value

            return 0  # اگر هیچ مقداری پیدا نشد

        except Exception as e:
            print(f"خطا در جستجوی مقدار: {str(e)}")
            return 0

    def read_financial_data(self, file_path):
        """خواندن داده‌های مالی از فایل اکسل"""
        try:
            df = pd.read_excel(file_path, header=None)
            year = str(file_path).split('_')[0].split('\\')[-1]

            data = {'سال': year}
            for metric, patterns in self.search_patterns.items():
                value = self.find_value_in_df(df, patterns)
                if value:
                    data[metric] = value
                    print(f"{metric}: {value:,.0f}")

            return data

        except Exception as e:
            print(f"خطا در خواندن فایل: {str(e)}")
            return None

        def calculate_ratios(self, data):  # تورفتگی درست شد
            """محاسبه نسبت‌های مالی با دقت بالا"""
            ratios = {}
            try:
                # نسبت‌های نقدینگی
                if data.get('بدهی جاری', 0) > 0:
                    # محاسبه نسبت جاری
                    current_ratio = data.get('دارایی جاری', 0) / data.get('بدهی جاری', 0)
                    if current_ratio != 0:
                        ratios['نسبت جاری'] = round(current_ratio, 6)

                    # محاسبه نسبت آنی
                    quick_ratio = (data.get('دارایی جاری', 0) - data.get('موجودی کالا', 0)) / data.get('بدهی جاری', 0)
                    if quick_ratio != 0:
                        ratios['نسبت آنی'] = round(quick_ratio, 6)

                # نسبت‌های سودآوری
                if data.get('فروش', 0) > 0:
                    # حاشیه سود ناخالص
                    if data.get('سود ناخالص', 0) != 0:
                        gross_margin = (data.get('سود ناخالص', 0) / data.get('فروش', 0)) * 100
                        ratios['حاشیه سود ناخالص'] = round(gross_margin, 6)

                    # حاشیه سود عملیاتی
                    if data.get('سود عملیاتی', 0) != 0:
                        operating_margin = (data.get('سود عملیاتی', 0) / data.get('فروش', 0)) * 100
                        ratios['حاشیه سود عملیاتی'] = round(operating_margin, 6)

                    # حاشیه سود خالص
                    if data.get('سود خالص', 0) != 0:
                        net_margin = (data.get('سود خالص', 0) / data.get('فروش', 0)) * 100
                        ratios['حاشیه سود خالص'] = round(net_margin, 6)

                # نسبت بدهی
                if data.get('کل دارایی ها', 0) > 0 and data.get('کل بدهی ها', 0) != 0:
                    debt_ratio = (data.get('کل بدهی ها', 0) / data.get('کل دارایی ها', 0)) * 100
                    ratios['نسبت بدهی'] = round(debt_ratio, 6)

                return ratios

            except Exception as e:
                print(f"خطا در محاسبه نسبت‌ها: {str(e)}")
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

                subheader_format = workbook.add_format({
                    'bold': True,
                    'align': 'center',
                    'valign': 'vcenter',
                    'bg_color': '#E6F1DC',
                    'border': 1
                })

                number_format = workbook.add_format({
                    'num_format': '#,##0',
                    'align': 'center',
                    'border': 1
                })

                ratio_format = workbook.add_format({
                    'num_format': '0.000000',
                    'align': 'center',
                    'border': 1
                })

                # داده‌های مالی
                financial_data = []
                ratios_data = []

                # تعریف متغیرها و نسبت‌ها
                variables = [
                    'دارایی جاری', 'کل دارایی ها', 'بدهی جاری', 'کل بدهی ها',
                    'فروش', 'سود ناخالص', 'سود عملیاتی', 'سود خالص',
                    'موجودی کالا', 'حساب های دریافتنی'
                ]

                ratios = [
                    'نسبت جاری', 'نسبت آنی', 'حاشیه سود ناخالص',
                    'حاشیه سود عملیاتی', 'حاشیه سود خالص', 'نسبت بدهی'
                ]

                years = ['1398', '1399', '1400', '1401', '1402']
                companies = sorted(results.keys())

                # ایجاد دیتافریم برای داده‌های مالی
                df_financial = pd.DataFrame(index=variables)
                df_ratios = pd.DataFrame(index=ratios)

                for company in companies:
                    company_data = {}
                    company_ratios = {}

                    for year in years:
                        if year in results[company]:
                            data = results[company][year]
                            col_name = f"{company} {year}"

                            # داده‌های مالی
                            company_data[col_name] = [data['متغیرها'].get(var, 0) for var in variables]

                            # نسبت‌های مالی
                            company_ratios[col_name] = [data['نسبت‌ها'].get(ratio, 0) for ratio in ratios]

                    # اضافه کردن داده‌ها به دیتافریم‌ها
                    df_temp = pd.DataFrame(company_data, index=variables)
                    df_financial = pd.concat([df_financial, df_temp], axis=1)

                    df_temp = pd.DataFrame(company_ratios, index=ratios)
                    df_ratios = pd.concat([df_ratios, df_temp], axis=1)

                # نوشتن داده‌های مالی
                df_financial.to_excel(writer, sheet_name='داده‌های مالی')
                worksheet1 = writer.sheets['داده‌های مالی']

                # نوشتن نسبت‌های مالی
                df_ratios.to_excel(writer, sheet_name='نسبت‌های مالی')
                worksheet2 = writer.sheets['نسبت‌های مالی']

                # تنظیم فرمت‌بندی برای هر دو شیت
                for ws, df, fmt in [(worksheet1, df_financial, number_format),
                                    (worksheet2, df_ratios, ratio_format)]:

                    # تنظیم عرض ستون‌ها
                    ws.set_column(0, 0, 20)  # ستون شاخص
                    ws.set_column(1, len(df.columns), 15)  # ستون‌های داده

                    # نوشتن هدر شرکت‌ها
                    for col, company in enumerate(companies):
                        ws.merge_range(0, col * 5 + 1, 0, col * 5 + 5, company, header_format)
                        for year_idx, year in enumerate(years):
                            ws.write(1, col * 5 + 1 + year_idx, year, subheader_format)

                    # نوشتن نام شاخص‌ها
                    for row, index in enumerate(df.index):
                        ws.write(row + 2, 0, index, header_format)

                    # نوشتن داده‌ها
                    for row in range(len(df.index)):
                        for col in range(len(df.columns)):
                            ws.write(row + 2, col + 1, df.iloc[row, col], fmt)

                    # تنظیم فریز پنل
                    ws.freeze_panes(2, 1)

                print(f"\nنتایج با موفقیت در فایل زیر ذخیره شد:\n{output_path}")
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