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
                'داراییهای جاری', 'دارایی های جاری', 'جمع دارایی های جاری'
            ],
            'کل دارایی ها': [
                'جمع دارایی‌ها', 'جمع داراییها', 'جمع کل دارایی‌ها',
                'کل دارایی‌ها', 'دارایی ها', 'جمع دارایی ها'
            ],
            'بدهی جاری': [
                'جمع بدهی‌های جاری', 'جمع بدهیهای جاری', 'بدهی‌های جاری',
                'بدهیهای جاری', 'بدهی های جاری', 'جمع بدهی های جاری'
            ],
            'کل بدهی ها': [
                'جمع بدهی‌ها', 'جمع بدهیها', 'جمع کل بدهی‌ها',
                'کل بدهی‌ها', 'بدهی ها', 'جمع بدهی ها'
            ],
            'فروش': [
                'درآمدهای عملیاتی', 'درآمد عملیاتی', 'فروش خالص',
                'فروش', 'درآمد حاصل از فروش', 'فروش و درآمد ارائه خدمات'
            ],
            'سود ناخالص': [
                'سود ناخالص', 'سود (زیان) ناخالص', 'سود و زیان ناخالص'
            ],
            'سود عملیاتی': [
                'سود عملیاتی', 'سود (زیان) عملیاتی', 'سود و زیان عملیاتی'
            ],
            'سود خالص': [
                'سود خالص', 'سود (زیان) خالص', 'سود خالص دوره'
            ],
            'موجودی کالا': [
                'موجودی مواد و کالا', 'موجودی کالا', 'موجودی‌های مواد و کالا'
            ],
            'حساب های دریافتنی': [
                'حساب‌های دریافتنی تجاری', 'حسابهای دریافتنی', 'دریافتنی‌های تجاری'
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
        """جستجوی مقدار در دیتافریم"""
        for keyword in keywords:
            for idx, row in df.iterrows():
                for col in df.columns:
                    cell_value = str(row[col]).strip()
                    if keyword in cell_value:
                        # جستجو در ستون‌های بعدی
                        for next_col in df.columns[df.columns.get_loc(col) + 1:]:
                            value = self.clean_number(row[next_col])
                            if value != 0:
                                return value

                        # جستجو در ستون‌های قبلی
                        for prev_col in df.columns[:df.columns.get_loc(col)]:
                            value = self.clean_number(row[prev_col])
                            if value != 0:
                                return value
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

    def calculate_ratios(self, data):
        """محاسبه نسبت‌های مالی"""
        ratios = {}

        if data.get('بدهی جاری', 0) > 0:
            ratios['نسبت جاری'] = data.get('دارایی جاری', 0) / data['بدهی جاری']
            ratios['نسبت آنی'] = (data.get('دارایی جاری', 0) - data.get('موجودی کالا', 0)) / data['بدهی جاری']

        if data.get('فروش', 0) > 0:
            ratios['حاشیه سود ناخالص'] = (data.get('سود ناخالص', 0) / data['فروش']) * 100
            ratios['حاشیه سود عملیاتی'] = (data.get('سود عملیاتی', 0) / data['فروش']) * 100
            ratios['حاشیه سود خالص'] = (data.get('سود خالص', 0) / data['فروش']) * 100

        if data.get('کل دارایی ها', 0) > 0:
            ratios['نسبت بدهی'] = (data.get('کل بدهی ها', 0) / data['کل دارایی ها']) * 100

        return ratios

    def save_results(self, results, output_path):
        """ذخیره نتایج در اکسل"""
        try:
            with pd.ExcelWriter(output_path) as writer:
                # داده‌های مالی
                financial_data = []
                for company, years in results.items():
                    for year, data in years.items():
                        row = {'شرکت': company, 'سال': year}
                        row.update(data['متغیرها'])
                        financial_data.append(row)

                df_financial = pd.DataFrame(financial_data)
                df_financial.to_excel(writer, sheet_name='داده‌های مالی', index=False)

                # نسبت‌های مالی
                ratios_data = []
                for company, years in results.items():
                    for year, data in years.items():
                        if data['نسبت‌ها']:
                            row = {'شرکت': company, 'سال': year}
                            row.update(data['نسبت‌ها'])
                            ratios_data.append(row)

                if ratios_data:
                    pd.DataFrame(ratios_data).to_excel(writer, sheet_name='نسبت‌های مالی', index=False)

            print(f"\nنتایج در فایل زیر ذخیره شد:\n{output_path}")
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