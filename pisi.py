import os
import pandas as pd
import warnings
from datetime import datetime
from pathlib import Path

warnings.filterwarnings('ignore')


class FinancialAnalyzer:
    def __init__(self, base_folder):
        self.base_folder = Path(base_folder)
        self.output_folder = self.base_folder / 'reports'
        self.output_folder.mkdir(exist_ok=True)

    def find_company_files(self, company_name):
        """یافتن فایل‌های مربوط به یک شرکت"""
        files = list(self.base_folder.glob(f'*{company_name}*.xlsx'))
        return sorted(files)

    def read_financial_data(self, excel_file):
        """خواندن اطلاعات مالی از فایل اکسل"""
        try:
            df = pd.read_excel(excel_file, engine='openpyxl')

            # جستجوی مقادیر با استفاده از کلمات کلیدی
            financial_data = {
                'سال': str(excel_file).split('_')[-1].split('.')[0],  # استخراج سال از نام فایل
                'دارایی جاری': self.find_value(df, ['دارایی جاری', 'دارایی های جاری']),
                'کل دارایی ها': self.find_value(df, ['جمع دارایی', 'کل دارایی']),
                'بدهی جاری': self.find_value(df, ['بدهی جاری', 'بدهی های جاری']),
                'کل بدهی ها': self.find_value(df, ['جمع بدهی', 'کل بدهی']),
                'فروش': self.find_value(df, ['فروش خالص', 'درآمد عملیاتی']),
                'سود ناخالص': self.find_value(df, ['سود ناخالص']),
                'سود عملیاتی': self.find_value(df, ['سود عملیاتی']),
                'سود خالص': self.find_value(df, ['سود خالص']),
                'موجودی کالا': self.find_value(df, ['موجودی کالا', 'موجودی مواد و کالا']),
                'حساب های دریافتنی': self.find_value(df, ['حساب های دریافتنی تجاری'])
            }
            return financial_data
        except Exception as e:
            print(f"خطا در خواندن فایل {excel_file.name}: {e}")
            return None

    def find_value(self, df, keywords):
        """یافتن مقدار با استفاده از کلمات کلیدی"""
        for keyword in keywords:
            try:
                # جستجو در ستون‌ها
                for col in df.columns:
                    mask = df[col].astype(str).str.contains(keyword, case=False, na=False)
                    if mask.any():
                        row = df.loc[mask].iloc[0]
                        # جستجوی عدد در سطر
                        for col_name in df.columns:
                            try:
                                value = pd.to_numeric(row[col_name])
                                if not pd.isna(value) and value != 0:
                                    return value
                            except:
                                continue
            except:
                continue
        return 0

    def calculate_ratios(self, data):
        """محاسبه نسبت‌های مالی"""
        try:
            ratios = {}

            # نسبت‌های نقدینگی
            if data['بدهی جاری'] != 0:
                ratios['نسبت جاری'] = data['دارایی جاری'] / data['بدهی جاری']
                ratios['نسبت آنی'] = (data['دارایی جاری'] - data['موجودی کالا']) / data['بدهی جاری']

            # نسبت‌های سودآوری
            if data['فروش'] != 0:
                ratios['حاشیه سود ناخالص'] = (data['سود ناخالص'] / data['فروش']) * 100
                ratios['حاشیه سود عملیاتی'] = (data['سود عملیاتی'] / data['فروش']) * 100
                ratios['حاشیه سود خالص'] = (data['سود خالص'] / data['فروش']) * 100

            # نسبت‌های اهرمی
            if data['کل دارایی ها'] != 0:
                ratios['نسبت بدهی'] = (data['کل بدهی ها'] / data['کل دارایی ها']) * 100

            return ratios
        except Exception as e:
            print(f"خطا در محاسبه نسبت‌ها: {e}")
            return None

    def analyze_companies(self, company_names):
        """تحلیل چند شرکت"""
        all_results = {}

        for company_name in company_names:
            print(f"\nتحلیل شرکت {company_name}:")
            company_files = self.find_company_files(company_name)

            if not company_files:
                print(f"هیچ فایلی برای شرکت {company_name} یافت نشد!")
                continue

            company_data = {}
            for file in company_files:
                financial_data = self.read_financial_data(file)
                if financial_data:
                    year = financial_data['سال']
                    ratios = self.calculate_ratios(financial_data)
                    company_data[year] = {
                        'financial_data': financial_data,
                        'ratios': ratios
                    }

            if company_data:
                all_results[company_name] = company_data
                self.display_company_results(company_name, company_data)
                self.save_company_report(company_name, company_data)

        return all_results

    def display_company_results(self, company_name, data):
        """نمایش نتایج تحلیل یک شرکت"""
        print(f"\n=== نتایج تحلیل شرکت {company_name} ===")

        for year, year_data in sorted(data.items()):
            print(f"\nسال {year}:")
            print("اطلاعات مالی اصلی:")
            for key, value in year_data['financial_data'].items():
                if key != 'سال':
                    print(f"{key}: {value:,.0f}")

            print("\nنسبت‌های مالی:")
            for key, value in year_data['ratios'].items():
                print(f"{key}: {value:.2f}")

    def save_company_report(self, company_name, data):
        """ذخیره گزارش در اکسل"""
        try:
            filename = self.output_folder / f"{company_name}_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

            with pd.ExcelWriter(filename) as writer:
                # تبدیل داده‌های مالی به دیتافریم
                financial_data = []
                ratios_data = []

                for year, year_data in data.items():
                    # داده‌های مالی
                    for key, value in year_data['financial_data'].items():
                        if key != 'سال':
                            financial_data.append({
                                'سال': year,
                                'شاخص': key,
                                'مقدار': value
                            })

                    # نسبت‌ها
                    for key, value in year_data['ratios'].items():
                        ratios_data.append({
                            'سال': year,
                            'نسبت': key,
                            'مقدار': value
                        })

                # ذخیره در شیت‌های جداگانه
                pd.DataFrame(financial_data).to_excel(writer, sheet_name='اطلاعات مالی', index=False)
                pd.DataFrame(ratios_data).to_excel(writer, sheet_name='نسبت‌های مالی', index=False)

            print(f"\nگزارش در مسیر زیر ذخیره شد:\n{filename}")

        except Exception as e:
            print(f"خطا در ذخیره گزارش: {e}")


def main():
    try:
        print("\n=== سیستم تحلیل مالی شرکت‌ها ===")
        print(f"زمان اجرا: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

        folder_path = input("\nلطفاً مسیر پوشه حاوی فایل‌های اکسل را وارد کنید: ").strip()
        if not os.path.exists(folder_path):
            print("خطا: مسیر وارد شده وجود ندارد!")
            return

        analyzer = FinancialAnalyzer(folder_path)

        # دریافت نام شرکت‌ها
        companies = []
        print("\nلطفاً نام 5 شرکت را وارد کنید (برای پایان، Enter خالی بزنید):")
        while len(companies) < 5:
            company = input(f"نام شرکت {len(companies) + 1}: ").strip()
            if not company:
                break
            companies.append(company)

        if companies:
            results = analyzer.analyze_companies(companies)
            if results:
                print("\nتحلیل با موفقیت انجام شد.")
            else:
                print("\nخطا در انجام تحلیل!")

    except Exception as e:
        print(f"\nخطای غیرمنتظره: {e}")
    finally:
        print("\nپایان برنامه")


if __name__ == "__main__":
    main()