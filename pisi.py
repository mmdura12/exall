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
        """خواندن دقیق داده‌های مالی از فایل اکسل"""
        try:
            # خواندن فایل اکسل با تنظیمات خاص
            df = pd.read_excel(excel_file, header=None, dtype=str)
            year = str(excel_file).split('_')[0].split('\\')[-1]  # استخراج سال از نام فایل

            # تعریف کلمات کلیدی دقیق‌تر برای جستجو
            search_terms = {
                'دارایی جاری': [
                    'جمع دارایی‌های جاری',
                    'جمع داراییهای جاری',
                    'جمع دارایی های جاری',
                    'دارایی‌های جاری',
                    'داراییهای جاری',
                    'دارایی های جاری'
                ],
                'کل دارایی ها': [
                    'جمع کل دارایی‌ها',
                    'جمع دارایی‌ها',
                    'جمع داراییها',
                    'جمع دارایی ها',
                    'کل دارایی‌ها',
                    'کل داراییها'
                ],
                'بدهی جاری': [
                    'جمع بدهی‌های جاری',
                    'جمع بدهیهای جاری',
                    'جمع بدهی های جاری',
                    'بدهی‌های جاری',
                    'بدهیهای جاری'
                ],
                'کل بدهی ها': [
                    'جمع کل بدهی‌ها',
                    'جمع بدهی‌ها',
                    'جمع بدهیها',
                    'کل بدهی‌ها',
                    'کل بدهیها'
                ],
                'فروش': [
                    'درآمدهای عملیاتی',
                    'درآمد های عملیاتی',
                    'فروش خالص',
                    'فروش و درآمد ارائه خدمات',
                    'درآمد عملیاتی'
                ],
                'سود ناخالص': [
                    'سود (زیان) ناخالص',
                    'سود(زیان)ناخالص',
                    'سود/زیان ناخالص',
                    'سود وزیان ناخالص',
                    'سود ناخالص'
                ],
                'سود عملیاتی': [
                    'سود (زیان) عملیاتی',
                    'سود(زیان)عملیاتی',
                    'سود/زیان عملیاتی',
                    'سود وزیان عملیاتی',
                    'سود عملیاتی'
                ],
                'سود خالص': [
                    'سود (زیان) خالص',
                    'سود(زیان)خالص',
                    'سود/زیان خالص',
                    'سود وزیان خالص',
                    'سود خالص دوره'
                ],
                'موجودی کالا': [
                    'موجودی مواد و کالا',
                    'موجودی کالا و مواد',
                    'موجودی مواد، کالا و قطعات',
                    'موجودی کالا'
                ],
                'حساب های دریافتنی': [
                    'حساب‌های دریافتنی تجاری',
                    'حسابهای دریافتنی تجاری',
                    'حساب های دریافتنی تجاری',
                    'دریافتنی‌های تجاری',
                    'دریافتنیهای تجاری'
                ]
            }

            financial_data = {'سال': year}

            # جستجوی دقیق‌تر در فایل اکسل
            for metric, patterns in search_patterns.items():
                value = 0
                for pattern in patterns:
                    for idx, row in df.iterrows():
                        for col in df.columns:
                            cell_value = str(row[col]).strip()
                            if pattern.strip().replace(' ', '') in cell_value.replace(' ', ''):
                                # جستجو در ستون‌های بعدی برای یافتن مقدار عددی
                                for next_col in df.columns[col + 1:]:
                                    try:
                                        val = str(row[next_col]).strip()
                                        val = val.replace(',', '').replace('(', '-').replace(')', '')
                                        if val and val not in ['-', '0']:
                                            num = float(val)
                                            if num != 0:
                                                value = num
                                                break
                                    except:
                                        continue
                                if value != 0:
                                    break
                        if value != 0:
                            break
                    if value != 0:
                        break
                financial_data[metric] = value

            return financial_data

        except Exception as e:
            print(f"خطا در خواندن فایل {excel_file}: {e}")
            return None

    def find_value_in_excel(self, df, keywords):
        """جستجوی دقیق مقادیر در فایل اکسل"""
        for keyword in keywords:
            for i in range(len(df)):
                row = df.iloc[i]
                for col in df.columns:
                    cell_value = str(row[col]).strip()
                    if keyword.strip() in cell_value:
                        # جستجو در ستون‌های بعدی برای یافتن مقدار
                        for j in range(col + 1, len(df.columns)):
                            try:
                                next_value = str(df.iloc[i, j]).strip()
                                # حذف کاراکترهای اضافی
                                next_value = next_value.replace(',', '').replace('(', '-').replace(')', '')
                                if next_value and next_value != '-':
                                    value = float(next_value)
                                    if value != 0:
                                        return abs(value)  # برگرداندن مقدار مثبت
                            except:
                                continue

                        # جستجو در ستون‌های قبلی اگر مقداری پیدا نشد
                        for j in range(col - 1, -1, -1):
                            try:
                                prev_value = str(df.iloc[i, j]).strip()
                                prev_value = prev_value.replace(',', '').replace('(', '-').replace(')', '')
                                if prev_value and prev_value != '-':
                                    value = float(prev_value)
                                    if value != 0:
                                        return abs(value)
                            except:
                                continue
        return None

    def read_financial_data(self, excel_file):
        """خواندن و پردازش داده‌های مالی"""
        try:
            # خواندن فایل با تنظیمات مختلف
            df = pd.read_excel(excel_file, header=None, dtype=str)
            year = str(excel_file).split('_')[0].split('\\')[-1]

            financial_data = {'سال': year}
            found_values = {}

            # جستجوی مقادیر با روش‌های مختلف
            for metric, patterns in self.search_patterns.items():
                value = self.find_value_in_excel(df, patterns)
                if value is not None:
                    found_values[metric] = value
                else:
                    # جستجو با حالت‌های مختلف نوشتاری
                    modified_patterns = [p.replace('‌', ' ').replace('ی', 'ي') for p in patterns]
                    value = self.find_value_in_excel(df, modified_patterns)
                    if value is not None:
                        found_values[metric] = value

            # اطمینان از معتبر بودن داده‌ها
            if found_values:
                financial_data.update(found_values)
                print(f"\nداده‌های مالی یافت شده برای سال {year}:")
                for key, value in found_values.items():
                    print(f"{key}: {value:,.0f}")
                return financial_data
            else:
                print(f"هیچ داده مالی معتبری در فایل {excel_file.name} یافت نشد!")
                return None

        except Exception as e:
            print(f"خطا در پردازش فایل {excel_file.name}: {str(e)}")
            return None

        def calculate_financial_ratios(self, data):
            """محاسبه نسبت‌های مالی با دقت بالا"""
            ratios = {}

            try:
                # نسبت‌های نقدینگی
                if data.get('بدهی جاری', 0) > 0:
                    ratios['نسبت جاری'] = data.get('دارایی جاری', 0) / data['بدهی جاری']
                    ratios['نسبت آنی'] = (data.get('دارایی جاری', 0) - data.get('موجودی کالا', 0)) / data['بدهی جاری']

                # نسبت‌های سودآوری
                if data.get('فروش', 0) > 0:
                    ratios['حاشیه سود ناخالص'] = (data.get('سود ناخالص', 0) / data['فروش']) * 100
                    ratios['حاشیه سود عملیاتی'] = (data.get('سود عملیاتی', 0) / data['فروش']) * 100
                    ratios['حاشیه سود خالص'] = (data.get('سود خالص', 0) / data['فروش']) * 100

                # نسبت‌های اهرمی
                if data.get('کل دارایی ها', 0) > 0:
                    ratios['نسبت بدهی'] = (data.get('کل بدهی ها', 0) / data['کل دارایی ها']) * 100

                return ratios

            except Exception as e:
                print(f"خطا در محاسبه نسبت‌ها: {str(e)}")
                return {}

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

                # ذخیره گزارش جامع
            if all_results:
                self.save_comprehensive_report(all_results)

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

    def save_comprehensive_report(self, all_results):
        """ذخیره همه نتایج در یک فایل اکسل جامع"""
        try:
            filename = self.output_folder / f"تحلیل_جامع_شرکت‌ها_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

            with pd.ExcelWriter(filename) as writer:
                # شیت داده‌های مالی
                financial_data = []
                for company, years_data in all_results.items():
                    for year, data in years_data.items():
                        row_data = {
                            'نام شرکت': company,
                            'سال': year
                        }
                        row_data.update(data['financial_data'])
                        financial_data.append(row_data)

                df_financial = pd.DataFrame(financial_data)
                df_financial.to_excel(writer, sheet_name='داده‌های مالی', index=False)

                # شیت نسبت‌ها
                ratios_data = []
                for company, years_data in all_results.items():
                    for year, data in years_data.items():
                        if data['ratios']:
                            row_data = {
                                'نام شرکت': company,
                                'سال': year
                            }
                            row_data.update(data['ratios'])
                            ratios_data.append(row_data)

                df_ratios = pd.DataFrame(ratios_data)
                df_ratios.to_excel(writer, sheet_name='نسبت‌های مالی', index=False)

                # شیت تحلیل روند
                trend_analysis = []
                for company in all_results.keys():
                    company_years = sorted(all_results[company].keys())
                    if len(company_years) >= 2:
                        latest_year = company_years[-1]
                        previous_year = company_years[-2]

                        latest_data = all_results[company][latest_year]['financial_data']
                        previous_data = all_results[company][previous_year]['financial_data']

                        for metric in latest_data.keys():
                            if metric != 'سال':
                                try:
                                    change = ((latest_data[metric] - previous_data[metric]) /
                                              previous_data[metric] * 100 if previous_data[metric] != 0 else 0)
                                    trend_analysis.append({
                                        'نام شرکت': company,
                                        'شاخص': metric,
                                        f'مقدار {latest_year}': latest_data[metric],
                                        f'مقدار {previous_year}': previous_data[metric],
                                        'درصد تغییر': change
                                    })
                                except:
                                    continue

                df_trend = pd.DataFrame(trend_analysis)
                df_trend.to_excel(writer, sheet_name='تحلیل روند', index=False)

                # تنظیم عرض ستون‌ها
                for sheet_name in writer.sheets:
                    worksheet = writer.sheets[sheet_name]
                    for idx, col in enumerate(df_financial.columns):
                        max_length = max(
                            df_financial[col].astype(str).apply(len).max(),
                            len(str(col))
                        ) + 2
                        worksheet.set_column(idx, idx, max_length)

            print(f"\nگزارش جامع در مسیر زیر ذخیره شد:\n{filename}")
            return True

        except Exception as e:
            print(f"خطا در ذخیره گزارش جامع: {e}")
            return False


def main():
    try:
        print("\n=== سیستم محاسبه متغیرها و نسبت‌های مالی ===")
        print(f"زمان اجرا: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"کاربر: {os.getenv('USERNAME', 'unknown')}")
        print("=" * 50)

        # دریافت مسیر
        folder_path = input("\nلطفاً مسیر پوشه حاوی فایل‌های اکسل را وارد کنید: ").strip()
        if not os.path.exists(folder_path):
            print("خطا: مسیر وارد شده وجود ندارد!")
            return

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
            print("هیچ شرکتی برای محاسبه وارد نشده است!")
            return

        # پردازش هر شرکت
        all_data = {}
        for company in companies:
            print(f"\nمحاسبه اطلاعات شرکت {company}:")

            # یافتن فایل‌های شرکت
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
                    ratios = analyzer.calculate_financial_ratios(data)
                    company_data[year] = {
                        'متغیرها': data,
                        'نسبت‌ها': ratios
                    }

            if company_data:
                all_data[company] = company_data

        # ذخیره نتایج در اکسل
        if all_data:
            output_file = Path(folder_path) / f"نتایج_محاسبات_مالی_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

            with pd.ExcelWriter(output_file) as writer:
                # شیت متغیرها
                variables_data = []
                for company, years in all_data.items():
                    for year, data in years.items():
                        row = {'شرکت': company, 'سال': year}
                        row.update(data['متغیرها'])
                        variables_data.append(row)

                pd.DataFrame(variables_data).to_excel(writer, sheet_name='متغیرهای مالی', index=False)

                # شیت نسبت‌ها
                ratios_data = []
                for company, years in all_data.items():
                    for year, data in years.items():
                        if data['نسبت‌ها']:
                            row = {'شرکت': company, 'سال': year}
                            row.update(data['نسبت‌ها'])
                            ratios_data.append(row)

                pd.DataFrame(ratios_data).to_excel(writer, sheet_name='نسبت‌های مالی', index=False)

            print(f"\nنتایج محاسبات در فایل زیر ذخیره شد:")
            print(output_file)
        else:
            print("\nهیچ داده‌ای برای ذخیره‌سازی یافت نشد!")

    except Exception as e:
        print(f"\nخطای غیرمنتظره: {str(e)}")
    finally:
        print("\nپایان برنامه")


if __name__ == "__main__":
    main()