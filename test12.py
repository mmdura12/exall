from pathlib import Path
import pandas as pd
from typing import Dict, Tuple, Optional
from helper_functions import clean_persian_text, convert_to_number
from financial_ratios import FinancialRatioCalculator
from decimal import Decimal, ROUND_HALF_UP
import pandas as pd
import re
from decimal import Decimal
from typing import Dict, Optional
from helper_functions import safe_divide
import sys
import logging
from datetime import datetime
from pathlib import Path
from typing import Optional
import pandas as pd
from financial_analyzer import FinancialAnalyzer

class FinancialAnalyzer:
    def __init__(self, input_folder: str):
        self.input_folder = Path(input_folder)
        self.ratio_calculator = FinancialRatioCalculator()

        # عبارات جستجو برای متغیرها
        self.search_terms = {
            "وجه نقد": ["موجودی نقد", "وجه نقد", "موجودی نقد و بانک"],
            "حساب دریافتنی": ["حسابهای دریافتنی تجاری", "دریافتنی‌های تجاری"],
            "موجودی کالا": ["موجودی مواد و کالا", "موجودی کالا", "موجودی‌ها"],
            # ... سایر متغیرها
        }

    def find_value_in_df(self, df: pd.DataFrame, search_terms: list) -> Decimal:
        """
        یافتن مقدار در دیتافریم با استفاده از عبارات جستجو
        """
        try:
            for term in search_terms:
                term = clean_persian_text(term)
                for idx, row in df.iterrows():
                    for col in df.columns:
                        cell_value = clean_persian_text(str(row[col]))
                        if term in cell_value:
                            for value_col in df.columns:
                                value = str(row[value_col])
                                number = convert_to_number(value)
                                if number != 0:
                                    return number
            return Decimal('0')

        except Exception as e:
            print(f"خطا در جستجوی مقدار: {str(e)}")
            return Decimal('0')

    def process_file(self, file_path: Path) -> Tuple[Dict, Dict]:
        """
        پردازش فایل اکسل و استخراج داده‌ها
        """
        try:
            print(f"\nدر حال پردازش فایل: {file_path.name}")

            df = pd.read_excel(
                file_path,
                header=None,
                dtype=str,
                na_filter=False
            )

            # استخراج متغیرها
            variables = {}
            for var_name, search_terms in self.search_terms.items():
                value = self.find_value_in_df(df, search_terms)
                variables[var_name] = value
                print(f"{var_name}: {float(value):,.0f}")

            # محاسبه نسبت‌ها
            ratios = self.ratio_calculator.calculate_all_ratios(variables)

            # نمایش نسبت‌ها
            print("\nنسبت‌های مالی محاسبه شده:")
            for ratio_name, ratio_value in ratios.items():
                print(f"{ratio_name}: {float(ratio_value):.2f}")

            return variables, ratios

        except Exception as e:
            print(f"خطا در پردازش فایل {file_path.name}: {str(e)}")
            return {}, {}

        def safe_divide(numerator: Decimal, denominator: Decimal) -> Decimal:
            """
            تقسیم ایمن دو عدد با در نظر گرفتن حالت تقسیم بر صفر
            """
            try:
                if numerator is None or denominator is None:
                    return Decimal('0')

                numerator = Decimal(str(numerator))
                denominator = Decimal(str(denominator))

                if denominator == 0:
                    return Decimal('0')

                result = numerator / denominator
                return result.quantize(Decimal('0.0001'), rounding=ROUND_HALF_UP)

            except Exception as e:
                print(f"خطا در تقسیم: {str(e)}")
                return Decimal('0')

        def clean_persian_text(text: str) -> str:
            """
            پاکسازی و یکسان‌سازی متن فارسی
            """
            if not isinstance(text, str):
                return ''

            # حذف فاصله‌های اضافی
            text = re.sub(r'\s+', ' ', text)

            # یکسان‌سازی کاراکترها
            replacements = {
                'ي': 'ی', 'ك': 'ک',
                '٠': '۰', '١': '۱', '٢': '۲', '٣': '۳', '٤': '۴',
                '٥': '۵', '٦': '۶', '٧': '۷', '٨': '۸', '٩': '۹'
            }

            for old, new in replacements.items():
                text = text.replace(old, new)

            return text.strip()

        def convert_to_number(value: str) -> Decimal:
            """
            تبدیل متن به عدد با پشتیبانی از فرمت‌های مختلف
            """
            try:
                if pd.isna(value) or not str(value).strip():
                    return Decimal('0')

                value = str(value)

                # تبدیل اعداد فارسی به انگلیسی
                persian_nums = '۰۱۲۳۴۵۶۷۸۹'
                english_nums = '0123456789'
                for persian, english in zip(persian_nums, english_nums):
                    value = value.replace(persian, english)

                # پاکسازی کاراکترهای غیرعددی
                value = value.replace(',', '').replace('٬', '')
                value = value.replace('(', '-').replace(')', '')
                value = value.replace('−', '-').replace('–', '-')

                # فقط نگه داشتن اعداد و علامت‌های خاص
                value = ''.join(c for c in value if c.isdigit() or c in '.-')

                return Decimal(value) if value else Decimal('0')

            except Exception as e:
                print(f"خطا در تبدیل مقدار {value} به عدد: {str(e)}")
                return Decimal('0')


class FinancialRatioCalculator:
    """
    کلاس محاسبه‌کننده نسبت‌های مالی
    """

    @staticmethod
    def calculate_liquidity_ratios(data: Dict[str, Decimal]) -> Dict[str, Decimal]:
        """
        محاسبه نسبت‌های نقدینگی
        """
        return {
            "نسبت جاری": safe_divide(
                data["دارایی جاری"],
                data["بدهی جاری"]
            ),
            "نسبت آنی": safe_divide(
                data["دارایی جاری"] - data["موجودی کالا"],
                data["بدهی جاری"]
            ),
            "نسبت وجه نقد": safe_divide(
                data["وجه نقد"],
                data["بدهی جاری"]
            )
        }

    @staticmethod
    def calculate_profitability_ratios(data: Dict[str, Decimal]) -> Dict[str, Decimal]:
        """
        محاسبه نسبت‌های سودآوری
        """
        return {
            "بازده دارایی ها": safe_divide(
                data["سود خالص"],
                data["کل دارایی ها"]
            ) * Decimal('100'),
            "بازده حقوق صاحبان سهام": safe_divide(
                data["سود خالص"],
                data["حقوق صاحبان سهام"]
            ) * Decimal('100'),
            "حاشیه سود خالص": safe_divide(
                data["سود خالص"],
                data["فروش"]
            ) * Decimal('100'),
            "حاشیه سود عملیاتی": safe_divide(
                data["سود عملیاتی"],
                data["فروش"]
            ) * Decimal('100'),
            "حاشیه سود ناخالص": safe_divide(
                data["سود ناخالص"],
                data["فروش"]
            ) * Decimal('100')
        }

    @staticmethod
    def calculate_activity_ratios(data: Dict[str, Decimal]) -> Dict[str, Decimal]:
        """
        محاسبه نسبت‌های فعالیت
        """
        return {
            "دوره وصول مطالبات": safe_divide(
                data["حساب دریافتنی"] * Decimal('365'),
                data["فروش"]
            ),
            "گردش حساب دریافتنی": safe_divide(
                data["فروش"],
                data["حساب دریافتنی"]
            ),
            "گردش موجودی کالا": safe_divide(
                data["بهای تمام شده کالای فروش رفته"],
                data["موجودی کالا"]
            )
        }

    @staticmethod
    def calculate_leverage_ratios(data: Dict[str, Decimal]) -> Dict[str, Decimal]:
        """
        محاسبه نسبت‌های اهرمی
        """
        return {
            "نسبت بدهی به دارایی": safe_divide(
                data["کل بدهی ها"],
                data["کل دارایی ها"]
            ) * Decimal('100')
        }

    def calculate_all_ratios(self, data: Dict[str, Decimal]) -> Dict[str, Decimal]:
        """
        محاسبه همه نسبت‌های مالی
        """
        all_ratios = {}
        all_ratios.update(self.calculate_liquidity_ratios(data))
        all_ratios.update(self.calculate_profitability_ratios(data))
        all_ratios.update(self.calculate_activity_ratios(data))
        all_ratios.update(self.calculate_leverage_ratios(data))
        return all_ratios

    def setup_logging(output_dir: Path) -> None:
        """
        راه‌اندازی سیستم ثبت لاگ
        """
        log_file = output_dir / f"financial_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file, encoding='utf-8'),
                logging.StreamHandler(sys.stdout)
            ]
        )

    def get_valid_path(prompt: str) -> Optional[Path]:
        """
        دریافت و اعتبارسنجی مسیر ورودی
        """
        try:
            while True:
                path_str = input(prompt).strip()

                if not path_str:
                    logging.error("مسیر نمی‌تواند خالی باشد!")
                    continue

                path = Path(path_str)

                if not path.exists():
                    logging.error("مسیر وارد شده وجود ندارد!")
                    retry = input("آیا می‌خواهید دوباره تلاش کنید؟ (بله/خیر) ").strip().lower()
                    if retry != 'بله':
                        return None
                    continue

                return path

        except Exception as e:
            logging.error(f"خطا در دریافت مسیر: {str(e)}")
            return None

    def save_summary_report(analyzer: FinancialAnalyzer, all_results: dict, company_name: str) -> None:
        """
        ذخیره گزارش خلاصه تحلیل
        """
        try:
            summary_file = analyzer.output_dir / f"{company_name}_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"

            with open(summary_file, 'w', encoding='utf-8') as f:
                f.write(f"گزارش تحلیل مالی شرکت {company_name}\n")
                f.write(f"تاریخ تحلیل: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"تحلیلگر: {sys.argv[1] if len(sys.argv) > 1 else 'mmdura12'}\n")
                f.write("-" * 50 + "\n\n")

                for year, data in sorted(all_results['ratios'].items()):
                    f.write(f"\nسال {year}:\n")
                    f.write("-" * 20 + "\n")

                    # نمایش نسبت‌های کلیدی
                    key_ratios = [
                        "نسبت جاری",
                        "نسبت آنی",
                        "بازده دارایی ها",
                        "بازده حقوق صاحبان سهام",
                        "حاشیه سود خالص"
                    ]

                    for ratio in key_ratios:
                        if ratio in data:
                            f.write(f"{ratio}: {float(data[ratio]):.2f}\n")

                f.write("\n" + "-" * 50 + "\n")
                f.write("پایان گزارش")

            logging.info(f"گزارش خلاصه در مسیر زیر ذخیره شد:\n{summary_file}")

        except Exception as e:
            logging.error(f"خطا در ذخیره گزارش خلاصه: {str(e)}")

    def main() -> None:
        """
        تابع اصلی برنامه
        """
        try:
            print("\n=== سیستم تحلیل مالی ===")
            print(f"تاریخ و زمان: {datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')} UTC")
            print(f"کاربر: {sys.argv[1] if len(sys.argv) > 1 else 'mmdura12'}")
            print("-" * 30 + "\n")

            # دریافت مسیر پوشه ورودی
            input_folder = get_valid_path("لطفا مسیر پوشه حاوی فایل‌های اکسل را وارد کنید: ")
            if not input_folder:
                logging.error("برنامه به دلیل عدم دریافت مسیر معتبر خاتمه می‌یابد.")
                return

            # ایجاد پوشه خروجی و راه‌اندازی لاگینگ
            output_dir = input_folder / "Financial_Reports"
            output_dir.mkdir(exist_ok=True)
            setup_logging(output_dir)

            # ایجاد آنالایزر
            analyzer = FinancialAnalyzer(str(input_folder))
            logging.info(f"آنالایزر با مسیر ورودی {input_folder} ایجاد شد.")

            # دریافت نام شرکت
            while True:
                company_name = input("نام شرکت را وارد کنید: ").strip()
                if company_name:
                    break
                print("نام شرکت نمی‌تواند خالی باشد!")

            # تحلیل شرکت
            logging.info(f"شروع تحلیل شرکت {company_name}")
            results = analyzer.analyze_company(company_name)

            if results and results.get('variables') and results.get('ratios'):
                logging.info("تحلیل با موفقیت انجام شد!")

                # ذخیره نتایج
                excel_file = analyzer.create_excel_report(results, company_name)
                charts_file = analyzer.create_charts(results, company_name)
                save_summary_report(analyzer, results, company_name)

                print("\n=== نتایج تحلیل ===")
                print(f"1. گزارش اکسل: {excel_file}")
                print(f"2. نمودارها: {charts_file}")
                print(f"3. مسیر خروجی: {output_dir}")
                print("\nلطفا فایل‌های خروجی را بررسی کنید.")

            else:
                logging.error(f"خطا در تحلیل شرکت {company_name}")
                print("\nمتاسفانه تحلیل با خطا مواجه شد. لطفا لاگ‌ها را بررسی کنید.")

        except KeyboardInterrupt:
            print("\n\nبرنامه توسط کاربر متوقف شد.")
            logging.info("برنامه توسط کاربر متوقف شد.")

        except Exception as e:
            logging.error(f"خطای کلی در اجرای برنامه: {str(e)}")
            print("\nخطایی در اجرای برنامه رخ داد. لطفا لاگ‌ها را بررسی کنید.")

        finally:
            print("\n=== پایان برنامه ===")

    if __name__ == "__main__":
        main()