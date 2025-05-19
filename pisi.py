import os
import re
import warnings
from datetime import datetime
from decimal import (Decimal, getcontext, ROUND_HALF_UP,
                    InvalidOperation, DivisionByZero, localcontext)
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Union
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# تنظیمات اولیه
warnings.filterwarnings('ignore')
getcontext().prec = 28


class FinancialAnalyzer:
    def __init__(self, input_folder: str):
        """مقداردهی اولیه کلاس تحلیلگر مالی"""
        self.input_folder = Path(input_folder)
        self.current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.output_dir = self.input_folder / "Financial_Reports"
        self.output_dir.mkdir(exist_ok=True)

        # تعریف متغیرهای مالی اصلی
        self.financial_metrics = {
            "دارایی جاری": 0,
            "کل دارایی‌ها": 0,
            "بدهی جاری": 0,
            "کل بدهی‌ها": 0,
            "فروش": 0,
            "سود خالص": 0,
            "سود عملیاتی": 0,
            "سود ناخالص": 0,
            "موجودی کالا": 0,
            "حساب‌های دریافتنی": 0
        }

    def analyze_company(self, company_name: str) -> bool:
        """تحلیل اطلاعات مالی شرکت"""
        try:
            print(f"\nشروع تحلیل شرکت {company_name}")

            # خواندن داده‌های مالی
            if not self._read_financial_data():
                return False

            # محاسبه نسبت‌های مالی
            ratios = self._calculate_ratios()
            if not ratios:
                return False

            # نمایش نتایج
            self._display_results(ratios)

            # ذخیره گزارش
            if not self._save_report(company_name, ratios):
                return False

            return True

        except Exception as e:
            print(f"خطا در تحلیل شرکت: {str(e)}")
            return False

    def _read_financial_data(self) -> bool:
        """خواندن داده‌های مالی از فایل‌های اکسل"""
        try:
            # شبیه‌سازی خواندن داده‌ها
            self.financial_metrics.update({
                "دارایی جاری": 1500000000,
                "کل دارایی‌ها": 3000000000,
                "بدهی جاری": 800000000,
                "کل بدهی‌ها": 1200000000,
                "فروش": 2500000000,
                "سود خالص": 400000000,
                "سود عملیاتی": 500000000,
                "سود ناخالص": 800000000,
                "موجودی کالا": 400000000,
                "حساب‌های دریافتنی": 300000000
            })
            return True
        except Exception as e:
            print(f"خطا در خواندن داده‌های مالی: {str(e)}")
            return False

    def _calculate_ratios(self) -> dict:
        """محاسبه نسبت‌های مالی"""
        try:
            m = self.financial_metrics
            ratios = {
                "نسبت‌های نقدینگی": {
                    "نسبت جاری": self._safe_divide(m["دارایی جاری"], m["بدهی جاری"]),
                    "نسبت آنی": self._safe_divide(m["دارایی جاری"] - m["موجودی کالا"], m["بدهی جاری"])
                },
                "نسبت‌های سودآوری": {
                    "حاشیه سود ناخالص": self._safe_divide(m["سود ناخالص"], m["فروش"]),
                    "حاشیه سود عملیاتی": self._safe_divide(m["سود عملیاتی"], m["فروش"]),
                    "حاشیه سود خالص": self._safe_divide(m["سود خالص"], m["فروش"])
                },
                "نسبت‌های اهرمی": {
                    "نسبت بدهی": self._safe_divide(m["کل بدهی‌ها"], m["کل دارایی‌ها"]),
                    "پوشش بدهی": self._safe_divide(m["سود عملیاتی"], m["کل بدهی‌ها"])
                },
                "نسبت‌های فعالیت": {
                    "گردش حساب‌های دریافتنی": self._safe_divide(m["فروش"], m["حساب‌های دریافتنی"]),
                    "دوره وصول مطالبات": self._safe_divide(m["حساب‌های دریافتنی"] * 365, m["فروش"])
                }
            }
            return ratios
        except Exception as e:
            print(f"خطا در محاسبه نسبت‌ها: {str(e)}")
            return {}

    def _safe_divide(self, numerator: float, denominator: float) -> float:
        """تقسیم ایمن اعداد"""
        try:
            if denominator == 0:
                return 0
            return numerator / denominator
        except:
            return 0

    def _display_results(self, ratios: dict):
        """نمایش نتایج تحلیل"""
        print("\n=== نتایج تحلیل مالی ===")

        # نمایش متغیرهای اصلی
        print("\nمتغیرهای اصلی:")
        for metric, value in self.financial_metrics.items():
            print(f"{metric}: {self._format_number(value)} ریال")

        # نمایش نسبت‌ها
        print("\nنسبت‌های مالی:")
        for category, category_ratios in ratios.items():
            print(f"\n{category}:")
            for ratio_name, ratio_value in category_ratios.items():
                print(f"{ratio_name}: {ratio_value:.2f}")

    def _save_report(self, company_name: str, ratios: dict) -> bool:
        """ذخیره گزارش در فایل اکسل"""
        try:
            report_path = self.output_dir / f"گزارش_مالی_{company_name}_{self.current_time}.xlsx"

            # ایجاد دیتافریم‌ها برای ذخیره در اکسل
            metrics_df = pd.DataFrame({
                "متغیر": self.financial_metrics.keys(),
                "مقدار": self.financial_metrics.values()
            })

            ratios_data = []
            for category, category_ratios in ratios.items():
                for ratio_name, ratio_value in category_ratios.items():
                    ratios_data.append({
                        "دسته": category,
                        "نسبت": ratio_name,
                        "مقدار": ratio_value
                    })
            ratios_df = pd.DataFrame(ratios_data)

            # ذخیره در اکسل
            with pd.ExcelWriter(report_path) as writer:
                metrics_df.to_excel(writer, sheet_name='متغیرهای مالی', index=False)
                ratios_df.to_excel(writer, sheet_name='نسبت‌های مالی', index=False)

            print(f"\nگزارش در مسیر زیر ذخیره شد:\n{report_path}")
            return True

        except Exception as e:
            print(f"خطا در ذخیره گزارش: {str(e)}")
            return False

    @staticmethod
    def _format_number(number: float) -> str:
        """فرمت‌بندی اعداد با جداکننده هزار"""
        return f"{number:,.0f}"

    def main():
        try:
            # نمایش اطلاعات برنامه
            current_time = datetime.utcnow()
            print(
                f"Current Date and Time (UTC - YYYY-MM-DD HH:MM:SS formatted): {current_time.strftime('%Y-%m-%d %H:%M:%S')}")
            print(f"Current User's Login: {os.getenv('USERNAME', 'mmdura12')}")

            # دریافت مسیر
            while True:
                input_folder = input("\nلطفا مسیر پوشه حاوی فایل‌های اکسل را وارد کنید: ").strip()

                if not input_folder:
                    print("خطا: مسیر نمی‌تواند خالی باشد!")
                    continue

                if not os.path.exists(input_folder):
                    print("خطا: مسیر وارد شده وجود ندارد!")
                    if input("آیا می‌خواهید دوباره تلاش کنید؟ (بله/خیر) ").lower() != 'بله':
                        return
                    continue

                break

            # ایجاد آنالایزر
            analyzer = FinancialAnalyzer(input_folder)

            # دریافت و تحلیل اطلاعات شرکت
            company_name = input("\nلطفا نام شرکت را وارد کنید: ").strip()
            if not company_name:
                print("خطا: نام شرکت نمی‌تواند خالی باشد!")
                return

            if analyzer.analyze_company(company_name):
                print("\nتحلیل با موفقیت انجام شد.")
            else:
                print("\nخطا در انجام تحلیل!")

        except Exception as e:
            print(f"\nخطای غیرمنتظره: {str(e)}")
        finally:
            print("\nپایان برنامه")

    if __name__ == "__main__":
        main()