import numpy as np
import logging
import cv2
from scipy.ndimage import gaussian_filter
from skimage.feature import peak_local_max

logger = logging.getLogger(__name__)


class PrecisePersonDetector:
    def __init__(self):
        # پارامترهای بهینه شده برای کار با مدل‌های crowd counting مانند CSRNet
        self.min_distance = 2           # فاصله حداقل بین نقاط تشخیص
        self.gaussian_sigma = 1.3       # سیگمای گاوسی برای صاف کردن ملایم اما نه زیاد
        self.density_multiplier = 330   # ضریب تبدیل مجموع تراکم به تعداد نفرات (مخصوص CSRNet)
        self.min_peak_threshold = 0.05  # مقدار حداقلی برای آستانه نقاط حداکثر محلی
        logger.info("PrecisePersonDetector initialized with tuned params")

    def normalize_density(self, density_map):
        """نرمال‌سازی نقشه تراکم به بازه 0 تا 1"""
        if density_map.max() > density_map.min():
            return (density_map - density_map.min()) / (density_map.max() - density_map.min())
        return density_map

    def enhance_density_map(self, density_map):
        """بهبود نقشه تراکم برای تشخیص بهتر نقاط"""
        try:
            # نرمال‌سازی اولیه
            density_norm = self.normalize_density(density_map)

            # تقویت کنتراست با CLAHE
            density_uint8 = (density_norm * 255).astype(np.uint8)
            clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
            enhanced = clahe.apply(density_uint8)
            enhanced_float = enhanced.astype(np.float32) / 255.0

            # صاف‌سازی با فیلتر گاوسی (کاهش نویز و نوسان)
            smoothed = gaussian_filter(enhanced_float, sigma=self.gaussian_sigma)

            # افزایش وضوح لبه‌ها برای کمک به تشخیص نقاط
            kernel = np.array([[-1, -1, -1], [-1, 9, -1], [-1, -1, -1]])
            sharpened = cv2.filter2D(smoothed, -1, kernel)

            # نرمال‌سازی نهایی
            final = np.clip(sharpened, 0, 1)
            return final
        except Exception as e:
            logger.error(f"Error in density map enhancement: {str(e)}")
            return density_map

    def remove_nearby_points(self, points, min_distance=2):
        """
        حذف نقاط خیلی نزدیک به هم (برای جلوگیری از شمردن چندباره)
        """
        if len(points) == 0:
            return np.empty((0, 2), dtype=int)
        final_points = [points[0]]
        for point in points[1:]:
            distances = np.sqrt(np.sum((np.array(final_points) - point) ** 2, axis=1))
            if np.all(distances >= min_distance):
                final_points.append(point)
        return np.array(final_points)

    def detect_local_maxima(self, density_map):
        try:
            # بهبود و پیش‌پردازش نقشه تراکم
            processed_map = self.enhance_density_map(density_map)

            # تخمین اولیه تعداد نفرات براساس مدل
            total_density = np.sum(density_map)
            estimated_count = int(total_density * self.density_multiplier)
            logger.info(f"Total density: {total_density:.3f} | Initial density-based estimate: {estimated_count}")

            # محاسبه threshold پویا براساس توزیع مقادیر ملایم نقشه
            if np.any(processed_map > 0):
                dynamic_thresh = max(self.min_peak_threshold, np.percentile(processed_map[processed_map > 0], 70))
            else:
                dynamic_thresh = self.min_peak_threshold

            # نقاط ماکزیمم محلی
            coordinates = peak_local_max(
                processed_map,
                min_distance=self.min_distance,
                threshold_abs=dynamic_thresh,
                exclude_border=False
            )

            # اگر کم بود، threshold را پایین‌تر می‌آوریم
            if len(coordinates) < estimated_count * 0.5 and estimated_count > 0:
                alt_thresh = max(self.min_peak_threshold / 2, np.percentile(processed_map[processed_map > 0], 50))
                extra = peak_local_max(
                    processed_map,
                    min_distance=1,
                    threshold_abs=alt_thresh,
                    exclude_border=False
                )
                if len(extra) > 0:
                    coordinates = np.vstack([coordinates, extra])

            # حذف نقاط تکراری نزدیک به هم
            if len(coordinates) > 0:
                coordinates = self.remove_nearby_points(coordinates, min_distance=self.min_distance)

            logger.info(f"Detected points: {len(coordinates)} | Estimated count: {estimated_count}")

            # اگر باز هم نقاط صفر بود و تراکم غیرصفر است، به عنوان fallback نقاط با مقدار بالاتر از میانه را قبول کن
            if len(coordinates) == 0 and estimated_count > 0:
                fallback_mask = processed_map > np.percentile(processed_map, 80)
                y, x = np.where(fallback_mask)
                coordinates = np.column_stack((y, x))
                if len(coordinates) > 0:
                    coordinates = self.remove_nearby_points(coordinates, min_distance=self.min_distance)
                logger.warning(f"Fallback: used percentile mask, got {len(coordinates)} points")

            return coordinates

        except Exception as e:
            logger.error(f"Error in person detection: {str(e)}")
            raise