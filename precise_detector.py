import numpy as np
import logging
import cv2
from scipy.ndimage import gaussian_filter, maximum_filter
from skimage.feature import peak_local_max
from scipy import ndimage

logger = logging.getLogger(__name__)


class PrecisePersonDetector:
    def __init__(self):
        # پارامترهای بهینه‌سازی شده برای تشخیص دقیق‌تر
        self.min_distance = 2  # کاهش فاصله برای تشخیص افراد نزدیک به هم
        self.gaussian_sigma = 0.8  # کاهش بیشتر سیگما برای حفظ جزئیات
        self.density_multiplier = 40  # تنظیم ضریب برای تخمین واقعی‌تر
        self.min_peak_threshold = 0.005  # کاهش بیشتر آستانه برای تشخیص بهتر
        self.adaptive_threshold = True  # فعال‌سازی آستانه تطبیقی
        logger.info("PrecisePersonDetector initialized with highly optimized params")

    def adaptive_thresholding(self, image):
        """اعمال آستانه‌گذاری تطبیقی پیشرفته"""
        block_size = 11
        C = 2
        return cv2.adaptiveThreshold(
            (image * 255).astype(np.uint8),
            255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY,
            block_size,
            C
        )

    def normalize_density(self, density_map):
        """نرمال‌سازی پیشرفته با تقویت مناطق کم‌تراکم"""
        eps = 1e-8
        normalized = (density_map - density_map.min()) / (density_map.max() - density_map.min() + eps)
        # تقویت مناطق کم‌تراکم با تبدیل گاما
        gamma = 0.7
        normalized = np.power(normalized, gamma)
        return np.clip(normalized, 0, 1)

    def enhance_density_map(self, density_map):
        """بهبود پیشرفته نقشه تراکم با تکنیک‌های متعدد"""
        try:
            # نرمال‌سازی اولیه
            density_norm = self.normalize_density(density_map)

            # تقویت لبه‌ها
            edges = cv2.Canny((density_norm * 255).astype(np.uint8), 50, 150)
            edges = cv2.dilate(edges, None)

            # تقویت کنتراست با CLAHE چند مرحله‌ای
            density_uint8 = (density_norm * 255).astype(np.uint8)
            clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(4, 4))
            enhanced = clahe.apply(density_uint8)
            enhanced = clahe.apply(enhanced)

            # اعمال فیلتر دوطرفه با پارامترهای بهینه
            bilateral = cv2.bilateralFilter(enhanced, 5, 50, 50)

            # ترکیب تصویر اصلی با لبه‌های تقویت شده
            enhanced_float = bilateral.astype(np.float32) / 255.0
            edges_float = edges.astype(np.float32) / 255.0
            combined = cv2.addWeighted(enhanced_float, 0.7, edges_float, 0.3, 0)

            # اعمال فیلتر شارپ‌کننده
            kernel = np.array([[-1, -1, -1], [-1, 9, -1], [-1, -1, -1]]) / 9
            sharpened = cv2.filter2D(combined, -1, kernel)

            # حذف نویز با فیلتر مدین
            denoised = cv2.medianBlur((sharpened * 255).astype(np.uint8), 3)

            # نرمال‌سازی نهایی
            final = denoised.astype(np.float32) / 255.0
            return np.clip(final, 0, 1)

        except Exception as e:
            logger.error(f"Error in density map enhancement: {str(e)}")
            return density_map

    def remove_nearby_points(self, points, min_distance=2):
        """حذف نقاط تکراری با الگوریتم بهینه‌شده"""
        if len(points) == 0:
            return np.empty((0, 2), dtype=int)

        def distance_matrix(points):
            return np.sqrt(((points[:, np.newaxis] - points) ** 2).sum(axis=2))

        # محاسبه ماتریس فاصله
        distances = distance_matrix(points)
        np.fill_diagonal(distances, np.inf)

        # حذف نقاط نزدیک به هم
        mask = np.ones(len(points), dtype=bool)
        for i in range(len(points)):
            if mask[i]:
                mask[distances[i] < min_distance] = False
                mask[i] = True

        return points[mask]

    def detect_local_maxima(self, density_map):
        try:
            # پیش‌پردازش و بهبود نقشه تراکم
            processed_map = self.enhance_density_map(density_map)

            # تخمین تعداد افراد با روش بهبود یافته
            total_density = np.sum(density_map)
            estimated_count = max(1, int(total_density * self.density_multiplier))
            logger.info(f"Total density: {total_density:.3f} | Initial estimate: {estimated_count}")

            # آستانه‌های چندگانه برای تشخیص بهتر
            thresholds = [
                np.percentile(processed_map[processed_map > 0], p)
                for p in [40, 50, 60, 70, 80]
            ]

            coordinates = np.empty((0, 2), dtype=int)

            # تشخیص چند مرحله‌ای با پارامترهای مختلف
            for thresh in thresholds:
                points = peak_local_max(
                    processed_map,
                    min_distance=self.min_distance,
                    threshold_abs=thresh,
                    exclude_border=False,
                    num_peaks=int(estimated_count * 1.5)  # اجازه تشخیص نقاط بیشتر
                )
                if len(points) > 0:
                    coordinates = np.vstack([coordinates, points])

            # حذف نقاط تکراری با الگوریتم بهبود یافته
            if len(coordinates) > 0:
                coordinates = self.remove_nearby_points(coordinates, self.min_distance)

            # روش جایگزین برای مناطق کم‌تراکم
            if len(coordinates) < estimated_count * 0.4:
                logger.info("Using enhanced alternative detection method...")

                # استفاده از آستانه‌گذاری تطبیقی
                if self.adaptive_threshold:
                    binary = self.adaptive_thresholding(processed_map)
                    # یافتن مراکز اجزای متصل
                    labeled, num_features = ndimage.label(binary)
                    centers = ndimage.center_of_mass(binary, labeled, range(1, num_features + 1))
                    if centers:
                        alt_coords = np.array(centers).astype(int)
                        if len(alt_coords) > 0:
                            coordinates = np.vstack([coordinates, alt_coords]) if len(coordinates) > 0 else alt_coords

            # فیلتر نهایی و حذف نقاط تکراری
            if len(coordinates) > 0:
                coordinates = self.remove_nearby_points(coordinates, self.min_distance)

            logger.info(f"Final detection: {len(coordinates)} points | Target estimate: {estimated_count}")
            return coordinates

        except Exception as e:
            logger.error(f"Error in person detection: {str(e)}")
            raise