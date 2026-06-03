import os
import unittest
from PIL import Image
from processor import PhotoProcessor

class TestPhotoProcessor(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.project_dir = os.path.dirname(os.path.abspath(__file__))
        cls.test_data_dir = os.path.join(cls.project_dir, "test_data")
        cls.excel_path = os.path.join(cls.test_data_dir, "students.xlsx")
        cls.output_dir = os.path.join(cls.test_data_dir, "test_output")
        os.makedirs(cls.output_dir, exist_ok=True)
        
    def test_01_get_excel_headers(self):
        """測試讀取 Excel 標頭"""
        headers = PhotoProcessor.get_excel_headers(self.excel_path)
        self.assertIn("身分證字號", headers)
        self.assertIn("學生姓名", headers)
        self.assertIn("學號", headers)
        
    def test_02_load_excel_mapping(self):
        """測試讀取 Excel 身分證至學號對照"""
        mapping = PhotoProcessor.load_excel_mapping(self.excel_path, "身分證字號", "學號")
        self.assertEqual(mapping.get("A123456789"), "112001")
        self.assertEqual(mapping.get("B987654321"), "112002")
        self.assertEqual(mapping.get("C111111111"), "112003")
        self.assertIsNone(mapping.get("D999999999")) # 此學號不在 Excel 中
        
    def test_03_scan_source_directory(self):
        """測試掃描圖片資料夾"""
        photos = PhotoProcessor.scan_source_directory(self.test_data_dir)
        # 應掃描到 4 張測試照片
        filenames = [p['filename'] for p in photos]
        self.assertTrue(any("A123456789.jpg" in f for f in filenames))
        self.assertTrue(any("B987654321.png" in f for f in filenames))
        self.assertTrue(any("C111111111.jpeg" in f for f in filenames))
        self.assertTrue(any("D999999999.jpg" in f for f in filenames))
        
    def test_04_calculate_crop_box_center(self):
        """測試裁切框計算 (置中裁切)"""
        # 原圖 600x800，目標 300x400 (比例 3:4)
        # 原圖與目標比例一致，故最大裁切寬高應為 600x800
        x1, y1, x2, y2 = PhotoProcessor.calculate_crop_box(600, 800, 300, 400, scale=1.0, vertical_offset=0.0)
        self.assertEqual((x1, y1, x2, y2), (0, 0, 600, 800))
        
        # 原圖 800x600，目標 300x400 (比例 3:4)
        # 目標為窄高，原圖為寬扁，應以高度為基準
        # 最大裁切高度 = 600，對應寬度 = 600 * (300/400) = 450
        # 置中裁切 x 軸起點應為 (800 - 450) / 2 = 175
        x1, y1, x2, y2 = PhotoProcessor.calculate_crop_box(800, 600, 300, 400, scale=1.0, vertical_offset=0.0)
        self.assertEqual((x1, y1, x2, y2), (175, 0, 625, 600))

    def test_05_calculate_crop_box_offset(self):
        """測試裁切框垂直偏移與邊界限制 (Clamp)"""
        # 原圖 600x800，目標 300x400，縮放 0.8，垂直偏移 0.15 (往上移動裁切框，即 y1 變小)
        # 最大裁切 w, h 為 600x800
        # 縮放後為 480x640
        # 預設中心 y = 400。偏移 0.15 代表 center_y 向上移動 0.15 * 800 = 120 像素，新中心 y = 280
        # 新裁切框高 640，y1 = 280 - 320 = -40 -> 超出邊界
        # Clamp 後，y1 應為 0，y2 應為 640
        x1, y1, x2, y2 = PhotoProcessor.calculate_crop_box(600, 800, 300, 400, scale=0.8, vertical_offset=0.15)
        self.assertEqual(y1, 0)
        self.assertEqual(y2, 640)
        self.assertEqual(x2 - x1, 480)

    def test_06_process_image_preview(self):
        """測試相片處理預覽 (不儲存檔案)"""
        src_path = os.path.join(self.test_data_dir, "A123456789.jpg")
        preview_img = PhotoProcessor.process_image(
            src_path, dest_path=None, target_width=300, target_height=400, scale=1.0, vertical_offset=0.0
        )
        self.assertIsInstance(preview_img, Image.Image)
        self.assertEqual(preview_img.size, (300, 400))

    def test_07_process_image_save(self):
        """測試相片處理並存檔 (轉檔與裁切)"""
        src_path = os.path.join(self.test_data_dir, "B987654321.png") # 原圖為 PNG
        dest_path = os.path.join(self.output_dir, "112002.jpg")       # 輸出為 JPG
        
        result = PhotoProcessor.process_image(
            src_path, dest_path=dest_path, target_width=354, target_height=472, scale=0.9, vertical_offset=0.1, output_format="JPEG"
        )
        self.assertTrue(result)
        self.assertTrue(os.path.exists(dest_path))
        
        # 驗證輸出圖片規格
        with Image.open(dest_path) as out_img:
            self.assertEqual(out_img.size, (354, 472))
            self.assertEqual(out_img.format, "JPEG")

if __name__ == "__main__":
    unittest.main()
