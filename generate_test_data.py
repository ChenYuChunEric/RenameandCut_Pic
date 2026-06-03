import os
from PIL import Image, ImageDraw, ImageFont
import openpyxl

def generate_test_data():
    project_dir = os.path.dirname(os.path.abspath(__file__))
    test_dir = os.path.join(project_dir, "test_data")
    os.makedirs(test_dir, exist_ok=True)
    
    # 1. 建立 Excel 對照表
    excel_path = os.path.join(test_dir, "students.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "新生名冊"
    
    # 寫入標頭
    headers = ["身分證字號", "學生姓名", "學號"]
    ws.append(headers)
    
    # 寫入測試資料
    students = [
        ("A123456789", "張小明", "112001"),
        ("B987654321", "李小華", "112002"),
        ("C111111111", "王大同", "112003"),
    ]
    for s in students:
        ws.append(s)
        
    wb.save(excel_path)
    print(f"已產生測試 Excel: {excel_path}")
    
    # 2. 建立測試照片
    # 每張圖片故意使用不同的尺寸與格式，以便測試相片處理與轉檔
    photo_configs = [
        ("A123456789", "張小明 (JPG)", (600, 800), "JPEG", ".jpg", "#F39C12"), # 窄長，黃橘色背景
        ("B987654321", "李小華 (PNG)", (800, 600), "PNG", ".png", "#3498DB"),  # 寬扁，藍色背景
        ("C111111111", "王大同 (JPEG)", (500, 500), "JPEG", ".jpeg", "#2ECC71"), # 正方形，綠色背景
        ("D999999999", "無對照 (JPG)", (400, 600), "JPEG", ".jpg", "#9B59B6"),  # 不在 Excel 中，紫色背景
    ]
    
    for filename_id, label, size, fmt, ext, bg_color in photo_configs:
        img = Image.new("RGB", size, bg_color)
        draw = ImageDraw.Draw(img)
        
        w, h = size
        # 繪製簡單的人臉示意圖 (圓形 + 矩形身體)
        # 頭部
        draw.ellipse([w*0.3, h*0.2, w*0.7, h*0.5], fill="#FFE0B2", outline="#E0F2F1", width=3)
        # 眼睛
        draw.ellipse([w*0.4, h*0.3, w*0.45, h*0.35], fill="#000000")
        draw.ellipse([w*0.55, h*0.3, w*0.6, h*0.35], fill="#000000")
        # 嘴巴
        draw.arc([w*0.45, h*0.38, w*0.55, h*0.44], start=0, end=180, fill="#FF5252", width=3)
        # 身體 (肩膀)
        draw.rectangle([w*0.2, h*0.55, w*0.8, h*0.9], fill="#37474F", outline="#CFD8DC", width=2)
        
        # 標記文字
        # 為了避免找不到字型檔，我們不載入 ttf，改用簡單劃線標記檔名或預設字型
        try:
            draw.text((10, 10), label, fill="#FFFFFF")
            draw.text((10, 30), f"Size: {w}x{h}", fill="#FFFFFF")
        except Exception:
            pass
            
        file_path = os.path.join(test_dir, f"{filename_id}{ext}")
        img.save(file_path, format=fmt)
        print(f"已產生測試相片: {file_path}")
        
    print(f"\n所有測試資料已成功產生於: {test_dir}")

if __name__ == "__main__":
    generate_test_data()
