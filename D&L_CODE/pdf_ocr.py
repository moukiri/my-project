import paddleocr
import cv2
import numpy as np
import fitz  # PyMuPDF
import json
import re
from PIL import Image
import os

class PDFHandwritingOCR:
    def __init__(self):
        # 初始化PaddleOCR，专门用于日语识别
        self.ocr = paddleocr.PaddleOCR(
            use_angle_cls=True, 
            lang='japan',  # 日语识别
            use_gpu=False,
            show_log=False
        )
        
        # 日语月份映射
        self.month_mapping = {
            '1月': '2025-01', '１月': '2025-01', '一月': '2025-01',
            '2月': '2025-02', '２月': '2025-02', '二月': '2025-02',
            '3月': '2025-03', '３月': '2025-03', '三月': '2025-03',
            '4月': '2025-04', '４月': '2025-04', '四月': '2025-04',
            '5月': '2025-05', '５月': '2025-05', '五月': '2025-05',
            '6月': '2025-06', '６月': '2025-06', '六月': '2025-06',
            '7月': '2025-07', '７月': '2025-07', '七月': '2025-07',
            '8月': '2025-08', '８月': '2025-08', '八月': '2025-08',
            '9月': '2025-09', '９月': '2025-09', '九月': '2025-09',
            '10月': '2025-10', '１０月': '2025-10', '十月': '2025-10',
            '11月': '2025-11', '１１月': '2025-11', '十一月': '2025-11',
            '12月': '2025-12', '１２月': '2025-12', '十二月': '2025-12'
        }

    def pdf_to_images(self, pdf_path, dpi=300):
        """
        将PDF转换为高质量图片
        """
        doc = fitz.open(pdf_path)
        images = []
        
        for page_num in range(len(doc)):
            page = doc[page_num]
            # 高DPI确保文字清晰
            mat = fitz.Matrix(dpi/72, dpi/72)
            pix = page.get_pixmap(matrix=mat)
            img_data = pix.tobytes("png")
            
            # 转换为OpenCV格式
            nparr = np.frombuffer(img_data, np.uint8)
            img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
            images.append(img)
            
        doc.close()
        return images

    def detect_red_marks(self, image):
        """
        检测红色标记（圆圈和×）
        """
        # 转换到HSV色彩空间，更容易检测红色
        hsv = cv2.cvtColor(image, cv2.COLOR_BGR2HSV)
        
        # 红色的HSV范围
        lower_red1 = np.array([0, 50, 50])
        upper_red1 = np.array([10, 255, 255])
        lower_red2 = np.array([170, 50, 50])
        upper_red2 = np.array([180, 255, 255])
        
        # 创建红色掩码
        mask1 = cv2.inRange(hsv, lower_red1, upper_red1)
        mask2 = cv2.inRange(hsv, lower_red2, upper_red2)
        red_mask = mask1 + mask2
        
        # 查找红色区域的轮廓
        contours, _ = cv2.findContours(red_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        
        red_regions = []
        for contour in contours:
            area = cv2.contourArea(contour)
            if area > 100:  # 过滤太小的区域
                x, y, w, h = cv2.boundingRect(contour)
                red_regions.append({
                    'bbox': (x, y, w, h),
                    'area': area,
                    'type': self.classify_red_mark(contour)
                })
        
        return red_regions

    def classify_red_mark(self, contour):
        """
        分类红色标记类型（圆圈或×）
        """
        # 简单的形状分析
        area = cv2.contourArea(contour)
        perimeter = cv2.arcLength(contour, True)
        
        if perimeter > 0:
            circularity = 4 * np.pi * area / (perimeter * perimeter)
            if circularity > 0.5:
                return 'circle'  # 圆圈
            else:
                return 'cross'   # ×或其他形状
        return 'unknown'

    def extract_handwriting_regions(self, image, red_regions):
        """
        基于红色标记提取手写区域
        """
        handwriting_regions = []
        
        for red_region in red_regions:
            x, y, w, h = red_region['bbox']
            mark_type = red_region['type']
            
            # 根据标记类型确定手写区域
            if mark_type == 'circle':
                # 圆圈内的手写，扩展搜索区域
                expand = 30
                hw_x = max(0, x - expand)
                hw_y = max(0, y - expand)
                hw_w = min(image.shape[1] - hw_x, w + 2*expand)
                hw_h = min(image.shape[0] - hw_y, h + 2*expand)
                
                handwriting_regions.append({
                    'bbox': (hw_x, hw_y, hw_w, hw_h),
                    'type': 'single_month',  # 单个月份
                    'red_mark': red_region
                })
                
            elif mark_type == 'cross':
                # ×标记，需要找到对应的范围
                # 这里需要更复杂的逻辑来确定影响范围
                handwriting_regions.append({
                    'bbox': (x, y, w, h),
                    'type': 'range_month',  # 范围月份
                    'red_mark': red_region
                })
        
        return handwriting_regions

    def enhance_handwriting_image(self, image_region):
        """
        增强手写图像，提高识别准确率
        """
        # 转为灰度
        if len(image_region.shape) == 3:
            gray = cv2.cvtColor(image_region, cv2.COLOR_BGR2GRAY)
        else:
            gray = image_region.copy()
        
        # 去噪
        denoised = cv2.fastNlMeansDenoising(gray)
        
        # 增强对比度
        clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8,8))
        enhanced = clahe.apply(denoised)
        
        # 锐化
        kernel = np.array([[-1,-1,-1], [-1,9,-1], [-1,-1,-1]])
        sharpened = cv2.filter2D(enhanced, -1, kernel)
        
        # 二值化
        _, binary = cv2.threshold(sharpened, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        
        return binary

    def recognize_month_text(self, image_region):
        """
        识别手写月份文字
        """
        # 图像增强
        enhanced_img = self.enhance_handwriting_image(image_region)
        
        # 使用PaddleOCR识别
        try:
            result = self.ocr.ocr(enhanced_img, cls=True)
            
            if result and result[0]:
                # 提取所有识别到的文字
                texts = []
                for line in result[0]:
                    text = line[1][0]
                    confidence = line[1][1]
                    if confidence > 0.3:  # 置信度阈值
                        texts.append(text)
                
                # 查找月份
                for text in texts:
                    month = self.parse_month(text)
                    if month:
                        return month
                        
        except Exception as e:
            print(f"OCR识别错误: {e}")
        
        return None

    def parse_month(self, text):
        """
        解析月份文字，返回标准格式
        """
        # 清理文字
        text = text.strip()
        
        # 直接匹配
        if text in self.month_mapping:
            return self.month_mapping[text]
        
        # 模糊匹配
        for month_text, standard_format in self.month_mapping.items():
            if month_text in text or text in month_text:
                return standard_format
        
        # 数字匹配
        numbers = re.findall(r'\d+', text)
        if numbers:
            try:
                month_num = int(numbers[0])
                if 1 <= month_num <= 12:
                    return f"2025-{month_num:02d}"
            except:
                pass
        
        return None

    def extract_item_info(self, image, red_regions):
        """
        提取项番和注番信息
        """
        # 这里需要根据表格结构来定位
        # 暂时返回简单的OCR结果
        try:
            result = self.ocr.ocr(image, cls=True)
            
            items = []
            if result and result[0]:
                for line in result[0]:
                    text = line[1][0]
                    bbox = line[0]
                    
                    # 查找JS开头的项番
                    if text.startswith('JS'):
                        items.append({
                            'type': 'item_number',
                            'text': text,
                            'bbox': bbox
                        })
                    # 查找T开头的注番
                    elif text.startswith('T'):
                        items.append({
                            'type': 'note_number', 
                            'text': text,
                            'bbox': bbox
                        })
            
            return items
            
        except Exception as e:
            print(f"提取项目信息错误: {e}")
            return []

    def process_pdf(self, pdf_path, output_file=None):
        """
        处理PDF文件的主函数
        """
        print(f"开始处理PDF: {pdf_path}")
        
        # 1. PDF转图片
        images = self.pdf_to_images(pdf_path)
        print(f"转换了 {len(images)} 页图片")
        
        all_results = []
        
        for page_num, image in enumerate(images):
            print(f"\n处理第 {page_num + 1} 页...")
            
            # 2. 检测红色标记
            red_regions = self.detect_red_marks(image)
            print(f"检测到 {len(red_regions)} 个红色标记")
            
            # 3. 提取手写区域
            handwriting_regions = self.extract_handwriting_regions(image, red_regions)
            
            # 4. 识别手写月份
            page_results = []
            for hw_region in handwriting_regions:
                x, y, w, h = hw_region['bbox']
                region_img = image[y:y+h, x:x+w]
                
                month = self.recognize_month_text(region_img)
                if month:
                    print(f"识别到月份: {month}")
                    page_results.append({
                        'month': month,
                        'type': hw_region['type'],
                        'bbox': hw_region['bbox']
                    })
            
            # 5. 提取项番信息
            items = self.extract_item_info(image, red_regions)
            
            all_results.append({
                'page': page_num + 1,
                'months': page_results,
                'items': items
            })
        
        # 保存结果
        if output_file:
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(all_results, f, ensure_ascii=False, indent=2)
            print(f"\n结果已保存到: {output_file}")
        
        return all_results

def main():
    """
    使用示例
    """
    ocr_processor = PDFHandwritingOCR()
    
    # 替换为你的PDF文件路径
    pdf_path = "test_document.pdf"
    output_file = "recognition_results.json"
    
    if os.path.exists(pdf_path):
        results = ocr_processor.process_pdf(pdf_path, output_file)
        
        # 打印结果摘要
        print("\n=== 识别结果摘要 ===")
        for result in results:
            print(f"第{result['page']}页:")
            for month_info in result['months']:
                print(f"  月份: {month_info['month']}")
                print(f"  类型: {month_info['type']}")
    else:
        print(f"PDF文件不存在: {pdf_path}")
        print("请修改pdf_path为实际文件路径")

if __name__ == "__main__":
    # 安装依赖检查
    try:
        import fitz
        print("✓ PyMuPDF 已安装")
    except ImportError:
        print("需要安装 PyMuPDF: pip install PyMuPDF")
    
    try:
        import paddleocr
        print("✓ PaddleOCR 已安装")
    except ImportError:
        print("需要安装 PaddleOCR")
    
    # 运行主程序
    main()