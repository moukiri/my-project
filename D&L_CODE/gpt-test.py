import pytesseract
import cv2
from pdf2image import convert_from_path
import re
import numpy as np
from tkinter.filedialog import askopenfilename
import matplotlib.pyplot as plt
from PIL import Image, ImageEnhance

# Tesseract 路径
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def extract_images_from_pdf(pdf_path):
    """从PDF提取高质量图像"""
    return convert_from_path(pdf_path, dpi=600, fmt='PNG')  # 提高到600 DPI

def detect_red_annotations(image):
    """检测红色标注区域"""
    img_array = np.array(image)
    
    # 转换到HSV色彩空间，更容易检测红色
    hsv = cv2.cvtColor(img_array, cv2.COLOR_RGB2HSV)
    
    # 定义红色范围（两个范围，因为红色在HSV中跨越0度）
    lower_red1 = np.array([0, 50, 50])
    upper_red1 = np.array([10, 255, 255])
    lower_red2 = np.array([170, 50, 50])
    upper_red2 = np.array([180, 255, 255])
    
    # 创建红色掩码
    mask1 = cv2.inRange(hsv, lower_red1, upper_red1)
    mask2 = cv2.inRange(hsv, lower_red2, upper_red2)
    red_mask = mask1 + mask2
    
    # 形态学操作，连接近邻的红色区域
    kernel = np.ones((5, 5), np.uint8)
    red_mask = cv2.morphologyEx(red_mask, cv2.MORPH_CLOSE, kernel)
    red_mask = cv2.morphologyEx(red_mask, cv2.MORPH_OPEN, kernel)
    
    return red_mask

def find_red_circles_and_text(image):
    """查找红色圆圈及其内部的手写文字"""
    red_mask = detect_red_annotations(image)
    
    # 查找轮廓
    contours, _ = cv2.findContours(red_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    circles = []
    img_array = np.array(image)
    
    for contour in contours:
        # 计算轮廓面积和周长
        area = cv2.contourArea(contour)
        perimeter = cv2.arcLength(contour, True)
        
        if area < 100:  # 过滤太小的区域
            continue
            
        # 获取边界框
        x, y, w, h = cv2.boundingRect(contour)
        
        # 圆形度检测（可选，因为手写圆圈可能不规则）
        if perimeter > 0:
            circularity = 4 * np.pi * area / (perimeter * perimeter)
        else:
            circularity = 0
            
        # 扩展区域以包含可能的文字
        margin = 10
        x_start = max(0, x - margin)
        y_start = max(0, y - margin)
        x_end = min(img_array.shape[1], x + w + margin)
        y_end = min(img_array.shape[0], y + h + margin)
        
        # 提取该区域
        roi = img_array[y_start:y_end, x_start:x_end]
        
        circles.append({
            'bbox': (x, y, w, h),
            'extended_bbox': (x_start, y_start, x_end, y_end),
            'roi': roi,
            'area': area,
            'circularity': circularity,
            'center': (x + w//2, y + h//2)
        })
    
    return circles

def detect_cross_marks(image):
    """检测红色的×标记"""
    red_mask = detect_red_annotations(image)
    
    # 查找轮廓
    contours, _ = cv2.findContours(red_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    crosses = []
    img_array = np.array(image)
    
    for contour in contours:
        area = cv2.contourArea(contour)
        
        if area < 50 or area > 2000:  # 过滤不合适的大小
            continue
            
        # 获取边界框
        x, y, w, h = cv2.boundingRect(contour)
        
        # ×的特征：宽高比接近1，不太圆
        aspect_ratio = w / h if h > 0 else 0
        
        if 0.5 < aspect_ratio < 2.0:  # 宽高比合理
            # 扩展区域
            margin = 20
            x_start = max(0, x - margin)
            y_start = max(0, y - margin)
            x_end = min(img_array.shape[1], x + w + margin)
            y_end = min(img_array.shape[0], y + h + margin)
            
            roi = img_array[y_start:y_end, x_start:x_end]
            
            crosses.append({
                'bbox': (x, y, w, h),
                'extended_bbox': (x_start, y_start, x_end, y_end),
                'roi': roi,
                'center': (x + w//2, y + h//2)
            })
    
    return crosses

def preprocess_for_handwriting(roi):
    """专门针对手写文字的预处理"""
    if len(roi.shape) == 3:
        gray = cv2.cvtColor(roi, cv2.COLOR_RGB2GRAY)
    else:
        gray = roi
    
    # 1. 增强对比度
    pil_img = Image.fromarray(gray)
    enhancer = ImageEnhance.Contrast(pil_img)
    enhanced = enhancer.enhance(2.0)  # 增强对比度
    enhanced_array = np.array(enhanced)
    
    # 2. 高斯模糊
    blurred = cv2.GaussianBlur(enhanced_array, (3, 3), 0)
    
    # 3. 自适应阈值
    thresh = cv2.adaptiveThreshold(blurred, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
                                  cv2.THRESH_BINARY, 11, 2)
    
    # 4. 形态学操作，连接断开的笔画
    kernel = np.ones((2, 2), np.uint8)
    processed = cv2.morphologyEx(thresh, cv2.MORPH_CLOSE, kernel)
    
    # 5. 膨胀操作，让笔画更粗一些
    dilated = cv2.dilate(processed, kernel, iterations=1)
    
    return dilated

def ocr_handwritten_month(roi):
    """专门识别手写月份的OCR"""
    processed = preprocess_for_handwriting(roi)
    
    # 多种OCR配置专门针对单个汉字或数字+月
    configs = [
        '--oem 3 --psm 8 -c tessedit_char_whitelist=0123456789月',  # 只允许数字和月
        '--oem 3 --psm 7 -c tessedit_char_whitelist=0123456789月',
        '--oem 3 --psm 13 -c tessedit_char_whitelist=0123456789月',
        '--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789月',
        '--oem 3 --psm 8',  # 单个字符模式
        '--oem 3 --psm 10', # 单个字符模式
    ]
    
    results = []
    
    for i, config in enumerate(configs):
        try:
            # 先尝试日文识别
            text_jp = pytesseract.image_to_string(processed, lang='jpn', config=config)
            text_jp = text_jp.strip().replace(' ', '').replace('\n', '')
            
            # 再尝试中文识别（因为月字在中日文中相同）
            text_ch = pytesseract.image_to_string(processed, lang='chi_sim', config=config)
            text_ch = text_ch.strip().replace(' ', '').replace('\n', '')
            
            if text_jp:
                results.append(text_jp)
            if text_ch and text_ch != text_jp:
                results.append(text_ch)
                
            print(f"手写识别Config{i}: 日文='{text_jp}' 中文='{text_ch}'")
            
        except Exception as e:
            print(f"手写OCR Config {i} 失败: {e}")
    
    return results

def extract_month_from_text(text_results):
    """从识别结果中提取月份"""
    month_patterns = [
        r'([0-9]{1,2})月',      # 数字+月
        r'([0-9]{1,2})',        # 单纯数字
        r'([一二三四五六七八九十]+)月',  # 中文数字+月
    ]
    
    # 中文数字转换
    chinese_to_num = {
        '一': 1, '二': 2, '三': 3, '四': 4, '五': 5,
        '六': 6, '七': 7, '八': 8, '九': 9, '十': 10,
        '十一': 11, '十二': 12
    }
    
    for text in text_results:
        if not text:
            continue
            
        print(f"分析文本: '{text}'")
        
        for pattern in month_patterns:
            matches = re.findall(pattern, text)
            for match in matches:
                try:
                    if match in chinese_to_num:
                        month_num = chinese_to_num[match]
                    else:
                        month_num = int(match)
                    
                    if 1 <= month_num <= 12:
                        print(f"识别到月份: {month_num}")
                        return f"2025-{month_num:02d}"
                except ValueError:
                    continue
    
    return None

def find_js_numbers_in_left_column(image, annotation_y):
    """在注释的Y坐标附近查找左侧的JS编号"""
    img_array = np.array(image)
    height, width = img_array.shape[:2]
    
    # 定义左侧区域（左边1/4）
    left_region = img_array[:, :width//4]
    
    # 转换为灰度并预处理
    if len(left_region.shape) == 3:
        gray = cv2.cvtColor(left_region, cv2.COLOR_RGB2GRAY)
    else:
        gray = left_region
    
    # 简单二值化
    _, thresh = cv2.threshold(gray, 127, 255, cv2.THRESH_BINARY)
    
    # OCR识别左侧区域
    config = '--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZJS'
    text = pytesseract.image_to_string(thresh, lang='eng', config=config)
    
    print(f"左侧区域识别文本:\n{text}")
    
    # 查找JS编号和T编号
    js_pattern = r'JS\s?(\d{4,}[A-Z]*)'
    t_pattern = r'T(\d{6,})'
    
    js_matches = re.findall(js_pattern, text, re.IGNORECASE)
    t_matches = re.findall(t_pattern, text, re.IGNORECASE)
    
    # 使用OCR获取文本位置信息来匹配Y坐标
    try:
        data = pytesseract.image_to_data(thresh, lang='eng', config=config, output_type=pytesseract.Output.DICT)
        
        js_numbers = []
        t_numbers = []
        
        for i, word in enumerate(data['text']):
            if not word.strip():
                continue
                
            word_y = data['top'][i] + data['height'][i] // 2
            
            # 检查是否在注释附近（±50像素范围内）
            if abs(word_y - annotation_y) <= 50:
                if re.match(r'JS\d+', word, re.IGNORECASE):
                    js_numbers.append(word.upper())
                elif re.match(r'T\d+', word, re.IGNORECASE):
                    t_numbers.append(word.upper())
        
        return js_numbers, t_numbers
        
    except Exception as e:
        print(f"位置匹配失败: {e}")
        # 返回所有找到的编号
        return [f"JS{js}" for js in js_matches], [f"T{t}" for t in t_matches]

def process_page(image, page_num):
    """处理单页图像"""
    print(f"\n{'='*60}")
    print(f"处理第 {page_num} 页")
    print(f"{'='*60}")
    
    results = []
    
    # 1. 查找红色圆圈（类型1：单行标注）
    circles = find_red_circles_and_text(image)
    print(f"发现 {len(circles)} 个红色圆圈区域")
    
    for i, circle in enumerate(circles):
        print(f"\n--- 处理圆圈 {i+1} ---")
        
        # OCR识别圆圈内的月份
        month_texts = ocr_handwritten_month(circle['roi'])
        month = extract_month_from_text(month_texts)
        
        if month:
            # 查找对应的JS编号
            annotation_y = circle['center'][1]
            js_numbers, t_numbers = find_js_numbers_in_left_column(image, annotation_y)
            
            for js_num in js_numbers:
                results.append({
                    'type': 'single_line',
                    'page': page_num,
                    '注番': t_numbers[0] if t_numbers else '',
                    '项番': js_num,
                    '日期': month,
                    'position': circle['center']
                })
                print(f"类型1识别结果: 注番={t_numbers[0] if t_numbers else '未知'}, 项番={js_num}, 日期={month}")
    
    # 2. 查找红色×标记（类型2：区段标注）
    crosses = detect_cross_marks(image)
    print(f"\n发现 {len(crosses)} 个红色×标记")
    
    for i, cross in enumerate(crosses):
        print(f"\n--- 处理×标记 {i+1} ---")
        
        # 在×标记右侧查找圆圈中的月份
        cross_x, cross_y = cross['center']
        
        # 查找×标记右侧的圆圈
        for circle in circles:
            circle_x, circle_y = circle['center']
            
            # 如果圆圈在×标记右侧且Y坐标接近
            if circle_x > cross_x and abs(circle_y - cross_y) <= 30:
                month_texts = ocr_handwritten_month(circle['roi'])
                month = extract_month_from_text(month_texts)
                
                if month:
                    # 查找×标记左侧对应的所有JS编号
                    js_numbers, t_numbers = find_js_numbers_in_left_column(image, cross_y)
                    
                    # 对于类型2，可能需要查找一个范围内的所有JS编号
                    for js_num in js_numbers:
                        results.append({
                            'type': 'range',
                            'page': page_num,
                            '注番': t_numbers[0] if t_numbers else '',
                            '项番': js_num,
                            '日期': month,
                            'position': cross['center']
                        })
                        print(f"类型2识别结果: 注番={t_numbers[0] if t_numbers else '未知'}, 项番={js_num}, 日期={month}")
                break
    
    return results

def main():
    # 选择PDF文件
    pdf_path = askopenfilename(title="选择PDF文件", filetypes=[("PDF files", "*.pdf")])
    if not pdf_path:
        print("未选择文件")
        return
    
    print("正在提取PDF图像...")
    images = extract_images_from_pdf(pdf_path)
    
    all_results = []
    
    for i, img in enumerate(images):
        page_results = process_page(img, i + 1)
        all_results.extend(page_results)
    
    # 输出最终结果
    print(f"\n{'='*80}")
    print("最终识别结果汇总")
    print(f"{'='*80}")
    
    if not all_results:
        print("未识别到任何有效信息")
    else:
        for result in all_results:
            print(f"页面: {result['page']}")
            print(f"类型: {'单行标注' if result['type'] == 'single_line' else '范围标注'}")
            print(f"注番: {result['注番']}")
            print(f"项番: {result['项番']}")
            print(f"日期: {result['日期']}")
            print(f"位置: {result['position']}")
            print("-" * 50)

if __name__ == "__main__":
    main()