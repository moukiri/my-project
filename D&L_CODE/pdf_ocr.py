import paddleocr
import cv2
import numpy as np
import fitz  # PyMuPDF
import json
import re
from PIL import Image
import os
import tkinter as tk
from tkinter import filedialog, messagebox

class PDFHandwritingOCR:
    def __init__(self):
        # 初始化PaddleOCR，专门用于日语识别
        self.ocr = paddleocr.PaddleOCR(
            use_textline_orientation=True,  # 替换 use_angle_cls
            lang='japan'  # 日语识别
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
            result = self.ocr.ocr(enhanced_img)
            
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
        try:
            result = self.ocr.ocr(image)
            
            js_items = []  # JS项番
            potential_notes = []  # 潜在的注番
            
            if result and result[0]:
                for line in result[0]:
                    text = line[1][0].strip()
                    bbox = line[0]
                    confidence = line[1][1]
                    
                    # 提取中心点坐标用于位置判断
                    center_y = (bbox[0][1] + bbox[2][1]) / 2
                    center_x = (bbox[0][0] + bbox[2][0]) / 2
                    
                    # 查找JS开头的项番
                    if text.startswith('JS'):
                        js_items.append({
                            'type': 'item_number',
                            'text': text,
                            'bbox': bbox,
                            'center_y': center_y,
                            'center_x': center_x,
                            'confidence': confidence
                        })
                    
                    # 查找潜在的注番 (字母+数字组合，且不是JS开头)
                    elif self.is_potential_note_number(text):
                        potential_notes.append({
                            'type': 'potential_note',
                            'text': text,
                            'bbox': bbox,
                            'center_y': center_y,
                            'center_x': center_x,
                            'confidence': confidence
                        })
            
            # 匹配注番和项番的对应关系
            note_item_pairs = self.match_notes_to_items(potential_notes, js_items)
            
            return {
                'js_items': js_items,
                'note_numbers': note_item_pairs,
                'all_potential_notes': potential_notes
            }
            
        except Exception as e:
            print(f"提取项目信息错误: {e}")
            return {'js_items': [], 'note_numbers': [], 'all_potential_notes': []}

    def is_potential_note_number(self, text):
        """
        判断文本是否可能是注番
        """
        import re
        
        # 清理文本
        text = text.strip().replace(' ', '')
        
        # 注番的可能模式：
        # 1. 字母开头 + 数字：HA05543, JA21671, RA11360, T614600
        # 2. 长度通常在6-8位
        # 3. 不是JS开头
        
        patterns = [
            r'^[A-Z]{1,3}\d{4,6}

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

def select_pdf_file():
    """
    打开文件选择器，让用户选择PDF文件
    """
    # 创建一个隐藏的根窗口
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    
    # 设置文件选择器
    file_path = filedialog.askopenfilename(
        title="选择要识别的PDF文件",
        filetypes=[
            ("PDF文件", "*.pdf"),
            ("所有文件", "*.*")
        ],
        initialdir=os.getcwd()  # 从当前目录开始
    )
    
    root.destroy()  # 销毁临时窗口
    return file_path

def select_output_folder():
    """
    选择结果保存文件夹
    """
    root = tk.Tk()
    root.withdraw()
    
    folder_path = filedialog.askdirectory(
        title="选择结果保存文件夹",
        initialdir=os.getcwd()
    )
    
    root.destroy()
    return folder_path

def main():
    """
    主函数 - 使用文件选择器
    """
    print("=== PaddleOCR PDF手写识别工具 ===")
    print()
    
    # 1. 选择PDF文件
    print("请选择要识别的PDF文件...")
    pdf_path = select_pdf_file()
    
    if not pdf_path:
        print("未选择文件，程序退出。")
        return
    
    if not os.path.exists(pdf_path):
        print(f"文件不存在: {pdf_path}")
        return
    
    print(f"已选择文件: {os.path.basename(pdf_path)}")
    print(f"文件路径: {pdf_path}")
    
    # 2. 选择保存位置
    print("\n请选择结果保存文件夹...")
    output_folder = select_output_folder()
    
    if not output_folder:
        # 如果没选择，使用PDF文件同目录
        output_folder = os.path.dirname(pdf_path)
        print(f"使用默认保存位置: {output_folder}")
    else:
        print(f"结果将保存到: {output_folder}")
    
    # 3. 生成输出文件名
    pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]
    output_file = os.path.join(output_folder, f"{pdf_name}_识别结果.json")
    
    # 4. 开始处理
    try:
        print("\n正在初始化OCR引擎...")
        ocr_processor = PDFHandwritingOCR()
        
        print("开始处理PDF文件...")
        results = ocr_processor.process_pdf(pdf_path, output_file)
        
        # 5. 显示结果摘要
        print("\n" + "="*60)
        print("🎉 识别完成！")
        print("="*60)
        
        total_months = 0
        total_items = 0
        
        for result in results:
            page_months = len(result['months'])
            page_items = len(result['items']['js_items'])
            total_months += page_months
            total_items += page_items
            
            print(f"第{result['page']}页: 识别到 {page_months} 个月份, {page_items} 个项番")
            
            # 显示识别到的月份
            for month_info in result['months']:
                print(f"  📅 月份: {month_info['month']} ({month_info['type']})")
            
            # 显示注番-项番对应关系
            if result['items']['note_numbers']:
                print(f"  📋 注番-项番对应关系:")
                for pair in result['items']['note_numbers'][:3]:  # 只显示前3个
                    print(f"    {pair['note_number']} → {pair['item_number']}")
                if len(result['items']['note_numbers']) > 3:
                    print(f"    ... 还有 {len(result['items']['note_numbers']) - 3} 个")
        
        print(f"\n📊 总计: {total_months} 个月份, {total_items} 个项番")
        print(f"💾 详细结果已保存到: {output_file}")
        
        # 6. 询问是否打开结果文件
        try:
            root = tk.Tk()
            root.withdraw()
            
            open_file = messagebox.askyesno(
                "识别完成", 
                f"识别完成！\n\n总计识别到:\n- {total_months} 个月份\n- {total_items} 个项番\n\n是否打开结果文件查看？"
            )
            
            if open_file:
                os.startfile(output_file)  # Windows
            
            root.destroy()
            
        except Exception as e:
            print(f"无法打开文件: {e}")
        
    except Exception as e:
        print(f"❌ 处理过程中出现错误: {e}")
        print("请检查:")
        print("1. PDF文件是否完整")
        print("2. 是否有足够的磁盘空间")
        print("3. 文件路径中是否包含特殊字符")

def quick_test():
    """
    快速测试函数 - 不使用文件选择器
    """
    print("=== 快速测试模式 ===")
    
    # 在当前目录查找PDF文件
    current_dir = os.getcwd()
    pdf_files = [f for f in os.listdir(current_dir) if f.lower().endswith('.pdf')]
    
    if pdf_files:
        print(f"在当前目录找到以下PDF文件:")
        for i, pdf_file in enumerate(pdf_files, 1):
            print(f"{i}. {pdf_file}")
        
        try:
            choice = input(f"\n请选择文件编号 (1-{len(pdf_files)}) 或按回车使用文件选择器: ").strip()
            
            if choice and choice.isdigit():
                choice_idx = int(choice) - 1
                if 0 <= choice_idx < len(pdf_files):
                    pdf_path = os.path.join(current_dir, pdf_files[choice_idx])
                    output_file = os.path.join(current_dir, f"{os.path.splitext(pdf_files[choice_idx])[0]}_结果.json")
                    
                    ocr_processor = PDFHandwritingOCR()
                    results = ocr_processor.process_pdf(pdf_path, output_file)
                    print(f"\n结果已保存到: {output_file}")
                    return
        except:
            pass
    
    # 如果没有找到文件或用户没选择，使用文件选择器
    main()

if __name__ == "__main__":
    # 安装依赖检查
    try:
        import fitz
        print("✓ PyMuPDF 已安装")
    except ImportError:
        print("❌ 需要安装 PyMuPDF: pip install PyMuPDF")
        exit(1)
    
    try:
        import paddleocr
        # 测试基础功能
        test_ocr = paddleocr.PaddleOCR(lang='ch')  # 删除show_log参数
        print("✓ PaddleOCR 初始化成功")
    except ImportError:
        print("❌ 需要安装 PaddleOCR: pip install paddlepaddle paddleocr")
        exit(1)
    except Exception as e:
        print(f"❌ PaddleOCR 初始化失败: {e}")
        exit(1)
    
    # 检查运行模式
    import sys
    
    if len(sys.argv) > 1 and sys.argv[1] == "--quick":
        # 快速测试模式
        quick_test()
    else:
        # 标准模式 - 使用文件选择器
        main()
,  # 1-3个字母 + 4-6个数字
            r'^[A-Z]\d{6}

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

if __name__ == "__main__":
    # 安装依赖检查
    try:
        import fitz
        print("✓ PyMuPDF 已安装")
    except ImportError:
        print("需要安装 PyMuPDF: pip install PyMuPDF")
    
    try:
        import paddleocr
        # 测试基础功能
        test_ocr = paddleocr.PaddleOCR(lang='ch')  # 删除show_log参数
        print("✓ PaddleOCR 初始化成功")
    except ImportError:
        print("需要安装 PaddleOCR")
    except Exception as e:
        print(f"PaddleOCR 初始化失败: {e}")
        print("可能需要更新 PaddleOCR 版本")
    
    # 运行主程序
    main()
,         # 1个字母 + 6个数字
            r'^[A-Z]{2}\d{5}

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
    pdf_path = "your_document.pdf"
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
        # 测试基础功能
        test_ocr = paddleocr.PaddleOCR(lang='ch')  # 删除show_log参数
        print("✓ PaddleOCR 初始化成功")
    except ImportError:
        print("需要安装 PaddleOCR")
    except Exception as e:
        print(f"PaddleOCR 初始化失败: {e}")
        print("可能需要更新 PaddleOCR 版本")
    
    # 运行主程序
    main()
,      # 2个字母 + 5个数字
        ]
        
        for pattern in patterns:
            if re.match(pattern, text) and not text.startswith('JS'):
                return True
        return False

    def match_notes_to_items(self, potential_notes, js_items):
        """
        根据位置关系匹配注番和项番
        注番通常在对应项番的上方
        """
        matched_pairs = []
        
        for js_item in js_items:
            js_y = js_item['center_y']
            js_x = js_item['center_x']
            
            # 找到在此JS项番上方且X坐标相近的注番
            candidates = []
            
            for note in potential_notes:
                note_y = note['center_y']
                note_x = note['center_x']
                
                # 条件：注番在JS上方，且X坐标差距不太大
                if (note_y < js_y and  # 注番在上方
                    abs(note_x - js_x) < 100 and  # X坐标相近 (可调整阈值)
                    js_y - note_y < 150):  # 距离不要太远 (可调整阈值)
                    
                    candidates.append({
                        'note': note,
                        'distance': js_y - note_y
                    })
            
            # 选择距离最近的作为对应的注番
            if candidates:
                closest = min(candidates, key=lambda x: x['distance'])
                matched_pairs.append({
                    'note_number': closest['note']['text'],
                    'item_number': js_item['text'],
                    'note_bbox': closest['note']['bbox'],
                    'item_bbox': js_item['bbox']
                })
        
        return matched_pairs

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
    pdf_path = "your_document.pdf"
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
        # 测试基础功能
        test_ocr = paddleocr.PaddleOCR(lang='ch')  # 删除show_log参数
        print("✓ PaddleOCR 初始化成功")
    except ImportError:
        print("需要安装 PaddleOCR")
    except Exception as e:
        print(f"PaddleOCR 初始化失败: {e}")
        print("可能需要更新 PaddleOCR 版本")
    
    # 运行主程序
    main()