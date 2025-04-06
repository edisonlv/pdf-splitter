import os
import re
import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Canvas, Scrollbar, Frame
import fitz  # PyMuPDF
from PIL import Image, ImageTk
import io

class PDFSplitterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF 分割工具")
        self.root.geometry("1200x800")
        
        # 将主窗口居中显示
        self.center_window(self.root, 1200, 800)
        
        self.pdf_document = None
        self.current_page = 0
        self.scale_factor = 1.0
        self.pdf_path = None
        self.output_dir = None
        
        # 新增：文件列表相关变量
        self.pdf_files = []  # 存储所有待处理的PDF文件路径
        self.current_file_index = -1  # 当前显示的文件索引
        
        # 新增：图幅尺寸相关变量
        self.page_sizes = {}  # 存储每个文件的页面尺寸信息 {file_path: {size_key: [page_numbers]}}
        self.page_templates = {}  # 存储每种尺寸的页码提取区域设置 {size_key: {current_coords, total_coords}}
        self.filename_templates = {}  # 存储每种尺寸的文件名提取区域设置 {size_key: [coords]}
        self.current_size_key = None  # 当前正在设置的尺寸key
        
        # 选择区域变量
        self.start_x = None
        self.start_y = None
        self.rect_id = None
        self.selected_regions = {}  # 修改为字典，key为文件路径，value为区域列表
        self.current_region_text = ""
        
        # 新增：两步选择模式的变量
        self.selection_step = 0  # 0: 未开始, 1: 选择当前页, 2: 选择总页数
        self.current_page_coords = None  # 存储"第n张"的坐标
        self.total_pages_coords = None   # 存储"共m张"的坐标
        self.template_coords_set = False  # 标记是否已设置模板坐标
        self.template_mode = None  # 记录当前模板模式
        
        # 添加文件名模板相关变量
        self.filename_template_coords = []  # 变更为列表，存储多个区域
        self.filename_template_mode = False
        self.custom_filenames = {}  # 存储自定义文件名 {file_path: {page_num: filename}}
        self.template_region_count = 0  # 记录已选择的区域数量
        
        self.create_widgets()
    
    def create_widgets(self):
        # 顶部控制区域
        control_frame = ttk.Frame(self.root)
        control_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Button(control_frame, text="添加PDF文件", command=self.add_pdf_files).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="设置输出目录", command=self.set_output_dir).pack(side=tk.LEFT, padx=5)
        self.split_button = ttk.Button(control_frame, text="分割PDF", command=self.split_pdf, state=tk.DISABLED)
        self.split_button.pack(side=tk.LEFT, padx=5)
        
        # 添加识别图幅按钮
        self.detect_size_button = ttk.Button(control_frame, text="识别图幅尺寸", 
                                           command=self.detect_page_sizes)
        self.detect_size_button.pack(side=tk.LEFT, padx=5)
        
        # 添加文件名模板按钮
        self.filename_button = ttk.Button(control_frame, text="设置文件名提取区域", 
                                        command=self.start_filename_template_selection)
        self.filename_button.pack(side=tk.LEFT, padx=5)
        
        # 修改页码模板设置按钮
        self.template_button1 = ttk.Button(control_frame, text="设置页码提取区域(双区域)", 
                                         command=lambda: self.start_template_selection(mode="double"))
        self.template_button1.pack(side=tk.LEFT, padx=5)
        
        self.template_button2 = ttk.Button(control_frame, text="设置页码提取区域(单区域)", 
                                         command=lambda: self.start_template_selection(mode="single"))
        self.template_button2.pack(side=tk.LEFT, padx=5)
        
        self.template_status = ttk.Label(control_frame, text="未设置页码提取区域")
        self.template_status.pack(side=tk.LEFT, padx=5)
        
        # 主内容区域
        main_frame = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 左侧区域：文件列表和PDF预览
        left_frame = ttk.Frame(main_frame)
        main_frame.add(left_frame, weight=3)
        
        # 文件列表区域
        file_frame = ttk.LabelFrame(left_frame, text="PDF文件列表")
        file_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.file_listbox = tk.Listbox(file_frame, height=6)
        self.file_listbox.pack(fill=tk.X, padx=5, pady=5)
        self.file_listbox.bind('<<ListboxSelect>>', self.on_file_select)
        
        # 添加文件列表操作按钮
        file_btn_frame = ttk.Frame(file_frame)
        file_btn_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(file_btn_frame, text="移除选中", command=self.remove_selected_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(file_btn_frame, text="清空列表", command=self.clear_file_list).pack(side=tk.LEFT, padx=5)
        
        # PDF预览区域
        pdf_frame = ttk.Frame(left_frame)
        pdf_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # PDF导航控制
        nav_frame = ttk.Frame(pdf_frame)
        nav_frame.pack(fill=tk.X)
        
        ttk.Button(nav_frame, text="上一页", command=self.prev_page).pack(side=tk.LEFT, padx=5)
        self.page_label = ttk.Label(nav_frame, text="页面: 0 / 0")
        self.page_label.pack(side=tk.LEFT, padx=5)
        ttk.Button(nav_frame, text="下一页", command=self.next_page).pack(side=tk.LEFT, padx=5)
        
        # PDF显示区域
        self.canvas_frame = ttk.Frame(pdf_frame)
        self.canvas_frame.pack(fill=tk.BOTH, expand=True)
        
        self.canvas = Canvas(self.canvas_frame, bg="gray")
        self.h_scrollbar = ttk.Scrollbar(self.canvas_frame, orient=tk.HORIZONTAL, command=self.canvas.xview)
        self.v_scrollbar = ttk.Scrollbar(self.canvas_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        
        self.canvas.config(xscrollcommand=self.h_scrollbar.set, yscrollcommand=self.v_scrollbar.set)
        
        self.h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        self.v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 绑定鼠标事件
        self.canvas.bind("<ButtonPress-1>", self.on_mouse_down)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_mouse_up)
        
        # 右侧区域：选定区域列表
        right_frame = ttk.Frame(main_frame)
        main_frame.add(right_frame, weight=1)
        
        # 区域列表
        region_frame = ttk.LabelFrame(right_frame, text="选定区域")
        region_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.region_listbox = tk.Listbox(region_frame)
        self.region_listbox.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.region_listbox.bind("<<ListboxSelect>>", self.on_region_select)
        self.region_listbox.bind("<Double-Button-1>", self.on_region_double_click)
        
        region_btn_frame = ttk.Frame(region_frame)
        region_btn_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(region_btn_frame, text="删除区域", command=self.remove_region).pack(side=tk.LEFT, padx=5)
        ttk.Button(region_btn_frame, text="清除所有", command=self.clear_regions).pack(side=tk.LEFT, padx=5)
        ttk.Button(region_btn_frame, text="修改内容", command=self.edit_selected_content).pack(side=tk.LEFT, padx=5)
    
    def add_pdf_files(self):
        """添加PDF文件到列表"""
        files = filedialog.askopenfilenames(
            title="选择PDF文件",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        
        if files:
            for file_path in files:
                if file_path not in self.pdf_files:
                    self.pdf_files.append(file_path)
                    self.selected_regions[file_path] = []  # 为新文件初始化区域列表
            
            self.update_file_list()
            
            # 随机选择一个文件进行预览
            if len(self.pdf_files) > 0:
                import random
                random_index = random.randint(0, len(self.pdf_files) - 1)
                self.file_listbox.selection_clear(0, tk.END)
                self.file_listbox.selection_set(random_index)
                self.load_pdf_file(random_index)
            
            # 启用分割按钮
            self.split_button.config(state=tk.NORMAL)

    def update_file_list(self):
        """更新文件列表显示"""
        self.file_listbox.delete(0, tk.END)
        for file_path in self.pdf_files:
            self.file_listbox.insert(tk.END, os.path.basename(file_path))

    def remove_selected_file(self):
        """移除选中的文件"""
        selection = self.file_listbox.curselection()
        if not selection:
            return
            
        index = selection[0]
        file_path = self.pdf_files[index]
        
        # 移除文件和相关数据
        self.pdf_files.pop(index)
        if file_path in self.selected_regions:
            del self.selected_regions[file_path]
        
        # 如果在自定义文件名中也存在，一并删除
        if file_path in self.custom_filenames:
            del self.custom_filenames[file_path]
        
        # 更新文件列表显示
        self.update_file_list()
        
        # 更新区域列表显示
        self.update_region_list()
        
        # 如果还有文件，选中下一个
        if self.pdf_files:
            next_index = min(index, len(self.pdf_files) - 1)
            self.file_listbox.selection_set(next_index)
            self.load_pdf_file(next_index)
        else:
            # 如果没有文件了，清空显示
            self.pdf_document = None
            self.pdf_path = None
            self.current_page = 0
            self.update_page_display()
            self.split_button.config(state=tk.DISABLED)

    def clear_file_list(self):
        """清空文件列表"""
        self.pdf_files = []
        self.selected_regions = {}
        self.custom_filenames = {}  # 同时清除自定义文件名
        self.pdf_document = None
        self.pdf_path = None
        self.current_page = 0
        
        # 更新界面显示
        self.update_file_list()
        self.update_region_list()  # 确保区域列表也被清空
        self.update_page_display()
        self.split_button.config(state=tk.DISABLED)

    def on_file_select(self, event):
        """文件选择事件处理"""
        selection = self.file_listbox.curselection()
        if not selection:
            return
            
        index = selection[0]
        self.load_pdf_file(index)

    def load_pdf_file(self, index):
        """加载指定索引的PDF文件"""
        if 0 <= index < len(self.pdf_files):
            try:
                if self.pdf_document:
                    self.pdf_document.close()
                
                self.current_file_index = index
                self.pdf_path = self.pdf_files[index]
                self.pdf_document = fitz.open(self.pdf_path)
                self.current_page = 0
                self.update_page_display()
                
                # 如果已设置模板，自动识别页码
                if self.template_coords_set:
                    self.scan_all_pages()
                
            except Exception as e:
                self.showerror("错误", f"无法打开PDF文件: {str(e)}")

    def update_page_display(self):
        if not self.pdf_document:
            return
        
        # 清除当前画布
        self.canvas.delete("all")
        
        # 获取当前页面
        page = self.pdf_document[self.current_page]
        
        # 渲染页面到图像
        zoom = 2.0 * self.scale_factor  # 提高分辨率
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        
        # 将PyMuPDF的Pixmap转换为PIL Image
        img_data = pix.samples
        img = Image.frombytes("RGB", [pix.width, pix.height], img_data)
        
        # 转换为Tkinter可用的PhotoImage
        self.tk_img = ImageTk.PhotoImage(image=img)
        
        # 在画布上显示图像
        self.canvas.create_image(0, 0, image=self.tk_img, anchor=tk.NW, tags="pdf_image")
        
        # 配置画布滚动区域
        self.canvas.config(scrollregion=self.canvas.bbox("all"))
        
        # 更新页面标签
        self.page_label.config(text=f"页面: {self.current_page + 1} / {len(self.pdf_document)}")
        
        # 重新绘制所有选定区域
        self.redraw_regions()
        
        # 记得保持所有模式的状态
        if self.filename_template_mode:
            self.redraw_filename_regions()
    
    def redraw_regions(self):
        """重新绘制当前页面上的所有选定区域"""
        if not self.pdf_path:
            return
        
        # 清除所有现有的区域矩形
        self.canvas.delete("region_*")
            
        # 如果当前文件有选定区域，绘制它们
        if self.pdf_path in self.selected_regions:
            # 获取当前文件的区域列表
            regions = self.selected_regions[self.pdf_path]
            
            # 绘制当前页面上的所有选定区域
            for i, region in enumerate(regions):
                if region['page'] == self.current_page:
                    x1, y1, x2, y2 = region['rect']
                    # 如果是文件名区域，使用蓝色
                    if region.get('is_filename', False):
                        self.canvas.create_rectangle(x1, y1, x2, y2, outline="blue", width=2, tags=f"region_{i}")
                        
                        # 如果存在多个区域坐标，也绘制它们
                        if 'all_coords' in region:
                            for j, coords in enumerate(region['all_coords']):
                                if j > 0:  # 跳过第一个坐标，因为它已经被绘制了
                                    x1, y1, x2, y2 = coords
                                    self.canvas.create_rectangle(x1, y1, x2, y2, outline="blue", width=2, 
                                                              tags=f"region_{i}_part_{j}")
                    else:
                        self.canvas.create_rectangle(x1, y1, x2, y2, outline="green", width=2, tags=f"region_{i}")
        
        # 如果正在设置文件名模板，显示已选择的模板区域
        if self.filename_template_mode:
            for i, coords in enumerate(self.filename_template_coords):
                x1, y1, x2, y2 = coords
                self.canvas.create_rectangle(x1, y1, x2, y2, outline="orange", width=2, 
                                          tags=f"filename_region_{i+1}")
        
        # 如果正在设置页码模板，显示已选择的模板区域
        if self.template_mode == "double" and self.current_page_coords:
            x1, y1, x2, y2 = self.current_page_coords
            self.canvas.create_rectangle(x1, y1, x2, y2, outline="purple", width=2, tags="template_current")
            
            if self.total_pages_coords:
                x1, y1, x2, y2 = self.total_pages_coords
                self.canvas.create_rectangle(x1, y1, x2, y2, outline="purple", width=2, tags="template_total")
        elif self.template_mode == "single" and self.current_page_coords:
            x1, y1, x2, y2 = self.current_page_coords
            self.canvas.create_rectangle(x1, y1, x2, y2, outline="purple", width=2, tags="template_single")
    
    def prev_page(self):
        if self.pdf_document and self.current_page > 0:
            self.current_page -= 1
            self.update_page_display()
    
    def next_page(self):
        if self.pdf_document and self.current_page < len(self.pdf_document) - 1:
            self.current_page += 1
            self.update_page_display()
    
    def zoom_in(self):
        self.scale_factor *= 1.25
        self.update_page_display()
    
    def zoom_out(self):
        self.scale_factor *= 0.8
        self.update_page_display()
    
    def zoom_reset(self):
        self.scale_factor = 1.0
        self.update_page_display()
    
    def on_mouse_down(self, event):
        if not self.pdf_document:
            return
        
        # 获取画布上的坐标
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)
        
        # 创建矩形
        self.rect_id = self.canvas.create_rectangle(
            self.start_x, self.start_y, self.start_x, self.start_y,
            outline="red", width=2, tags="selection"
        )
    
    def on_mouse_drag(self, event):
        if not self.rect_id:
            return
        
        # 获取当前位置
        cur_x = self.canvas.canvasx(event.x)
        cur_y = self.canvas.canvasy(event.y)
        
        # 更新矩形
        self.canvas.coords(self.rect_id, self.start_x, self.start_y, cur_x, cur_y)
    
    def on_mouse_up(self, event):
        if not self.rect_id or not self.pdf_document:
            return
        
        print(f"DEBUG: on_mouse_up - filename_mode={self.filename_template_mode}, selection_step={self.selection_step}") # Add this
        
        # 获取最终位置
        end_x = self.canvas.canvasx(event.x)
        end_y = self.canvas.canvasy(event.y)
        
        # 确保矩形有一定大小
        if abs(end_x - self.start_x) < 10 or abs(end_y - self.start_y) < 10:
            self.canvas.delete(self.rect_id)
            self.rect_id = None
            return
        
        # 获取矩形坐标
        x1 = min(self.start_x, end_x)
        y1 = min(self.start_y, end_y)
        x2 = max(self.start_x, end_x)
        y2 = max(self.start_y, end_y)
        
        if self.filename_template_mode:
            # 文件名模板模式
            self.template_region_count += 1
            self.filename_template_coords.append((x1, y1, x2, y2))
            
            # 提取文本
            text = self.extract_text_from_selection(x1, y1, x2, y2)
            text_preview = text.strip() if text else "无文本"
            
            # 创建一个橙色矩形标记已选区域 (改为橙色以与其他区域区分)
            self.canvas.create_rectangle(x1, y1, x2, y2, outline="orange", width=2, 
                                       tags=f"filename_region_{self.template_region_count}")
            
            # 删除选择框
            self.canvas.delete(self.rect_id)
            self.rect_id = None
            
            # 弹出选择对话框
            self.show_filename_selection_dialog(text_preview)
            return
        
        if self.selection_step > 0:  # 模板选择模式
            if self.template_mode == "double":
                if self.selection_step == 1:  # 选择"第n张"
                    self.current_page_coords = (x1, y1, x2, y2)
                    self.selection_step = 2
                    self.template_status.config(text='请框选"共m张"位置...')
                    self.showinfo('提示', '请框选"共m张"的位置')
                elif self.selection_step == 2:  # 选择"共m张"
                    self.total_pages_coords = (x1, y1, x2, y2)
                    self.template_coords_set = True
                    self.selection_step = 0
                    self.template_status.config(text='正在识别所有页面...')
                    self.scan_all_pages()
            else:  # 单区域模式
                self.current_page_coords = (x1, y1, x2, y2)
                self.template_coords_set = True
                self.selection_step = 0
                self.template_status.config(text='正在识别所有页面...')
                self.scan_all_pages()
        else:  # 普通选择模式
            if not self.pdf_path:
                return
                
            text = self.extract_text_from_selection(x1, y1, x2, y2)
            region = {
                'page': self.current_page,
                'rect': (x1, y1, x2, y2),
                'text': text if text else "",  # 即使没有文本也保存空字符串
                'has_text': bool(text.strip())  # 添加标记表示是否有文本
            }
            
            # 确保当前文件的区域列表已初始化
            if self.pdf_path not in self.selected_regions:
                self.selected_regions[self.pdf_path] = []
            
            # 添加到当前文件的区域列表
            self.selected_regions[self.pdf_path].append(region)
            self.update_region_list()
        
        # 删除临时选择矩形，重新绘制所有区域
        self.canvas.delete("selection")
        self.rect_id = None
        self.redraw_regions()
    
    def _extract_text_from_rect(self, page, rect):
        """Helper method to extract and sort text from a given rectangle on a page."""
        try:
            # Use dict mode for better structure
            blocks = page.get_text("dict", clip=rect)["blocks"]
            
            # Define line tolerance based on zoom (assuming default zoom/scale)
            # TODO: Consider passing scale factor if needed for dynamic tolerance
            line_tolerance = 5 / (2.0 * self.scale_factor) 
            
            all_spans = []
            for block in blocks:
                if block["type"] == 0:  # Text block
                    for line in block["lines"]:
                        for span in line["spans"]:
                            y_pos = (span["bbox"][1] + span["bbox"][3]) / 2
                            x_pos = span["bbox"][0]
                            all_spans.append((y_pos, x_pos, span["text"]))
            
            if not all_spans:
                return ""
                
            # Sort primarily by Y, then X
            all_spans.sort(key=lambda x: (x[0], x[1]))
            
            # Group spans into lines based on Y-coordinate and tolerance
            lines = []
            if all_spans:
                current_line = [all_spans[0]]
                last_y = all_spans[0][0]
                
                for span in all_spans[1:]:
                    span_y = span[0]
                    if abs(span_y - last_y) < line_tolerance:
                        current_line.append(span)
                    else:
                        # Sort current line by X before adding
                        current_line.sort(key=lambda x: x[1])
                        lines.append(current_line)
                        # Start new line
                        current_line = [span]
                        last_y = span_y
                
                # Add the last line
                current_line.sort(key=lambda x: x[1])
                lines.append(current_line)
            
            # Join text within lines and then join lines
            result_text = " ".join([" ".join(span[2] for span in line) for line in lines])
            
        except Exception as e:
            print(f"Warning: Text extraction using dict failed, falling back. Error: {e}")
            # Fallback to simple text extraction
            try:
                result_text = page.get_text("text", clip=rect)
            except Exception as fallback_e:
                print(f"Error: Fallback text extraction failed. Error: {fallback_e}")
                result_text = ""
        
        # Clean up whitespace
        text = result_text.replace('\\n', ' ').replace('\\r', ' ')
        text = re.sub(r'\\s+', ' ', text).strip()
        return text

    def extract_text_from_selection(self, x1, y1, x2, y2):
        if not self.pdf_document:
            return ""
        
        # Get current page
        page = self.pdf_document[self.current_page]
        
        # Calculate selection rectangle in PDF coordinates
        zoom = 2.0 * self.scale_factor
        pdf_rect = fitz.Rect(x1 / zoom, y1 / zoom, x2 / zoom, y2 / zoom)
        
        return self._extract_text_from_rect(page, pdf_rect)

    def update_region_list(self):
        """更新区域列表显示"""
        # 清除列表
        self.region_listbox.delete(0, tk.END)
        
        # 存储列表索引到文件路径和页码的映射
        self.region_map = {}
        current_index = 0
        
        # 检查是否有任何区域
        has_any_regions = False
        for file_path in self.pdf_files:
            if file_path in self.selected_regions and self.selected_regions[file_path]:
                has_any_regions = True
                break
        
        if not self.pdf_files:
            # 如果没有文件，显示提示信息
            self.region_listbox.insert(tk.END, "请先添加PDF文件")
            return
        elif not has_any_regions:
            # 如果有文件但没有区域，显示提示信息
            self.region_listbox.insert(tk.END, "请在PDF上选择区域")
            return
        
        # 按图幅尺寸分组显示
        for file_path in self.pdf_files:
            file_name = os.path.basename(file_path)
            regions = self.selected_regions.get(file_path, [])
            
            if not regions:
                continue  # 跳过没有区域的文件
            
            # 添加文件名作为标题
            self.region_listbox.insert(tk.END, f"=== {file_name} ===")
            current_index += 1
            
            # 如果有图幅信息，按图幅分组显示
            if file_path in self.page_sizes and self.page_sizes[file_path]:
                # 创建页码到尺寸的映射
                page_to_size = {}
                for size_key, pages in self.page_sizes[file_path].items():
                    for page_num in pages:
                        page_to_size[page_num] = size_key
                
                # 按图幅分组区域
                size_groups = {}
                for region in regions:
                    page_num = region['page']
                    size_key = page_to_size.get(page_num, "未知尺寸")
                    if size_key not in size_groups:
                        size_groups[size_key] = []
                    size_groups[size_key].append(region)
                
                # 显示每个图幅组的区域
                for size_key in sorted(size_groups.keys()):
                    # 添加图幅尺寸标题
                    template_status = "（已设置模板）" if size_key in self.page_templates else "（未设置模板）"
                    self.region_listbox.insert(tk.END, f"--- 图幅：{size_key} {template_status} ---")
                    current_index += 1
                    
                    # 显示该图幅下的区域
                    for region in sorted(size_groups[size_key], key=lambda x: x['page']):
                        page_num = region['page'] + 1
                        
                        if 'current_page' in region and 'filename' in region:
                            # 既有页码又有文件名
                            item_text = f"页面 {page_num}: [第 {region['current_page']} 张] [{region['filename']}]"
                        elif region.get('is_filename'):
                            # 只有文件名
                            item_text = f"页面 {page_num}: [文件名] {region['filename']}"
                        elif 'current_page' in region and 'total_pages' in region:
                            # 只有页码
                            item_text = f"页面 {page_num}: 第 {region['current_page']} 张 共 {region['total_pages']} 张"
                        else:
                            # 普通区域
                            text_preview = region['text'].strip()
                            if not text_preview:
                                item_text = f"页面 {page_num}: [无内容]"
                            else:
                                if len(text_preview) > 30:
                                    text_preview = text_preview[:30] + "..."
                                item_text = f"页面 {page_num}: {text_preview}"
                        
                        self.region_listbox.insert(tk.END, item_text)
                        # 保存映射关系
                        self.region_map[current_index] = (file_path, region['page'])
                        current_index += 1
                    
                    # 添加空行分隔不同图幅
                    self.region_listbox.insert(tk.END, "")
                    current_index += 1
            else:
                # 如果没有图幅信息，按原来的方式显示
                for region in sorted(regions, key=lambda x: x['page']):
                    page_num = region['page'] + 1
                    
                    if 'current_page' in region and 'filename' in region:
                        item_text = f"页面 {page_num}: [第 {region['current_page']} 张] [{region['filename']}]"
                    elif region.get('is_filename'):
                        item_text = f"页面 {page_num}: [文件名] {region['filename']}"
                    elif 'current_page' in region and 'total_pages' in region:
                        item_text = f"页面 {page_num}: 第 {region['current_page']} 张 共 {region['total_pages']} 张"
                    else:
                        text_preview = region['text'].strip()
                        if not text_preview:
                            item_text = f"页面 {page_num}: [无内容]"
                        else:
                            if len(text_preview) > 30:
                                text_preview = text_preview[:30] + "..."
                            item_text = f"页面 {page_num}: {text_preview}"
                    
                    self.region_listbox.insert(tk.END, item_text)
                    # 保存映射关系
                    self.region_map[current_index] = (file_path, region['page'])
                    current_index += 1
            
            # 添加空行分隔不同文件
            self.region_listbox.insert(tk.END, "")
            current_index += 1
    
    def on_region_select(self, event):
        """处理区域列表的单击选择事件"""
        # 获取选中的索引
        selection = self.region_listbox.curselection()
        if not selection:
            return
        
        index = selection[0]
        if index not in self.region_map:
            return  # 如果选择的是标题行或空行，则不做处理
            
        # 获取文件路径和页码
        file_path, page_num = self.region_map[index]
        
        # 如果当前文件与选中的不同，加载新文件
        if self.pdf_path != file_path:
            self.load_pdf_file(self.pdf_files.index(file_path))
        
        # 跳转到指定页面
        self.current_page = page_num
        self.update_page_display()
    
    def on_region_double_click(self, event):
        """处理区域列表的双击事件"""
        self.edit_selected_content()
    
    def edit_selected_content(self):
        """编辑选中区域的内容"""
        selection = self.region_listbox.curselection()
        if not selection:
            self.showwarning("警告", "请先选择一个区域")
            return
            
        index = selection[0]
        if index not in self.region_map:
            return
            
        # 获取文件路径和页码
        file_path, page_num = self.region_map[index]
        
        # 找到对应的区域
        region = None
        for r in self.selected_regions[file_path]:
            if r['page'] == page_num:
                region = r
                break
                
        if not region:
            return
            
        # 显示编辑对话框
        self.show_content_edit_dialog(file_path, page_num, region)
    
    def show_content_edit_dialog(self, file_path, page_num, region):
        """显示内容编辑对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("编辑页面内容")
        dialog.geometry("400x200")
        dialog.transient(self.root)
        dialog.grab_set()  # 模态对话框
        
        # 创建主框架
        main_frame = ttk.Frame(dialog, padding=0)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建标题栏 - 蓝色
        title_frame = tk.Frame(main_frame, height=6, bg="#3498db")
        title_frame.pack(fill=tk.X)
        
        # 内容区域
        content_frame = ttk.Frame(main_frame, padding=(20, 15, 20, 10))
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        header_label = ttk.Label(content_frame, 
                               text=f"编辑第 {page_num + 1} 页内容", 
                               font=("Arial", 12, "bold"))
        header_label.pack(anchor=tk.W, pady=(0, 10))
        
        # 文本输入框
        text_frame = ttk.Frame(content_frame)
        text_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        text_entry = ttk.Entry(text_frame)
        text_entry.pack(fill=tk.X, padx=5, pady=5)
        
        # 设置当前内容
        current_text = ""
        if region.get('is_filename', False):
            current_text = region.get('filename', '')
            if current_text == '[无内容]':
                current_text = ''
        else:
            current_text = region.get('text', '')
            if current_text == '[无内容]':
                current_text = ''
        text_entry.insert(0, current_text)
        text_entry.focus_set()  # 设置焦点
        text_entry.select_range(0, tk.END)  # 选中所有文本
        
        # 分隔线
        separator = ttk.Separator(main_frame, orient="horizontal")
        separator.pack(fill=tk.X, padx=0, pady=0)
        
        # 按钮区域
        btn_frame = ttk.Frame(main_frame, padding=(15, 10))
        btn_frame.pack(fill=tk.X)
        
        def on_save():
            new_text = text_entry.get().strip()
            if not new_text:
                self.showwarning("警告", "请输入内容")
                return
                
            # 更新区域信息
            if region.get('is_filename', False):
                # 如果是文件名区域
                region['filename'] = new_text
                region['filename_text'] = f"文件名: {new_text}"
                region['text'] = f"文件名: {new_text}"
                
                # 更新自定义文件名字典
                if file_path not in self.custom_filenames:
                    self.custom_filenames[file_path] = {}
                self.custom_filenames[file_path][page_num] = new_text
            else:
                # 如果是普通区域
                region['text'] = new_text
            
            # 更新显示
            self.update_region_list()
            dialog.destroy()
            
            # 显示成功消息
            self.showinfo("成功", "内容已更新")
        
        def on_cancel():
            dialog.destroy()
        
        style = ttk.Style()
        style.configure("Accent.TButton", font=("Arial", 10))
        
        # 取消按钮
        cancel_btn = ttk.Button(btn_frame, text="取消", command=on_cancel)
        cancel_btn.pack(side=tk.LEFT, padx=5)
        
        # 保存按钮
        save_btn = ttk.Button(btn_frame, text="保存", style="Accent.TButton", command=on_save)
        save_btn.pack(side=tk.RIGHT, padx=5)
        
        # 绑定回车键到保存按钮
        dialog.bind("<Return>", lambda e: save_btn.invoke())
        dialog.bind("<Escape>", lambda e: cancel_btn.invoke())
        
        # 居中显示对话框
        self.center_dialog(dialog)
        
        dialog.wait_window()
    
    def remove_region(self):
        """删除选中的区域"""
        selection = self.region_listbox.curselection()
        if not selection:
            return
        
        index = selection[0]
        
        # 找到对应的文件和区域
        current_index = 0
        for file_path in self.pdf_files:
            regions = self.selected_regions.get(file_path, [])
            if index - current_index < len(regions) + 1:  # +1 是因为文件标题行
                region_index = index - current_index - 1  # -1 是因为文件标题行
                if region_index >= 0:
                    del self.selected_regions[file_path][region_index]
                break
            current_index += len(regions) + 2  # +2 是因为文件标题行和空行
        
        self.update_region_list()
        self.update_page_display()
    
    def clear_regions(self):
        """清除所有区域"""
        if self.pdf_path:
            self.selected_regions[self.pdf_path] = []
        self.update_region_list()
        self.update_page_display()
        # Removing the line below which is causing the error
        # self.update_text_display("选择一个区域以查看提取的文本")
    
    def on_page_mode_change(self, event):
        # 根据页码模式启用或禁用手动输入控件
        is_manual = (self.page_mode.get() == "手动指定")
        
        state = tk.NORMAL if is_manual else tk.DISABLED
        for child in self.manual_frame.winfo_children():
            try:
                child.config(state=state)
            except:
                pass  # 某些控件可能没有state属性
        
        self.apply_button.config(state=state)
    
    def apply_page_info(self):
        try:
            current_page = self.current_page_var.get()
            total_pages = self.total_pages_var.get()
            
            if current_page <= 0 or total_pages <= 0 or current_page > total_pages:
                self.showerror("错误", "页码信息无效")
                return
            
            # 创建一个新的手动区域
            manual_region = {
                'page': self.current_page,
                'rect': (10, 10, 100, 30),  # 一个小区域，仅用于标记
                'text': f"第 {current_page} 张 共 {total_pages} 张",
                'manual': True,
                'current_page': current_page,
                'total_pages': total_pages
            }
            
            # 添加到区域列表
            self.selected_regions.append(manual_region)
            self.update_region_list()
            
            self.showinfo("页码信息", f"已为页面 {self.current_page + 1} 设置页码信息: 第 {current_page} 张 共 {total_pages} 张")
        except Exception as e:
            self.showerror("错误", f"设置页码信息失败: {str(e)}")
    
    def scan_all_pages(self):
        """遍历所有PDF文件并识别页码"""
        if not self.current_page_coords:  # 至少需要一个坐标
            return
            
        try:
            # 显示进度窗口
            progress_window = tk.Toplevel(self.root)
            progress_window.title("识别进度")
            progress_window.geometry("300x150")
            progress_window.transient(self.root)  # 设置为主窗口的子窗口
            self.center_dialog(progress_window)  # 居中显示
            
            progress_label = ttk.Label(progress_window, text="正在识别页码...")
            progress_label.pack(pady=10)
            
            progress_var = tk.DoubleVar()
            progress_bar = ttk.Progressbar(progress_window, variable=progress_var, maximum=100)
            progress_bar.pack(fill=tk.X, padx=20, pady=10)
            
            # 计算总页数和需要处理的页面
            total_pages_to_process = 0
            pages_to_process = []  # [(file_path, page_num), ...]
            
            for file_path in self.pdf_files:
                if self.current_size_key:
                    # 如果是为特定图幅设置模板，只处理该图幅的页面
                    if self.current_size_key in self.page_sizes[file_path]:
                        page_nums = self.page_sizes[file_path][self.current_size_key]
                        total_pages_to_process += len(page_nums)
                        pages_to_process.extend([(file_path, page_num) for page_num in page_nums])
                else:
                    # 如果不是特定图幅，处理所有页面
                    pdf = fitz.open(file_path)
                    total_pages_to_process += len(pdf)
                    pages_to_process.extend([(file_path, i) for i in range(len(pdf))])
                    pdf.close()
            
            processed_pages = 0
            total_pages_recognized = 0
            
            # 保存现有的文件名区域
            saved_filename_regions = {}
            for file_path in self.pdf_files:
                if file_path in self.selected_regions:
                    filename_regions = [region for region in self.selected_regions[file_path] 
                                      if region.get('is_filename', False)]
                    if filename_regions:
                        saved_filename_regions[file_path] = filename_regions
            
            # 处理每个页面
            current_pdf = None
            current_file = None
            
            for file_path, page_num in pages_to_process:
                # 如果是新文件，关闭旧文件并打开新文件
                if current_file != file_path:
                    if current_pdf:
                        current_pdf.close()
                    current_pdf = fitz.open(file_path)
                    current_file = file_path
                    
                    # 更新进度标签
                    progress_label.config(text=f"正在处理: {os.path.basename(file_path)}")
                    progress_window.update()
                    
                    # 清除当前文件的页码识别结果，但保留文件名区域和其他图幅的结果
                    if file_path in self.selected_regions:
                        if self.current_size_key:
                            # 如果是特定图幅，只清除该图幅的页面结果
                            self.selected_regions[file_path] = [
                                region for region in self.selected_regions[file_path]
                                if region.get('is_filename', False) or  # 保留文件名区域
                                page_num not in self.page_sizes[file_path].get(self.current_size_key, [])  # 保留其他图幅的结果
                            ]
                        else:
                            # 如果不是特定图幅，保留文件名区域
                            self.selected_regions[file_path] = [
                                region for region in self.selected_regions[file_path]
                                if region.get('is_filename', False)
                            ]
                    else:
                        self.selected_regions[file_path] = []
                
                # 获取页面
                page = current_pdf[page_num]
                
                # 计算选择区域在PDF坐标系中的位置
                zoom = 2.0 * self.scale_factor
                
                # 提取页码文本
                x1, y1, x2, y2 = self.current_page_coords
                pdf_rect1 = fitz.Rect(x1/zoom, y1/zoom, x2/zoom, y2/zoom)
                current_text = page.get_text("text", clip=pdf_rect1)
                
                # 提取当前页码
                current_number = self.extract_number_from_text(current_text)
                
                if self.template_mode == "double" and self.total_pages_coords:
                    # 提取总页数
                    x1, y1, x2, y2 = self.total_pages_coords
            # 处理每个PDF文件
            for file_path in self.pdf_files:
                # 清除当前文件的页码识别结果，但保留文件名区域
                if file_path in self.selected_regions:
                    # 只保留文件名区域
                    self.selected_regions[file_path] = [region for region in self.selected_regions[file_path] 
                                                     if region.get('is_filename', False)]
                else:
                    self.selected_regions[file_path] = []
                
                # 打开PDF文件
                current_pdf = fitz.open(file_path)
                file_pages = len(current_pdf)
                
                progress_label.config(text=f"正在处理: {os.path.basename(file_path)}")
                progress_window.update()
                
                # 遍历当前文件的所有页面
                for page_num in range(file_pages):
                    # 获取页面
                    page = current_pdf[page_num]
                    
                    # 计算选择区域在PDF坐标系中的位置
                    zoom = 2.0 * self.scale_factor
                    
                    # 提取页码文本
                    x1, y1, x2, y2 = self.current_page_coords
                    pdf_rect1 = fitz.Rect(x1/zoom, y1/zoom, x2/zoom, y2/zoom)
                    current_text = page.get_text("text", clip=pdf_rect1)
                    
                    # 提取当前页码
                    current_number = self.extract_number_from_text(current_text)
                    
                    if self.template_mode == "double" and self.total_pages_coords:
                        # 提取总页数
                        x1, y1, x2, y2 = self.total_pages_coords
                        pdf_rect2 = fitz.Rect(x1/zoom, y1/zoom, x2/zoom, y2/zoom)
                        total_text = page.get_text("text", clip=pdf_rect2)
                        total_number = self.extract_number_from_text(total_text)
                    else:
                        total_number = None
                    
                    if current_number:  # 只要有当前页码就记录
                        # 查找这个页面是否已有区域（例如文件名区域）
                        existing_region = None
                        for i, region in enumerate(self.selected_regions[file_path]):
                            if region['page'] == page_num:
                                existing_region = region
                                existing_index = i
                                break
                        
                        text = f"第 {current_number} 张" + (f" 共 {total_number} 张" if total_number else "")
                        
                        if existing_region:
                            # 如果已有区域，将页码信息添加到现有区域
                            self.selected_regions[file_path][existing_index].update({
                                'current_page': current_number,
                                'total_pages': total_number if total_number else file_pages,
                                'page_text': text,
                            })
                            # 如果不是文件名区域，还要更新text字段
                            if not existing_region.get('is_filename', False):
                                self.selected_regions[file_path][existing_index]['text'] = text
                        else:
                            # 创建新区域
                            region = {
                                'page': page_num,
                                'rect': self.current_page_coords,
                                'text': text,
                                'current_page': current_number,
                                'total_pages': total_number if total_number else file_pages,
                                'page_text': text,
                            }
                            self.selected_regions[file_path].append(region)
                            print(f"DEBUG: Added region for {file_path}, page {page_num}: {region}") # Add this line
                            self.update_region_list()
                        
                        total_pages_recognized += 1
                    
                    total_pages_processed += 1
                    # 更新进度
                    progress = (processed_files * 100.0 / total_files) + (page_num + 1) * 100.0 / (file_pages * total_files)
                    progress_var.set(progress)
                    progress_window.update()
                
                current_pdf.close()
                processed_files += 1
            
            progress_window.destroy()
            
            # 更新区域列表显示
            self.update_region_list()
            
            # 显示识别结果
            mode_text = "双区域" if self.template_mode == "double" else "单区域"
            self.template_status.config(text=f'{mode_text}模板设置完成 (已识别 {total_pages_recognized}/{total_pages_processed} 页)')
            self.showinfo('完成', f'所有文件页码识别完成！\n共识别出 {total_pages_recognized}/{total_pages_processed} 页的页码信息。')
            
        except Exception as e:
            self.showerror('错误', f'识别页码时出错：{str(e)}')
            if 'progress_window' in locals() and progress_window.winfo_exists():
                progress_window.destroy()
    
    def extract_number_from_text(self, text):
        """从文本中提取数字"""
        numbers = re.findall(r'\d+', text)
        return int(numbers[0]) if numbers else None
    
    def start_template_selection(self, mode="double"):
        """开始模板选择过程"""
        if not self.pdf_document:
            self.showwarning('警告', '请先打开PDF文件')
            return
        
        # 保存文件名模板状态，以便页码模板选择完成后恢复
        saved_filename_template_mode = self.filename_template_mode 
        saved_filename_template_coords = self.filename_template_coords.copy()
        
        # 暂时关闭文件名模板模式
        self.filename_template_mode = False
            
        self.selection_step = 1
        self.template_coords_set = False
        self.current_page_coords = None
        self.total_pages_coords = None
        self.template_mode = mode  # 记录当前模板模式
        
        if mode == "double":
            self.template_status.config(text='请框选"第n张"位置...')
            self.showinfo('提示', '请框选"第n张"的位置')
        else:
            self.template_status.config(text='请框选页码位置...')
            self.showinfo('提示', '请框选页码位置')
            
        # 如果正在为特定尺寸设置模板，更新状态显示
        if self.current_size_key:
            self.template_status.config(text=f'正在为图幅 {self.current_size_key} 设置模板...')
            
        # 下面的操作将在鼠标操作完成后执行
    
    def split_pdf(self):
        """分割所有PDF文件"""
        if not self.pdf_files:
            self.showwarning("警告", "请先添加PDF文件")
            return
        
        try:
            # 确定输出目录
            output_dir = self.output_dir
            if not output_dir:
                output_dir = os.path.dirname(self.pdf_files[0]) or '.'
            
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # 显示总进度窗口
            progress_window = tk.Toplevel(self.root)
            progress_window.title("处理进度")
            progress_window.geometry("300x150")
            progress_window.transient(self.root)  # 设置为主窗口的子窗口
            self.center_dialog(progress_window)  # 居中显示
            
            progress_label = ttk.Label(progress_window, text="正在处理文件...")
            progress_label.pack(pady=10)
            
            progress_var = tk.DoubleVar()
            progress_bar = ttk.Progressbar(progress_window, variable=progress_var, maximum=100)
            progress_bar.pack(fill=tk.X, padx=20, pady=10)
            
            total_files = len(self.pdf_files)
            processed_files = 0
            
            for file_path in self.pdf_files:
                progress_label.config(text=f"正在处理: {os.path.basename(file_path)}")
                progress_window.update()
                
                # 处理单个文件
                self.process_single_file(file_path, output_dir, progress_window)
                
                processed_files += 1
                progress_var.set(processed_files * 100 / total_files)
            
            progress_window.destroy()
            self.showinfo("完成", f"所有PDF文件处理完成\n保存到: {output_dir}")
            
        except Exception as e:
            self.showerror("错误", f"处理文件时出错：{str(e)}")
            if 'progress_window' in locals() and progress_window.winfo_exists():
                progress_window.destroy()

    def process_single_file(self, file_path, output_dir, progress_window=None):
        """处理单个PDF文件"""
        # 获取该文件的区域信息
        regions = self.selected_regions.get(file_path, [])
        custom_filenames = self.custom_filenames.get(file_path, {})
        
        # 打开PDF文件
        pdf_document = fitz.open(file_path)
        total_pages = len(pdf_document)
        
        try:
            # 如果没有任何区域信息，但有自定义文件名，则尝试根据自定义文件名分割
            if (not regions or all(region.get('is_filename', False) for region in regions)) and custom_filenames:
                result = self.process_by_filenames(file_path, output_dir, custom_filenames, progress_window)
                if result:  # 如果成功处理了，直接返回
                    return
                # 如果处理失败（比如所有页面都是无内容），继续执行下面的逻辑
            
            # 如果只有一页，且没有识别到任何内容或区域，直接复制整个文件
            if total_pages == 1 and (not regions or all(not region.get('text', '').strip() for region in regions)):
                output_filename = os.path.splitext(os.path.basename(file_path))[0] + "_完整文件.pdf"
                output_path = os.path.join(output_dir, output_filename)
                
                # 直接复制整个文件
                output_doc = fitz.open()
                output_doc.insert_pdf(pdf_document)
                self.optimize_and_save_pdf(output_doc, output_path)
                output_doc.close()
                return
            
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            
            if self.template_mode == "double":
                # 双区域模式的处理逻辑
                documents = self.process_double_mode(regions)
            else:
                # 单区域模式的处理逻辑
                documents = self.process_single_mode(regions)
            
            # 创建分割后的PDF文件
            for i, doc_info in enumerate(documents, 1):
                # 更新进度信息
                if progress_window:
                    progress_window.update()
                
                output_doc = fitz.open()
                
                if self.template_mode == "double":
                    total_pages, pages = doc_info
                    for page_index, current_page in pages:
                        output_doc.insert_pdf(pdf_document, from_page=page_index, to_page=page_index)
                    first_page = pages[0][1]
                    last_page = pages[-1][1]
                    
                    # 检查是否有自定义文件名
                    custom_name = None
                    if file_path in self.custom_filenames and pages[0][0] in self.custom_filenames[file_path]:
                        custom_name = self.custom_filenames[file_path][pages[0][0]]
                    
                    if custom_name and custom_name != "[无内容]":
                        output_filename = f"{custom_name}.pdf"
                    else:
                        output_filename = f"{base_name}_{i:02d}_第{first_page}至{last_page}张共{total_pages}张.pdf"
                else:
                    start_page, end_page, page_numbers = doc_info
                    for page_index in range(start_page, end_page + 1):
                        output_doc.insert_pdf(pdf_document, from_page=page_index, to_page=page_index)
                    
                    # 检查是否有自定义文件名
                    custom_name = None
                    if file_path in self.custom_filenames and start_page in self.custom_filenames[file_path]:
                        custom_name = self.custom_filenames[file_path][start_page]
                    
                    if custom_name and custom_name != "[无内容]":
                        output_filename = f"{custom_name}.pdf"
                    else:
                        output_filename = f"{base_name}_{i:02d}_第{page_numbers[0]}至{page_numbers[-1]}张.pdf"
                    
                    # 使用更优化的参数保存PDF
                    output_path = os.path.join(output_dir, output_filename)
                    
                    # 调用自定义优化方法
                    self.optimize_and_save_pdf(output_doc, output_path)
                    
                    output_doc.close()
            
        finally:
            pdf_document.close()

    def process_by_filenames(self, file_path, output_dir, custom_filenames, parent_progress_window=None):
        """根据连续的相同文件名来分割PDF"""
        try:
            # 打开PDF文件
            pdf_document = fitz.open(file_path)
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            
            # 过滤掉标记为无内容的页面
            valid_filenames = {page_num: name for page_num, name in custom_filenames.items() 
                             if name != "[无内容]"}
            
            if not valid_filenames:
                self.showwarning("警告", "没有找到有效的文件名，无法进行分割。")
                return False
            
            # 按页码排序自定义文件名
            sorted_pages = sorted(valid_filenames.keys())
            
            if not sorted_pages:
                return False
            
            # 按文件名组织页面
            current_name = valid_filenames[sorted_pages[0]]
            start_page = sorted_pages[0]
            groups = []
            
            # 遍历所有页面，将连续相同文件名的页面分为一组
            for i in range(1, len(sorted_pages)):
                page_num = sorted_pages[i]
                current_filename = valid_filenames[page_num]
                
                # 如果文件名变化或者页面不连续，创建一个新组
                if current_filename != current_name or page_num != sorted_pages[i-1] + 1:
                    groups.append((start_page, sorted_pages[i-1], current_name))
                    start_page = page_num
                    current_name = current_filename
            
            # 添加最后一组
            groups.append((start_page, sorted_pages[-1], current_name))
            
            # 使用父进度窗口或创建新窗口
            local_progress_window = None
            progress_label = None
            progress_var = None
            
            if parent_progress_window:
                # 使用父窗口中的标签
                progress_label = parent_progress_window.winfo_children()[0]
                progress_var = None  # 使用父窗口的进度条
            else:
                # 创建新的进度窗口
                local_progress_window = tk.Toplevel(self.root)
                local_progress_window.title("生成PDF中")
                local_progress_window.geometry("300x150")
                local_progress_window.transient(self.root)  # 设置为主窗口的子窗口
                self.center_dialog(local_progress_window)  # 居中显示
                
                progress_label = ttk.Label(local_progress_window, text="正在生成PDF文件...")
                progress_label.pack(pady=10)
                
                progress_var = tk.DoubleVar()
                progress_bar = ttk.Progressbar(local_progress_window, variable=progress_var, maximum=100)
                progress_bar.pack(fill=tk.X, padx=20, pady=10)
            
            # 创建分割后的PDF文件
            for i, (start, end, name) in enumerate(groups, 1):
                # 更新进度标签
                if progress_label:
                    progress_label.config(text=f"正在生成: {name}.pdf ({i}/{len(groups)})")
                
                # 更新进度条
                if progress_var:
                    progress_var.set(i * 100 / len(groups))
                
                # 更新窗口
                if local_progress_window:
                    local_progress_window.update()
                elif parent_progress_window:
                    parent_progress_window.update()
                
                output_doc = fitz.open()
                
                # 添加页面到新文档
                for page_index in range(start, end + 1):
                    output_doc.insert_pdf(pdf_document, from_page=page_index, to_page=page_index)
                
                # 使用识别的文件名
                output_filename = f"{name}.pdf"
                
                # 清理文件名中的非法字符
                output_filename = re.sub(r'[\\/*?:"<>|]', '_', output_filename)
                
                output_path = os.path.join(output_dir, output_filename)
                
                # 使用优化方法
                self.optimize_and_save_pdf(output_doc, output_path)
                
                output_doc.close()
            
            if local_progress_window:
                local_progress_window.destroy()
                
            pdf_document.close()
            
            # 只有在使用本地窗口时才显示完成消息
            if not parent_progress_window:
                self.showinfo("完成", f"根据文件名分割完成，共生成 {len(groups)} 个文件。")
            
            return True
            
        except Exception as e:
            self.showerror("错误", f"按文件名分割时出错：{str(e)}")
            if 'local_progress_window' in locals() and local_progress_window:
                local_progress_window.destroy()
            return False
            
    def optimize_and_save_pdf(self, doc, output_path):
        """优化PDF文件并保存
        
        应用多种优化技术来减小PDF文件大小：
        - 垃圾回收：移除未使用的对象
        - 压缩：对PDF内容进行压缩
        - 清理：移除冗余对象
        - 优化图像：对图像应用压缩
        """
        try:
            # 第一步：优化每一页中的图像
            for page_num in range(len(doc)):
                page = doc[page_num]
                
                # 获取页面上的图像对象
                img_list = page.get_images(full=True)
                for img_index, img_info in enumerate(img_list):
                    try:
                        xref = img_info[0]  # 图像的xref号
                        
                        # 获取图像数据
                        base_image = doc.extract_image(xref)
                        if not base_image:
                            continue
                            
                        image_bytes = base_image["image"]
                        image_ext = base_image["ext"]
                        
                        # 仅处理像素图像格式（不处理矢量图像）
                        if image_ext.lower() in ("jpg", "jpeg", "png"):
                            # 使用PIL打开图像
                            img = Image.open(io.BytesIO(image_bytes))
                            
                            # 确定合适的压缩质量
                            # 根据图像大小动态调整质量：越大的图像使用越低的质量
                            img_size = len(image_bytes) / 1024  # KB
                            
                            if img_size > 1000:  # > 1MB
                                quality = 65
                            elif img_size > 500:  # > 500KB
                                quality = 75
                            elif img_size > 100:  # > 100KB
                                quality = 80
                            else:
                                quality = 85
                            
                            # 创建内存缓冲区存储压缩图像
                            output_buffer = io.BytesIO()
                            if image_ext.lower() in ("jpg", "jpeg"):
                                img.save(output_buffer, format="JPEG", quality=quality, optimize=True)
                            elif image_ext.lower() == "png":
                                img.save(output_buffer, format="PNG", optimize=True, 
                                       compress_level=9)  # 最高压缩级别
                            
                            compressed_bytes = output_buffer.getvalue()
                            
                            # 如果压缩成功并且新的图像比原始图像小，则替换
                            if len(compressed_bytes) < len(image_bytes):
                                # 替换原图像
                                doc.update_image(xref, compressed_bytes)
                    except Exception:
                        # 如果处理单个图像出错，继续处理其他图像
                        continue
            
            # 第二步：使用PyMuPDF的内置优化功能保存PDF
            doc.save(
                output_path,
                garbage=4,         # 最高级别的垃圾收集
                deflate=True,      # 使用压缩
                clean=True,        # 清理和优化结构
                pretty=False,      # 不美化（节省空间）
                ascii=False,       # 使用二进制格式（更紧凑）
                compress=True      # 压缩文件流
            )
            
            return True
        
        except Exception as e:
            print(f"优化PDF时出错: {str(e)}")
            # 如果优化失败，使用基本设置保存
            doc.save(output_path, garbage=4, deflate=True)
            return False

    def process_double_mode(self, regions):
        """处理双区域模式的页面分组"""
        documents = []
        temp_docs = {}
        
        # 对页面进行分组
        for region in regions:
            if 'current_page' not in region:
                continue
                
            current_page = region['current_page']
            total_pages = region['total_pages']
            page_index = region['page']
            
            doc_start_index = page_index - (current_page - 1)
            doc_key = (total_pages, doc_start_index)
            
            if doc_key not in temp_docs:
                temp_docs[doc_key] = []
            temp_docs[doc_key].append((page_index, current_page))
        
        # 将分组后的文档转换为列表
        for (total_pages, _), pages in temp_docs.items():
            pages.sort(key=lambda x: x[1])
            documents.append((total_pages, pages))
        
        # 按文档第一页的页码排序
        documents.sort(key=lambda x: x[1][0][1])
        return documents

    def process_single_mode(self, regions):
        """处理单区域模式的页面分组"""
        documents = []
        current_group = []
        last_page_number = None
        
        # 按页面索引排序
        sorted_regions = sorted(regions, key=lambda x: x['page'])
        
        for region in sorted_regions:
            current_number = region['current_page']
            
            # 如果是第一个页码或者页码序列断开（重新从1开始），创建新组
            if last_page_number is None or current_number == 1:
                if current_group:
                    start_page = current_group[0]['page']
                    end_page = current_group[-1]['page']
                    page_numbers = [r['current_page'] for r in current_group]
                    documents.append((start_page, end_page, page_numbers))
                current_group = [region]
            else:
                current_group.append(region)
            
            last_page_number = current_number
        
        # 处理最后一组
        if current_group:
            start_page = current_group[0]['page']
            end_page = current_group[-1]['page']
            page_numbers = [r['current_page'] for r in current_group]
            documents.append((start_page, end_page, page_numbers))
        
        return documents

    def set_output_dir(self):
        """设置输出目录"""
        dir_path = filedialog.askdirectory(
            title="选择输出目录",
            initialdir=self.output_dir if self.output_dir else os.path.expanduser("~")
        )
        if dir_path:
            self.output_dir = dir_path
            self.showinfo("提示", f"输出目录已设置为：\n{dir_path}")

    def start_filename_template_selection(self):
        """开始文件名模板选择过程"""
        if not self.pdf_document:
            self.showwarning('警告', '请先打开PDF文件')
            return
            
        self.filename_template_mode = True
        self.filename_template_coords = []  # 重置为空列表
        self.template_region_count = 0
        
        # 清除之前的选择框
        self.canvas.delete("filename_region_*")
        
        # 如果当前正在为特定图幅尺寸设置模板，显示相应的提示
        if self.current_size_key:
            self.template_status.config(text=f"正在为图幅 {self.current_size_key} 设置文件名提取区域...")
            self.showinfo('提示', f'请为图幅 {self.current_size_key} 框选要作为文件名的文本区域')
        else:
            self.template_status.config(text="正在设置文件名提取区域...")
            self.showinfo('提示', '请框选要作为文件名的文本区域')

    def show_filename_selection_dialog(self, text_preview):
        """显示文件名选择对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("文件名区域选择")
        dialog.geometry("440x240")
        dialog.transient(self.root)
        dialog.grab_set()  # 模态对话框
        
        # 创建主框架
        main_frame = ttk.Frame(dialog, padding=0)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建标题栏 - 蓝色
        title_frame = tk.Frame(main_frame, height=6, bg="#3498db")
        title_frame.pack(fill=tk.X)
        
        # 内容区域
        content_frame = ttk.Frame(main_frame, padding=(20, 15, 20, 10))
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        header_text = f"已选择区域 {self.template_region_count}"
        header_label = ttk.Label(content_frame, text=header_text, font=("Arial", 12, "bold"))
        header_label.pack(anchor=tk.W, pady=(0, 5))
        
        # 识别文本标签
        text_label = ttk.Label(content_frame, text="识别文本：", font=("Arial", 10))
        text_label.pack(anchor=tk.W, pady=(5, 0))
        
        # 文本框架
        text_frame = ttk.Frame(content_frame, relief="solid", borderwidth=1)
        text_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 文本预览 - 使用Text组件代替Label以支持更好的文本显示
        text_display = tk.Text(text_frame, height=2, width=40, wrap=tk.WORD, font=("Arial", 10))
        text_display.insert(tk.END, text_preview if text_preview else "无文本")
        text_display.config(state=tk.DISABLED, bd=0, padx=8, pady=8)
        text_display.pack(fill=tk.BOTH, expand=True)
        
        # 提示
        prompt_label = ttk.Label(content_frame, text="请选择下一步操作：", font=("Arial", 10))
        prompt_label.pack(anchor=tk.W, pady=(5, 0))
        
        # 分隔线
        separator = ttk.Separator(main_frame, orient="horizontal")
        separator.pack(fill=tk.X, padx=0, pady=0)
        
        # 按钮区域
        btn_frame = ttk.Frame(main_frame, padding=(15, 10))
        btn_frame.pack(fill=tk.X)
        
        def on_continue():
            dialog.destroy()
            # 继续选择，不做任何事
        
        def on_finish():
            dialog.destroy()
            self.filename_template_mode = False
            # 完成选择，开始扫描
            self.scan_filename_template()
        
        def on_cancel():
            dialog.destroy()
            # 取消选择，移除最后添加的区域
            if self.filename_template_coords:
                self.filename_template_coords.pop()
                self.template_region_count -= 1
                self.canvas.delete(f"filename_region_{self.template_region_count + 1}")
            
            # 如果还有其他选择的区域，保持选择模式
            if self.filename_template_coords:
                # 如果还有已选区域，弹出新对话框展示已有选择
                combined_text = [f"区域{i+1}: {text}" for i, text in enumerate(
                    [self.extract_text_from_selection(*coords) for coords in self.filename_template_coords]
                )]
                self.show_current_selections_dialog(combined_text)
            else:
                # 如果没有选择任何区域，则完全取消
                self.filename_template_mode = False
                self.canvas.delete("filename_region_*")
        
        style = ttk.Style()
        style.configure("Accent.TButton", font=("Arial", 10))
        
        cancel_btn = ttk.Button(btn_frame, text="取消", command=on_cancel)
        cancel_btn.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(btn_frame, text="继续选择", style="Accent.TButton", 
                 command=on_continue).pack(side=tk.RIGHT, padx=5)
        
        finish_btn = ttk.Button(btn_frame, text="完成", style="Accent.TButton", 
                              command=on_finish)
        finish_btn.pack(side=tk.RIGHT, padx=5)
        
        # 设置默认按钮和键盘快捷键
        finish_btn.focus_set()
        dialog.bind("<Return>", lambda event: finish_btn.invoke())
        dialog.bind("<Escape>", lambda event: cancel_btn.invoke())
        
        # 居中显示对话框
        self.center_dialog(dialog, 440, 240)
        
        dialog.wait_window()  # 等待对话框关闭

    def show_current_selections_dialog(self, texts):
        """显示当前已选择的区域信息"""
        dialog = tk.Toplevel(self.root)
        dialog.title("当前已选择区域")
        dialog.geometry("440x320")
        dialog.transient(self.root)
        dialog.grab_set()  # 模态对话框
        
        # 创建主框架
        main_frame = ttk.Frame(dialog, padding=0)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建标题栏 - 蓝色
        title_frame = tk.Frame(main_frame, height=6, bg="#3498db")
        title_frame.pack(fill=tk.X)
        
        # 内容区域
        content_frame = ttk.Frame(main_frame, padding=(20, 15, 20, 10))
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        header_label = ttk.Label(content_frame, text="已选择的区域", font=("Arial", 12, "bold"))
        header_label.pack(anchor=tk.W, pady=(0, 10))
        
        # 创建一个框架来容纳文本列表
        list_frame = ttk.Frame(content_frame, relief="solid", borderwidth=1)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 使用Text组件显示文本，带有更好的样式
        text_display = tk.Text(list_frame, height=10, width=40, wrap=tk.WORD, font=("Arial", 10))
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(list_frame, command=text_display.yview)
        scrollbar.pack(fill=tk.Y, side=tk.RIGHT)
        text_display.config(yscrollcommand=scrollbar.set, bd=0, padx=8, pady=8)
        text_display.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        
        # 插入文本，每个区域使用不同的标签样式
        text_display.tag_configure("header", font=("Arial", 10, "bold"), foreground="#3498db")
        
        for i, line in enumerate(texts):
            area_num = i + 1
            text_display.insert(tk.END, f"区域 {area_num}: ", "header")
            
            # 提取文本内容（去掉"区域X: "前缀）
            content = line[line.find(":")+1:].strip()
            text_display.insert(tk.END, f"{content}\n")
        
        text_display.config(state=tk.DISABLED)  # 设为只读
        
        # 分隔线
        separator = ttk.Separator(main_frame, orient="horizontal")
        separator.pack(fill=tk.X, padx=0, pady=0)
        
        # 按钮区域
        btn_frame = ttk.Frame(main_frame, padding=(15, 10))
        btn_frame.pack(fill=tk.X)
        
        def on_continue():
            dialog.destroy()
            # 继续选择区域
            
        def on_complete():
            dialog.destroy()
            self.filename_template_mode = False
            # 完成选择，开始扫描
            self.scan_filename_template()
        
        style = ttk.Style()
        style.configure("Accent.TButton", font=("Arial", 10))
            
        continue_btn = ttk.Button(btn_frame, text="继续选择", command=on_continue)
        continue_btn.pack(side=tk.LEFT, padx=5)
        
        complete_btn = ttk.Button(btn_frame, text="完成", style="Accent.TButton", command=on_complete)
        complete_btn.pack(side=tk.RIGHT, padx=5)
        
        # 设置默认按钮和键盘快捷键
        complete_btn.focus_set()
        dialog.bind("<Return>", lambda event: complete_btn.invoke())
        dialog.bind("<Escape>", lambda event: continue_btn.invoke())
        
        # 居中显示对话框
        self.center_dialog(dialog, 440, 320)
        
        dialog.wait_window()  # 等待对话框关闭

    def scan_filename_template(self):
        """扫描所有文件的文件名模板区域"""
        if not self.filename_template_coords:
            return
        
        try:
            # 显示进度窗口
            progress_window = tk.Toplevel(self.root)
            progress_window.title("识别进度")
            progress_window.geometry("300x150")
            progress_window.transient(self.root)  # 设置为主窗口的子窗口
            self.center_dialog(progress_window)  # 居中显示
            
            progress_label = ttk.Label(progress_window, text="正在识别文件名...")
            progress_label.pack(pady=10)
            
            progress_var = tk.DoubleVar()
            progress_bar = ttk.Progressbar(progress_window, variable=progress_var, maximum=100)
            progress_bar.pack(fill=tk.X, padx=20, pady=10)
            
            # 计算需要处理的页面
            pages_to_process = []  # [(file_path, page_num), ...]
            for file_path in self.pdf_files:
                if self.current_size_key:
                    # 如果是为特定图幅设置模板，只处理该图幅的页面
                    if self.current_size_key in self.page_sizes[file_path]:
                        page_nums = self.page_sizes[file_path][self.current_size_key]
                        pages_to_process.extend([(file_path, page_num) for page_num in page_nums])
                else:
                    # 如果不是特定图幅，处理所有页面
                    pdf = fitz.open(file_path)
                    pages_to_process.extend([(file_path, i) for i in range(len(pdf))])
                    pdf.close()
            
            total_pages = len(pages_to_process)
            processed_pages = 0
            total_pages_recognized = 0
            
            # 处理每个页面
            current_pdf = None
            current_file = None
            
            for file_path, page_num in pages_to_process:
                # 如果是新文件，关闭旧文件并打开新文件
                if current_file != file_path:
                    if current_pdf:
                        current_pdf.close()
                    current_pdf = fitz.open(file_path)
                    current_file = file_path
                    
                    progress_label.config(text=f"正在处理: {os.path.basename(file_path)}")
                    progress_window.update()
                    
                    # 确保文件的区域列表已初始化
                    if file_path not in self.selected_regions:
                        self.selected_regions[file_path] = []
                    if file_path not in self.custom_filenames:
                        self.custom_filenames[file_path] = {}
                
                # 获取页面
                page = current_pdf[page_num]
                
                # 计算选择区域在PDF坐标系中的位置
                zoom = 2.0 * self.scale_factor
                
                # 保持原始选择顺序提取文本
                combined_text = []
                
                # 按照用户框选的原始顺序提取文本
                for coords in self.filename_template_coords:
                    x1, y1, x2, y2 = coords
                    pdf_rect = fitz.Rect(x1/zoom, y1/zoom, x2/zoom, y2/zoom)
                    
                    # 提取文本并清理
                    text = self.extract_text_from_region(page, pdf_rect)
                    if text:
                        combined_text.append(text)
                
                # 使用所有框选区域的坐标
                all_coords = self.filename_template_coords
                
                # 组合所有区域的文本，用"-"连接
                if combined_text:
                    final_text = "-".join(combined_text)
                    # 清理非法文件名字符
                    final_text = re.sub(r'[\\/*?:"<>|]', '_', final_text)
                    self.custom_filenames[file_path][page_num] = final_text
                    total_pages_recognized += 1
                else:
                    final_text = ""
                
                # 查找这个页面是否已有区域
                existing_region = None
                existing_index = -1
                for i, region in enumerate(self.selected_regions[file_path]):
                    if region['page'] == page_num:
                        existing_region = region
                        existing_index = i
                        break
                
                if existing_region:
                    # 更新现有区域
                    self.selected_regions[file_path][existing_index].update({
                        'all_coords': all_coords,
                        'is_filename': True,
                        'filename': final_text if final_text else "[无内容]",
                        'filename_text': f"文件名: {final_text if final_text else '[无内容]'}"
                    })
                    if 'page_text' in existing_region:
                        self.selected_regions[file_path][existing_index]['text'] = \
                            f"{existing_region['page_text']} | {final_text if final_text else '[无内容]'}"
                    else:
                        self.selected_regions[file_path][existing_index]['text'] = \
                            f"文件名: {final_text if final_text else '[无内容]'}"
                else:
                    # 创建新区域
                    filename_region = {
                        'page': page_num,
                        'rect': all_coords[0],
                        'all_coords': all_coords,
                        'text': f"文件名: {final_text if final_text else '[无内容]'}",
                        'is_filename': True,
                        'filename': final_text if final_text else "[无内容]",
                        'filename_text': f"文件名: {final_text if final_text else '[无内容]'}"
                    }
                    self.selected_regions[file_path].append(filename_region)
                
                # 如果当前正在为特定图幅尺寸设置模板，保存模板信息
                if self.current_size_key:
                    if self.current_size_key not in self.filename_templates:
                        self.filename_templates[self.current_size_key] = []
                    # 只保存一次模板坐标
                    if not self.filename_templates[self.current_size_key]:
                        self.filename_templates[self.current_size_key] = all_coords
                
                processed_pages += 1
                # 更新进度
                progress = (processed_pages * 100.0) / total_pages
                progress_var.set(progress)
                progress_window.update()
            
            # 关闭最后打开的PDF文件
            if current_pdf:
                current_pdf.close()
            
            progress_window.destroy()
            
            # 更新区域列表显示
            self.update_region_list()
            
            # 显示识别结果
            size_text = f" (图幅: {self.current_size_key})" if self.current_size_key else ""
            self.showinfo('完成', f'文件名识别完成{size_text}！\n共识别出 {total_pages_recognized}/{total_pages} 页的文件名。')
            
        except Exception as e:
            self.showerror('错误', f'识别文件名时出错：{str(e)}')
            if 'progress_window' in locals() and progress_window.winfo_exists():
                progress_window.destroy()

    def extract_text_from_region(self, page, rect):
        """Extracts text from a specified region on a page using the helper method."""
        return self._extract_text_from_rect(page, rect)

    def center_window(self, window, width, height):
        """将窗口居中显示在屏幕上"""
        # 获取屏幕宽度和高度
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        
        # 计算窗口左上角的x, y坐标
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        
        # 确保窗口在屏幕内
        x = max(0, x)
        y = max(0, y)
        
        # 设置窗口位置
        window.geometry(f"{width}x{height}+{x}+{y}")
        window.update()  # 立即应用更改

    def center_dialog(self, dialog, width=None, height=None):
        """将对话框居中放置在主窗口上"""
        if not width or not height:
            # 如果未指定尺寸，获取对话框当前尺寸
            dialog.update_idletasks()  # 确保对话框已经绘制完成
            width = dialog.winfo_width()
            height = dialog.winfo_height()
            
            # 如果尺寸为0（可能是对话框还未完全初始化），使用默认尺寸
            if width <= 1 or height <= 1:
                width = 400
                height = 200
        
        # 获取主窗口的位置和尺寸
        root_x = self.root.winfo_rootx()
        root_y = self.root.winfo_rooty()
        root_width = self.root.winfo_width()
        root_height = self.root.winfo_height()
        
        # 计算对话框应该在的位置（居中）
        x = root_x + (root_width - width) // 2
        y = root_y + (root_height - height) // 2
        
        # 确保对话框在屏幕内
        screen_width = dialog.winfo_screenwidth()
        screen_height = dialog.winfo_screenheight()
        x = max(0, min(x, screen_width - width))
        y = max(0, min(y, screen_height - height))
        
        # 设置对话框位置
        dialog.geometry(f"{width}x{height}+{x}+{y}")
        dialog.lift()  # 确保对话框在顶层
        dialog.focus_set()  # 设置焦点
        dialog.update()  # 立即应用更改

    def show_custom_messagebox(self, title, message, message_type="info", buttons=None):
        """自定义消息框，确保居中显示
        
        参数:
            title: 对话框标题
            message: 显示的消息
            message_type: 消息类型 ("info", "warning", "error", "question")
            buttons: 按钮列表，如 [("确定", lambda: dialog.destroy()), ("取消", lambda: dialog.destroy())]
        
        返回:
            选择的按钮索引或None
        """
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.transient(self.root)
        dialog.grab_set()  # 模态对话框
        
        # 设置初始大小
        dialog.geometry("400x220")
        
        # 配置窗口属性
        dialog.resizable(False, False)
        
        # 创建主框架，使用浅灰色背景
        main_frame = ttk.Frame(dialog)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建标题框架 - 顶部带有颜色条
        title_frame = tk.Frame(main_frame, height=6)
        if message_type == "info":
            title_frame.configure(bg="#3498db")  # 蓝色
        elif message_type == "warning":
            title_frame.configure(bg="#f39c12")  # 橙色
        elif message_type == "error":
            title_frame.configure(bg="#e74c3c")  # 红色
        elif message_type == "question":
            title_frame.configure(bg="#2ecc71")  # 绿色
        title_frame.pack(fill=tk.X)
        
        # 内容框架
        content_frame = ttk.Frame(main_frame, padding=(20, 15, 20, 10))
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建图标和消息的水平布局
        icon_message_frame = ttk.Frame(content_frame)
        icon_message_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 创建图标
        icon_frame = ttk.Frame(icon_message_frame, width=60)
        icon_frame.pack(side=tk.LEFT, padx=(0, 15))
        
        icon_label = ttk.Label(icon_frame)
        if message_type == "info":
            icon_text = "ℹ"
            icon_color = "#3498db"  # 蓝色
        elif message_type == "warning":
            icon_text = "⚠"
            icon_color = "#f39c12"  # 橙色
        elif message_type == "error":
            icon_text = "✖"
            icon_color = "#e74c3c"  # 红色
        elif message_type == "question":
            icon_text = "?"
            icon_color = "#2ecc71"  # 绿色
        
        icon_label.configure(text=icon_text, foreground=icon_color, 
                            font=("Arial", 32, "bold"))
        icon_label.pack(side=tk.LEFT)
        
        # 创建消息文本框架
        msg_frame = ttk.Frame(icon_message_frame)
        msg_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 消息文本 - 使用更好的字体和行间距
        msg_label = ttk.Label(
            msg_frame, 
            text=message, 
            wraplength=300, 
            justify="left",
            font=("Arial", 11)
        )
        msg_label.pack(anchor=tk.W, fill=tk.BOTH, expand=True)
        
        # 分隔线
        separator = ttk.Separator(main_frame, orient="horizontal")
        separator.pack(fill=tk.X, padx=0, pady=0)
        
        # 创建按钮框架 - 底部浅灰色背景
        button_frame = ttk.Frame(main_frame, padding=(15, 10))
        button_frame.pack(fill=tk.X)
        
        result = None
        
        if not buttons:
            # 默认按钮
            if message_type == "question":
                buttons = [("是", 1), ("否", 0)]
            else:
                buttons = [("确定", 0)]
        
        # 确定按钮样式
        style = ttk.Style()
        style.configure("Accent.TButton", font=("Arial", 10))
        
        # 创建按钮 - 右对齐
        for i, (text, value) in enumerate(reversed(buttons)):
            btn = ttk.Button(
                button_frame, 
                text=text,
                style="Accent.TButton",
                command=lambda v=value: [dialog.destroy(), setattr(dialog, 'result', v)]
            )
            # 主要按钮在最右侧
            btn.pack(side=tk.RIGHT, padx=5)
        
        # 确保第一个按钮获得焦点
        if buttons:
            first_button = button_frame.winfo_children()[-1]
            first_button.focus_set()
            
            # 设置回车键绑定到第一个按钮
            dialog.bind("<Return>", lambda event: first_button.invoke())
            # 设置Escape键绑定到取消（如果有）
            if len(buttons) > 1:
                dialog.bind("<Escape>", lambda event: button_frame.winfo_children()[0].invoke())
        
        # 居中显示对话框
        self.center_dialog(dialog)
        
        # 等待对话框关闭
        dialog.wait_window()
        
        # 返回结果
        return getattr(dialog, 'result', None)
    
    def redraw_filename_regions(self):
        """重新绘制文件名模板区域"""
        if not self.filename_template_mode or not self.filename_template_coords:
            return
            
        for i, coords in enumerate(self.filename_template_coords):
            x1, y1, x2, y2 = coords
            self.canvas.create_rectangle(x1, y1, x2, y2, outline="orange", width=2, 
                                      tags=f"filename_region_{i+1}")

    def showinfo(self, title, message):
        """显示信息消息框"""
        return self.show_custom_messagebox(title, message, "info")
        
    def showwarning(self, title, message):
        """显示警告消息框"""
        return self.show_custom_messagebox(title, message, "warning")
        
    def showerror(self, title, message):
        """显示错误消息框"""
        return self.show_custom_messagebox(title, message, "error")
        
    def askyesno(self, title, message):
        """显示是/否消息框"""
        result = self.show_custom_messagebox(title, message, "question", 
                                          [("是", True), ("否", False)])
        return result

    def detect_page_sizes(self):
        """识别所有PDF文件中的不同图幅尺寸"""
        if not self.pdf_files:
            self.showwarning("警告", "请先添加PDF文件")
            return
            
        try:
            # 显示进度窗口
            progress_window = tk.Toplevel(self.root)
            progress_window.title("识别图幅尺寸")
            progress_window.geometry("300x150")
            progress_window.transient(self.root)
            self.center_dialog(progress_window)
            
            progress_label = ttk.Label(progress_window, text="正在识别图幅尺寸...")
            progress_label.pack(pady=10)
            
            progress_var = tk.DoubleVar()
            progress_bar = ttk.Progressbar(progress_window, variable=progress_var, maximum=100)
            progress_bar.pack(fill=tk.X, padx=20, pady=10)
            
            # 清空之前的尺寸信息
            self.page_sizes = {}
            
            total_files = len(self.pdf_files)
            processed_files = 0
            
            all_sizes = set()  # 收集所有不同的尺寸
            
            for file_path in self.pdf_files:
                progress_label.config(text=f"正在处理: {os.path.basename(file_path)}")
                progress_window.update()
                
                pdf_doc = fitz.open(file_path)
                self.page_sizes[file_path] = {}
                
                for page_num in range(len(pdf_doc)):
                    page = pdf_doc[page_num]
                    # 获取页面尺寸（四舍五入到整数）
                    width = round(page.rect.width)
                    height = round(page.rect.height)
                    size_key = f"{width}x{height}"
                    
                    if size_key not in self.page_sizes[file_path]:
                        self.page_sizes[file_path][size_key] = []
                    self.page_sizes[file_path][size_key].append(page_num)
                    all_sizes.add(size_key)
                
                pdf_doc.close()
                processed_files += 1
                progress_var.set(processed_files * 100 / total_files)
            
            progress_window.destroy()
            
            # 显示识别结果对话框
            self.show_size_detection_results(all_sizes)
            
        except Exception as e:
            self.showerror("错误", f"识别图幅尺寸时出错：{str(e)}")
            if 'progress_window' in locals() and progress_window.winfo_exists():
                progress_window.destroy()

    def show_size_detection_results(self, all_sizes):
        """显示图幅尺寸识别结果"""
        dialog = tk.Toplevel(self.root)
        dialog.title("图幅尺寸识别结果")
        dialog.geometry("500x400")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 创建主框架
        main_frame = ttk.Frame(dialog, padding=0)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建标题栏
        title_frame = tk.Frame(main_frame, height=6, bg="#3498db")
        title_frame.pack(fill=tk.X)
        
        # 内容区域
        content_frame = ttk.Frame(main_frame, padding=(20, 15, 20, 10))
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        header_label = ttk.Label(content_frame, text="检测到以下图幅尺寸：", 
                               font=("Arial", 12, "bold"))
        header_label.pack(anchor=tk.W, pady=(0, 10))
        
        # 创建表格框架
        table_frame = ttk.Frame(content_frame)
        table_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建表格
        columns = ("尺寸", "页数", "页码提取区域", "文件名提取区域")
        tree = ttk.Treeview(table_frame, columns=columns, show="headings")
        
        # 设置列标题
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=100)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 填充数据
        for size in sorted(all_sizes):
            # 计算使用该尺寸的页面总数
            total_pages = sum(len(sizes[size]) for sizes in self.page_sizes.values() if size in sizes)
            page_status = "未设置" if size not in self.page_templates else "已设置"
            filename_status = "未设置" if size not in self.filename_templates else "已设置"
            tree.insert("", tk.END, values=(size, total_pages, page_status, filename_status))
        
        # 分隔线
        separator = ttk.Separator(main_frame, orient="horizontal")
        separator.pack(fill=tk.X, padx=0, pady=0)
        
        # 按钮区域
        btn_frame = ttk.Frame(main_frame, padding=(15, 10))
        btn_frame.pack(fill=tk.X)
        
        def on_set_page_template():
            selection = tree.selection()
            if not selection:
                self.showwarning("警告", "请先选择一个图幅尺寸")
                return
                
            size_key = tree.item(selection[0])["values"][0]
            self.current_size_key = size_key
            dialog.destroy()
            # 开始为选中的尺寸设置页码提取区域
            self.start_template_selection_for_size(size_key, template_type="page")
            
        def on_set_filename_template():
            selection = tree.selection()
            if not selection:
                self.showwarning("警告", "请先选择一个图幅尺寸")
                return
                
            size_key = tree.item(selection[0])["values"][0]
            self.current_size_key = size_key
            dialog.destroy()
            # 开始为选中的尺寸设置文件名提取区域
            self.start_template_selection_for_size(size_key, template_type="filename")
        
        # 设置页码提取区域按钮
        set_page_btn = ttk.Button(btn_frame, text="设置页码提取区域", 
                                command=on_set_page_template)
        set_page_btn.pack(side=tk.RIGHT, padx=5)
        
        # 设置文件名提取区域按钮
        set_filename_btn = ttk.Button(btn_frame, text="设置文件名提取区域", 
                                    command=on_set_filename_template)
        set_filename_btn.pack(side=tk.RIGHT, padx=5)
        
        # 关闭按钮
        close_btn = ttk.Button(btn_frame, text="关闭", command=dialog.destroy)
        close_btn.pack(side=tk.RIGHT, padx=5)
        
        # 居中显示对话框
        self.center_dialog(dialog)
        
        dialog.wait_window()

    def start_template_selection_for_size(self, size_key, template_type="page"):
        """为指定尺寸开始模板选择过程"""
        # 找到第一个具有该尺寸的页面并显示
        for file_path in self.pdf_files:
            if size_key in self.page_sizes[file_path]:
                page_num = self.page_sizes[file_path][size_key][0]
                # 加载对应的文件和页面
                self.load_pdf_file(self.pdf_files.index(file_path))
                self.current_page = page_num
                self.update_page_display()
                break
        
        # 根据模板类型开始相应的选择过程
        if template_type == "page":
            self.start_template_selection(mode="double")  # 或根据需要选择mode="single"
            self.template_status.config(text=f"正在为图幅 {size_key} 设置页码提取区域...")
        else:  # filename
            self.start_filename_template_selection()
            self.template_status.config(text=f"正在为图幅 {size_key} 设置文件名提取区域...")

def main():
    root = tk.Tk()
    app = PDFSplitterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
