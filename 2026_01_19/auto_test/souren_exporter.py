from lib.var import *

class ResultExporter:
    """结果导出器 - 将JSON结果自动转换为Excel,使用matplotlib绘制折线图"""
    
    def __init__(self):
        # 延迟导入，避免循环导入问题
        import souren_config
        
        # 获取配置
        self.log_dir = souren_config.EXECUTION_DIR
        
        # 应用配置值
        self.default_row_height = souren_config.EXCEL_DEFAULT_ROW_HEIGHT
        self.default_column_width = souren_config.EXCEL_DEFAULT_COLUMN_WIDTH
        self.header_row_height = souren_config.EXCEL_HEADER_ROW_HEIGHT
        
        # 应用列宽配置
        self.summary_column_widths = souren_config.EXCEL_SUMMARY_COLUMN_WIDTHS
        self.details_column_widths = souren_config.EXCEL_DETAILS_COLUMN_WIDTHS
        self.chart_column_widths = souren_config.EXCEL_CHART_COLUMN_WIDTHS
        
        # 应用颜色配置
        self.border_colors = souren_config.EXCEL_BORDER_COLORS
        self.generate_chart_colors_func = souren_config.generate_chart_colors
        
        # 图表步骤筛选
        self.chart_steps_filter = souren_config.CHART_STEPS_FILTER
        
        # 创建样式对象
        self.header_fill = PatternFill(
            start_color=souren_config.EXCEL_HEADER_FILL_COLOR, 
            end_color=souren_config.EXCEL_HEADER_FILL_COLOR, 
            fill_type="solid"
        )
        self.header_font = Font(
            color=souren_config.EXCEL_HEADER_FONT_COLOR, 
            bold=True, 
            size=souren_config.EXCEL_HEADER_FONT_SIZE
        )
        self.query_result_fill = PatternFill(
            start_color=souren_config.EXCEL_QUERY_RESULT_FILL, 
            end_color=souren_config.EXCEL_QUERY_RESULT_FILL, 
            fill_type="solid"
        )
        self.query_result_font = Font(
            color=souren_config.EXCEL_QUERY_RESULT_FONT, 
            bold=True
        )
        self.data_font = Font(size=souren_config.EXCEL_DATA_FONT_SIZE)
        
        # 对齐样式
        self.center_align = Alignment(horizontal="center", vertical="center")
        self.left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
        
        # 缓存动态生成的颜色
        self._chart_colors_cache = {}
    
    def _get_chart_colors(self, count: int) -> List[str]:
        """获取图表颜色，动态生成"""
        if count not in self._chart_colors_cache:
            self._chart_colors_cache[count] = self.generate_chart_colors_func(count)
        return self._chart_colors_cache[count]
    
    def _hex_to_matplotlib_color(self, hex_color: str) -> str:
        """将十六进制颜色转换为matplotlib可识别的格式"""
        if not hex_color.startswith('#'):
            return f"#{hex_color}"
        return hex_color
    
    def find_latest_json_result(self) -> Optional[str]:
        """查找最新的JSON结果文件"""
        try:
            if not os.path.exists(self.log_dir):
                return None
            
            json_files = []
            for filename in os.listdir(self.log_dir):
                if filename.startswith("souren_results_") and filename.endswith(".json"):
                    filepath = os.path.join(self.log_dir, filename)
                    if os.path.isfile(filepath):
                        mtime = os.path.getmtime(filepath)
                        json_files.append({"path": filepath, "mtime": mtime})
            
            if not json_files:
                return None
            
            json_files.sort(key=lambda x: x["mtime"], reverse=True)
            return json_files[0]["path"]
            
        except Exception:
            return None
    
    def load_json_results(self, json_file: str) -> List[Dict]:
        """加载JSON结果数据"""
        try:
            with open(json_file, 'r', encoding='utf-8') as f:
                content = f.read()
            
            if not content.strip():
                return []
            
            results = json.loads(content)
            return [results] if isinstance(results, dict) else results
            
        except Exception:
            return []
    
    def convert_to_excel(self, json_file: str = None) -> bool:
        """将JSON结果转换为Excel文件"""
        try:
            if json_file is None:
                json_file = self.find_latest_json_result()
                if json_file is None:
                    return False
            
            json_name = os.path.basename(json_file)
            excel_name = json_name.replace('.json', '.xlsx')
            excel_file = os.path.join(os.path.dirname(json_file), excel_name)
            
            # 检查文件是否被锁定，如果是则创建新文件名
            import time
            if os.path.exists(excel_file):
                try:
                    # 尝试重命名文件
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
                    backup_file = excel_file.replace('.xlsx', f'_{timestamp}.xlsx')
                    os.rename(excel_file, backup_file)
                    print(f"⚠️ 原文件已存在，已重命名为: {os.path.basename(backup_file)}")
                except:
                    # 如果重命名失败，使用新文件名
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    excel_name = f"souren_results_{timestamp}.xlsx"
                    excel_file = os.path.join(os.path.dirname(json_file), excel_name)
                    print(f"⚠️ 文件被锁定，使用新文件名: {excel_name}")
            
            results = self.load_json_results(json_file)
            if not results:
                return False
            
            # 创建Excel文件
            with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
                self._create_summary_sheet(writer, results)
                self._create_execution_details_sheet(writer, results)
            
            # 应用样式和图表
            self._apply_styles_and_charts(excel_file, results)
            
            print(f"✅ Excel文件已生成: {excel_file}")
            return True
            
        except Exception as e:
            print(f"❌ Excel文件导出失败: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def _create_summary_sheet(self, writer: pd.ExcelWriter, results: List[Dict]):
        """创建汇总表"""
        try:
            valid_results = [r for r in results if r.get('executed_steps', 0) > 0]
            
            if not valid_results:
                df_empty = pd.DataFrame({"状态": ["没有有效的执行记录"]})
                df_empty.to_excel(writer, sheet_name='执行汇总', index=False)
                return
            
            summary_data = []
            for i, result in enumerate(valid_results):
                executed = result.get('executed_steps', 0)
                passed = result.get('passed', 0)
                success_rate = round((passed / executed) * 100, 2) if executed > 0 else 0
                
                summary = {
                    "序号": i + 1,
                    "执行时间": result.get('timestamp_readable', result.get('timestamp', 'N/A')),
                    "SCV文件": result.get('file', 'N/A'),
                    "设备": self._extract_device_name(result),
                    "执行模式": result.get('mode', 'run_all'),
                    "循环次数": result.get('loop_count', 1),
                    "总步骤数": result.get('total_steps', result.get('total_executions', 0)),
                    "已执行步骤": executed,
                    "通过步骤": passed,
                    "失败步骤": result.get('failed', 0),
                    "成功率(%)": success_rate,
                    "总耗时(秒)": round(float(result.get('execution_time', result.get('duration', 0))), 2),
                    "状态": result.get('status', 'N/A'),
                    "状态消息": result.get('message', 'N/A')
                }
                summary_data.append(summary)
            
            if summary_data:
                df_summary = pd.DataFrame(summary_data)
                df_summary.to_excel(writer, sheet_name='执行汇总', index=False)
            
        except Exception as e:
            print(f"创建汇总表失败: {e}")
    
    def _extract_device_name(self, result: Dict) -> str:
        """提取设备名称"""
        device = result.get('device', {})
        if isinstance(device, dict):
            return device.get('name', device.get('address', 'N/A'))
        elif isinstance(device, str):
            return device
        else:
            return 'N/A'
    
    def _create_execution_details_sheet(self, writer: pd.ExcelWriter, results: List[Dict]):
        """创建详细执行记录表"""
        try:
            details_data = []
            
            for result_index, result in enumerate(results):
                if result.get('executed_steps', 0) == 0:
                    continue
                    
                execution_details = result.get('execution_details', [])
                loop_count = result.get('loop_count', 1)
                
                if result.get('mode', 'run_all') != 'loop_info' or loop_count <= 1:
                    for detail in execution_details:
                        details_data.append(self._create_detail_record(detail, result_index, result, 1, 1))
                    continue
                
                # 按循环轮次分组
                loop_groups = {}
                for detail in execution_details:
                    loop_index = detail.get('loop_index', 1)
                    if loop_index not in loop_groups:
                        loop_groups[loop_index] = []
                    loop_groups[loop_index].append(detail)
                
                for loop_index in sorted(loop_groups.keys()):
                    if loop_index > 1:
                        details_data.extend([{}, {}])
                    for detail in loop_groups[loop_index]:
                        details_data.append(self._create_detail_record(
                            detail, result_index, result, loop_index, loop_count
                        ))
            
            if details_data:
                columns = [
                    "汇总序号", "执行时间", "步骤序号", "循环轮次", "总循环数", 
                    "步骤内容", "步骤类型", "执行状态", "执行结果", "耗时(秒)"
                ]
                
                formatted_data = []
                for record in details_data:
                    formatted_record = {}
                    for col in columns:
                        formatted_record[col] = record.get(col, "")
                    formatted_data.append(formatted_record)
                
                df_details = pd.DataFrame(formatted_data, columns=columns)
                df_details.to_excel(writer, sheet_name='详细执行记录', index=False)
            
        except Exception as e:
            print(f"创建详细记录表失败: {e}")
    
    def _create_detail_record(self, detail: Dict, result_index: int, result: Dict, 
                            loop_index: int, total_loops: int) -> Dict:
        """创建单个详细记录"""
        return {
            "汇总序号": result_index + 1,
            "执行时间": result.get('timestamp_readable', result.get('timestamp', 'N/A')),
            "步骤序号": detail.get('original_step', detail.get('step', 'N/A')),
            "循环轮次": loop_index,
            "总循环数": total_loops,
            "步骤内容": detail.get('content', ''),
            "步骤类型": detail.get('type', 'Normal'),
            "执行状态": detail.get('status', 'N/A'),
            "执行结果": str(detail.get('result', 'N/A')),
            "耗时(秒)": round(float(detail.get('duration', 0)), 3)
        }
    
    def _extract_chart_data(self, results: List[Dict]) -> Dict:
        """提取图表数据，按步骤和循环分组，根据配置筛选步骤"""
        chart_data = {}
        
        for result in results:
            execution_details = result.get('execution_details', [])
            
            for detail in execution_details:
                step_content = detail.get('content', '')
                result_str = str(detail.get('result', ''))
                step_num = detail.get('original_step', detail.get('step', 0))
                loop_index = detail.get('loop_index', 1)
                
                # 根据配置筛选步骤
                if step_num not in self.chart_steps_filter:
                    continue
                
                # 只处理查询命令且有结果的数据
                if '?' in step_content and result_str and result_str != 'N/A':
                    numbers = self._extract_all_numbers(result_str)
                    
                    if numbers:
                        step_key = f"步骤{step_num}"
                        loop_key = f"循环{loop_index}"
                        
                        if step_key not in chart_data:
                            chart_data[step_key] = {
                                "step_num": int(step_num),
                                "content": step_content.replace('?', '').strip(),
                                "loops": {}
                            }
                        
                        # 存储该循环的数据点
                        if loop_key not in chart_data[step_key]["loops"]:
                            chart_data[step_key]["loops"][loop_key] = {
                                "loop_index": loop_index,
                                "data_points": numbers
                            }
        
        return chart_data
    
    def _extract_all_numbers(self, text: str) -> List[float]:
        """从文本中提取所有数字，保持原值"""
        try:
            if not text:
                return []
            
            clean_text = text.strip()
            error_keywords = ["仪器通信错误", "VI_ERROR_TMO", "Timeout", "通信失败"]
            if any(keyword in clean_text for keyword in error_keywords):
                return []
            
            pattern = r'[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?'
            matches = re.findall(pattern, clean_text)
            
            numbers = []
            for match in matches:
                try:
                    number = float(match)
                    numbers.append(number)
                except ValueError:
                    continue
            
            return numbers
        except Exception:
            return []
    
    def _create_matplotlib_line_chart(self, step_num: int, step_content: str, 
                                     data_by_loop: Dict[int, List[float]]) -> BytesIO:
        """使用matplotlib创建折线图，返回图像字节流"""
        try:
            # 设置中文字体支持（如果需要）
            plt.rcParams['font.sans-serif'] = ['SimHei', 'Arial Unicode MS', 'DejaVu Sans']
            plt.rcParams['axes.unicode_minus'] = False
            
            # 创建图形
            fig, ax = plt.subplots(figsize=(10, 6))
            
            # 动态生成颜色并转换为matplotlib格式
            colors = self._get_chart_colors(len(data_by_loop))
            matplotlib_colors = [self._hex_to_matplotlib_color(color) for color in colors]
            
            # 线条样式和标记
            line_styles = ['o-', 's-', '^-', 'v-', 'd-', 'p-', 'h-', '*-']
            markers = ['o', 's', '^', 'v', 'D', 'p', 'h', '*']
            
            # 绘制每个循环的数据
            for idx, (loop_num, data_points) in enumerate(sorted(data_by_loop.items())):
                if not data_points or len(data_points) < 2:
                    print(f"    ⚠️ 循环{loop_num}数据点不足 ({len(data_points)}个)，跳过")
                    continue
                
                # 生成x轴数据点编号
                x_data = list(range(1, len(data_points) + 1))
                
                # 选择颜色和样式
                color = matplotlib_colors[idx % len(matplotlib_colors)]
                line_style = line_styles[idx % len(line_styles)]
                marker = markers[idx % len(markers)]
                
                # 创建标签（显示前2个值）
                if len(data_points) >= 2:
                    label = f'循环{loop_num}: {data_points[0]:.3f} → {data_points[1]:.3f}'
                else:
                    label = f'循环{loop_num}'
                
                # 绘制折线
                ax.plot(x_data, data_points, line_style, color=color, 
                       label=label, markersize=8, linewidth=2, alpha=0.8)
                
                # 添加数据点标签
                for x, y in zip(x_data, data_points):
                    ax.text(x, y, f'{y:.3f}', ha='center', va='bottom', 
                           fontsize=8, color=color, alpha=0.7)
            
            # 检查是否成功绘制了数据
            if not ax.lines:
                plt.close(fig)
                print(f"    ⚠️ 没有成功绘制任何数据，跳过步骤{step_num}")
                return None
            
            # 设置图表属性
            ax.set_xlabel('数据点编号', fontsize=12)
            ax.set_ylabel('数值', fontsize=12)
            ax.set_title(f'步骤 {step_num} 折线图\n{step_content}', fontsize=14, pad=20)
            ax.legend(loc='best', fontsize=10)
            ax.grid(True, linestyle='--', alpha=0.5)
            
            # 自动调整布局
            plt.tight_layout()
            
            # 将图表保存到BytesIO
            img_data = BytesIO()
            plt.savefig(img_data, format='png', dpi=150, bbox_inches='tight')
            img_data.seek(0)
            plt.close(fig)  # 关闭图形释放内存
            
            print(f"    ✅ 步骤{step_num}折线图创建成功，包含 {len(ax.lines)} 条数据线")
            return img_data
            
        except Exception as e:
            print(f"❌ 创建matplotlib图表失败（步骤{step_num}）: {e}")
            import traceback
            traceback.print_exc()
            return None
    
    def _apply_styles_and_charts(self, excel_file: str, results: List[Dict]):
        """应用样式和图表到Excel文件，使用matplotlib绘制折线图"""
        try:
            wb = load_workbook(excel_file)
            
            # 1. 格式化执行汇总表
            if '执行汇总' in wb.sheetnames:
                self._format_summary_sheet(wb['执行汇总'])
            
            # 2. 格式化详细执行记录表
            if '详细执行记录' in wb.sheetnames:
                self._format_details_sheet(wb['详细执行记录'])
            
            # 3. 提取图表数据
            chart_data = self._extract_chart_data(results)
            
            if chart_data:
                # 4. 创建数据分析图表工作表（包含数据表格和matplotlib图表）
                self._create_chart_sheet_with_matplotlib(wb, chart_data)
            else:
                # 创建空的数据分析图表工作表
                if '数据分析图表' in wb.sheetnames:
                    wb.remove(wb['数据分析图表'])
                chart_ws = wb.create_sheet(title='数据分析图表')
                chart_ws['A1'] = "数据分析图表"
                chart_ws['A2'] = "⚠️ 没有找到配置步骤的查询数据"
                print("⚠️ 没有找到配置步骤的查询数据")
            
            # 5. 调整列宽行高（应用配置文件中的设置）
            self._adjust_dimensions(wb)
            
            wb.save(excel_file)
            
        except Exception as e:
            print(f"❌ 应用样式和图表失败: {e}")
            import traceback
            traceback.print_exc()
    
    def _format_summary_sheet(self, ws):
        """格式化汇总表 - 应用配置文件中的样式"""
        try:
            # 格式化表头 - 应用配置
            for cell in ws[1]:
                cell.fill = self.header_fill
                cell.font = self.header_font
                cell.alignment = self.center_align
                cell.border = self._create_border(self.border_colors[0])
            
            # 格式化数据行
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                for cell in row:
                    cell.border = self._create_border(self.border_colors[0])
                    cell.font = self.data_font
                    
                    # 设置对齐方式
                    col_idx = cell.column
                    if col_idx in [1, 5, 6, 7, 8, 9, 10, 11, 12, 13]:
                        cell.alignment = self.center_align
                    elif col_idx == 2:
                        cell.alignment = Alignment(horizontal="left", vertical="center")
                    else:
                        cell.alignment = self.left_align
                
                # 设置行高 - 应用配置
                ws.row_dimensions[row[0].row].height = self.default_row_height
            
        except Exception as e:
            print(f"❌ 格式化汇总表失败: {e}")
    
    def _format_details_sheet(self, ws):
        """格式化详细执行记录表 - 应用配置文件中的样式"""
        try:
            # 格式化表头 - 应用配置
            for cell in ws[1]:
                cell.fill = self.header_fill
                cell.font = self.header_font
                cell.alignment = self.center_align
                cell.border = self._create_border(self.border_colors[0])
            
            # 设置表头行高 - 应用配置
            ws.row_dimensions[1].height = self.header_row_height
            
            # 处理数据行
            for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row), start=2):
                if not any(cell.value for cell in row):
                    continue
                
                # 检查是否是查询命令
                step_content_cell = row[5]  # F列是步骤内容
                is_query = step_content_cell.value and '?' in str(step_content_cell.value)
                
                for cell in row:
                    cell.border = self._create_border(self.border_colors[0])
                    cell.font = self.data_font
                    
                    # 设置对齐方式
                    col_idx = cell.column
                    if col_idx in [1, 3, 4, 5, 7, 8, 10]:
                        cell.alignment = self.center_align
                    elif col_idx == 2:
                        cell.alignment = Alignment(horizontal="left", vertical="center")
                    else:
                        cell.alignment = self.left_align
                
                # 高亮显示查询结果行 - 应用配置
                if is_query:
                    for cell in row:
                        cell.fill = self.query_result_fill
                        cell.font = self.query_result_font
                
                # 设置数据行高 - 应用配置
                ws.row_dimensions[row_idx].height = self.default_row_height
            
        except Exception as e:
            print(f"❌ 格式化详细记录表失败: {e}")
    
    def _create_chart_sheet_with_matplotlib(self, wb, chart_data):
        """创建包含数据表格和matplotlib图表的综合分析工作表"""
        try:
            # 移除旧的图表工作表
            if '数据分析图表' in wb.sheetnames:
                wb.remove(wb['数据分析图表'])
            
            # 创建新的数据分析图表工作表
            chart_ws = wb.create_sheet(title='数据分析图表')
            
            # 1. 添加标题
            chart_ws['A1'] = "数据分析图表"
            chart_ws['A1'].font = Font(size=14, bold=True, color=self.header_fill.start_color.rgb)
            chart_ws.merge_cells('A1:L1')
            
            # 显示筛选的步骤信息
            chart_ws['A2'] = f"显示步骤: {', '.join(map(str, self.chart_steps_filter))}"
            chart_ws['A2'].font = Font(size=10, italic=True)
            
            # 2. 创建数据表格标题
            chart_ws['A4'] = "查询命令数据表格"
            chart_ws['A4'].font = Font(size=12, bold=True)
            
            # 3. 准备表格数据
            all_step_data = []
            
            # 按步骤编号排序
            sorted_steps = sorted(chart_data.items(), key=lambda x: x[1]["step_num"])
            
            print(f"📊 筛选出的步骤数据: {[step_info['step_num'] for step_key, step_info in sorted_steps]}")
            
            for step_key, step_info in sorted_steps:
                step_num = step_info["step_num"]
                step_content = step_info["content"]
                
                # 按循环排序
                sorted_loops = sorted(step_info["loops"].items(), 
                                     key=lambda x: int(x[1]["loop_index"]))
                
                for loop_key, loop_info in sorted_loops:
                    loop_index = loop_info["loop_index"]
                    data_points = loop_info["data_points"]
                    
                    # 创建数据行
                    row_data = [f"步骤{step_num}-循环{loop_index}({step_content})"]
                    
                    # 添加数据点，最多11个
                    for i in range(min(11, len(data_points))):
                        row_data.append(data_points[i])
                    
                    # 如果数据点不足11个，填充空值
                    for i in range(len(data_points), 11):
                        row_data.append("")
                    
                    all_step_data.append(row_data)
            
            if not all_step_data:
                chart_ws['A6'] = "没有找到查询命令的数字数据"
                return
            
            print(f"📊 找到 {len(all_step_data)} 条查询命令数据")
            
            # 4. 创建表格表头
            headers = ["数据系列", "数据点1", "数据点2", "数据点3", "数据点4", "数据点5", 
                      "数据点6", "数据点7", "数据点8", "数据点9", "数据点10", "数据点11"]
            
            start_row = 6
            for col_idx, header in enumerate(headers, 1):
                cell = chart_ws.cell(row=start_row, column=col_idx, value=header)
                cell.fill = self.header_fill
                cell.font = self.header_font
                cell.alignment = self.center_align
                cell.border = self._create_border(self.border_colors[0])
            
            # 设置表头行高
            chart_ws.row_dimensions[start_row].height = self.header_row_height
            
            # 5. 填充数据并设置颜色
            # 首先确定总循环数
            all_loops = set()
            for row_data in all_step_data:
                match = re.search(r'循环(\d+)', str(row_data[0]))
                if match:
                    all_loops.add(int(match.group(1)))
            
            total_loops = len(all_loops)
            
            # 动态生成颜色
            chart_colors = self._get_chart_colors(total_loops)
            print(f"🎨 为 {total_loops} 个循环生成了 {len(chart_colors)} 种颜色")
            
            for row_idx, row_data in enumerate(all_step_data, start=start_row + 1):
                # 提取循环编号
                loop_num = 1
                match = re.search(r'循环(\d+)', str(row_data[0]))
                if match:
                    loop_num = int(match.group(1))
                
                # 获取对应的颜色
                color_idx = (loop_num - 1) % len(chart_colors)
                loop_color = chart_colors[color_idx]
                
                for col_idx, cell_value in enumerate(row_data, 1):
                    cell = chart_ws.cell(row=row_idx, column=col_idx, value=cell_value)
                    
                    # 应用边框
                    border_color_idx = ((row_idx - start_row - 1) % len(self.border_colors))
                    cell.border = self._create_border(self.border_colors[border_color_idx])
                    
                    cell.alignment = self.center_align if col_idx > 1 else self.left_align
                    
                    # 设置字体颜色（包括数据点）
                    if cell_value not in (None, ""):
                        cell.font = Font(color=loop_color, size=10)
                        if col_idx == 1:  # 数据系列列加粗
                            cell.font = Font(color=loop_color, bold=True, size=10)
                
                # 设置数据行高
                chart_ws.row_dimensions[row_idx].height = self.default_row_height
            
            # 6. 使用matplotlib创建折线图
            chart_start_row = start_row + len(all_step_data) + 3
            
            # 按步骤分组数据
            step_data_by_loop = {}
            for step_key, step_info in sorted_steps:
                step_num = step_info["step_num"]
                step_content = step_info["content"]
                
                # 收集该步骤的所有循环数据
                loop_data_dict = {}
                for loop_key, loop_info in step_info["loops"].items():
                    loop_num = loop_info["loop_index"]
                    data_points = loop_info["data_points"]
                    if data_points and len(data_points) >= 2:
                        loop_data_dict[loop_num] = data_points
                    else:
                        print(f"  ⚠️ 步骤{step_num}-循环{loop_num}数据点不足 ({len(data_points) if data_points else 0}个)，跳过")
                
                if loop_data_dict:
                    step_data_by_loop[step_num] = {
                        "content": step_content,
                        "data_by_loop": loop_data_dict
                    }
            
            print(f"📈 将使用matplotlib创建 {len(step_data_by_loop)} 个步骤的折线图")
            
            chart_row_offset = 0
            charts_created = 0
            
            for step_num, step_info in sorted(step_data_by_loop.items()):
                step_content = step_info["content"]
                data_by_loop = step_info["data_by_loop"]
                
                print(f"  📊 处理步骤 {step_num}: {len(data_by_loop)} 个有效循环")
                
                # 创建matplotlib折线图
                img_data = self._create_matplotlib_line_chart(
                    step_num, step_content, data_by_loop
                )
                
                if img_data:
                    # 创建图像对象
                    img = Image(img_data)
                    
                    # 调整图像大小
                    img.width = 600  # 像素宽度
                    img.height = 400  # 像素高度
                    
                    # 将图像插入到Excel中
                    chart_cell = f"A{chart_start_row + chart_row_offset}"
                    chart_ws.add_image(img, chart_cell)
                    
                    # 在图表下方添加说明
                    explanation_row = chart_start_row + chart_row_offset + 20
                    chart_ws.cell(row=explanation_row, column=1, value=f"步骤{step_num}数据说明:")
                    chart_ws.cell(row=explanation_row + 1, column=1, value=f"• 图表类型: 折线图 (使用matplotlib生成)")
                    chart_ws.cell(row=explanation_row + 2, column=1, value=f"• 数据系列: {len(data_by_loop)} 个循环")
                    chart_ws.cell(row=explanation_row + 3, column=1, value=f"• 步骤内容: {step_content}")
                    
                    chart_row_offset += 25
                    charts_created += 1
                else:
                    print(f"  ⚠️ 步骤 {step_num} 的折线图创建失败，跳过")
            
            if charts_created == 0:
                chart_ws[f"A{chart_start_row}"] = "⚠️ 没有足够的数据来创建折线图"
                chart_ws[f"A{chart_start_row + 1}"] = "可能原因:"
                chart_ws[f"A{chart_start_row + 2}"] = "1. 数据点数量不足 (每个循环至少需要2个数据点)"
                chart_ws[f"A{chart_start_row + 3}"] = "2. 查询命令没有返回有效的数字数据"
                chart_ws[f"A{chart_start_row + 4}"] = "3. 数据提取失败"
                print("⚠️ 没有足够的数据来创建折线图")
            else:
                print(f"✅ 成功使用matplotlib创建了 {charts_created} 个折线图")
            
            # 7. 设置列宽 - 应用配置
            for col_letter, width in self.chart_column_widths.items():
                if col_letter == 'default':
                    # 设置默认列宽
                    for col in range(1, chart_ws.max_column + 1):
                        col_letter = get_column_letter(col)
                        if col_letter not in self.chart_column_widths:
                            chart_ws.column_dimensions[col_letter].width = width
                else:
                    if col_letter in chart_ws.column_dimensions:
                        chart_ws.column_dimensions[col_letter].width = width
            
        except Exception as e:
            print(f"❌ 创建数据分析图表失败: {e}")
            import traceback
            traceback.print_exc()
    
    def _create_border(self, color: str) -> Border:
        """创建边框 - 使用配置的颜色"""
        side = Side(style='thin', color=color)
        return Border(left=side, right=side, top=side, bottom=side)
    
    def _adjust_dimensions(self, wb):
        """调整列宽和行高 - 应用配置文件中的所有设置"""
        try:
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                
                # 设置行高
                for row in range(1, ws.max_row + 1):
                    if row == 1:
                        # 表头行高 - 应用配置
                        ws.row_dimensions[row].height = self.header_row_height
                    else:
                        # 数据行高 - 应用配置
                        ws.row_dimensions[row].height = self.default_row_height
                
                # 根据工作表类型应用不同的列宽设置
                if sheet_name == '执行汇总':
                    # 首先设置默认列宽
                    for col in ws.columns:
                        col_letter = get_column_letter(col[0].column)
                        ws.column_dimensions[col_letter].width = self.default_column_width
                    
                    # 然后应用特殊列宽配置
                    for col_letter, width in self.summary_column_widths.items():
                        if col_letter in ws.column_dimensions:
                            ws.column_dimensions[col_letter].width = width
                            
                elif sheet_name == '详细执行记录':
                    # 首先设置默认列宽
                    for col in ws.columns:
                        col_letter = get_column_letter(col[0].column)
                        ws.column_dimensions[col_letter].width = self.default_column_width
                    
                    # 然后应用特殊列宽配置
                    for col_letter, width in self.details_column_widths.items():
                        if col_letter in ws.column_dimensions:
                            ws.column_dimensions[col_letter].width = width
                            
                elif sheet_name == '数据分析图表':
                    # 应用图表列宽配置
                    for col_letter, width in self.chart_column_widths.items():
                        if col_letter == 'default':
                            # 设置默认列宽
                            for col in range(1, ws.max_column + 1):
                                col_letter = get_column_letter(col)
                                if col_letter not in self.chart_column_widths:
                                    ws.column_dimensions[col_letter].width = width
                        elif col_letter in ws.column_dimensions:
                            ws.column_dimensions[col_letter].width = width
                            
                else:
                    # 其他工作表设置默认列宽
                    for col in ws.columns:
                        col_letter = get_column_letter(col[0].column)
                        ws.column_dimensions[col_letter].width = self.default_column_width
                
                # 启用筛选（除了图表工作表）
                if sheet_name != '数据分析图表' and ws.max_row > 1:
                    try:
                        ws.auto_filter.ref = ws.dimensions
                    except:
                        pass
            
            print(f"✅ 已应用所有Excel导出配置")
            print(f"   - 图表显示步骤: {self.chart_steps_filter}")
            
        except Exception as e:
            print(f"❌ 调整列宽行高失败: {e}")

# 保留独立的导出功能
if __name__ == "__main__":
    exporter = ResultExporter()
    
    # 可以传入特定的json文件路径
    # x = r"D:\feifei\2026_01_05\auto_test\log\execution_20260113_163555\souren_results_20260113_163555.json"
    success = exporter.convert_to_excel()
    
    if success:
        print("✅ Excel文件导出成功!")
    else:
        print("❌ Excel文件导出失败")