from lib.var import *
import souren_config

class ResultExporter:

    def __init__(self):
        self.log_dir = souren_config.EXECUTION_DIR

        self.default_row_height = getattr(souren_config, 'EXCEL_DEFAULT_ROW_HEIGHT', 13.5)
        self.default_column_width = getattr(souren_config, 'EXCEL_DEFAULT_COLUMN_WIDTH', 9)
        self.header_row_height = getattr(souren_config, 'EXCEL_HEADER_ROW_HEIGHT', 20)

        self.summary_column_widths = getattr(souren_config, 'EXCEL_SUMMARY_COLUMN_WIDTHS', {})
        self.details_column_widths = getattr(souren_config, 'EXCEL_DETAILS_COLUMN_WIDTHS', {})
        self.chart_column_widths = getattr(souren_config, 'EXCEL_CHART_COLUMN_WIDTHS', {'default': 15})

        self.border_colors = getattr(souren_config, 'EXCEL_BORDER_COLORS', {"default": "000000"})

        self.header_fill = PatternFill(
            start_color=getattr(souren_config, 'EXCEL_HEADER_FILL_COLOR', "366092"),
            end_color=getattr(souren_config, 'EXCEL_HEADER_FILL_COLOR', "366092"),
            fill_type="solid"
        )
        self.header_font = Font(
            color=getattr(souren_config, 'EXCEL_HEADER_FONT_COLOR', "FFFFFF"),
            bold=True,
            size=getattr(souren_config, 'EXCEL_HEADER_FONT_SIZE', 11)
        )

        self.failed_font = Font(color="FF0000", bold=True)

        self.data_font = Font(size=getattr(souren_config, 'EXCEL_DATA_FONT_SIZE', 10))
        self.center_align = Alignment(horizontal="center", vertical="center")
        self.left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

        self.chart_style = {
            'base_width': 5,
            'width_per_point': 0.3,
            'height': 4,
            'dpi': 150,
            'bar_width': 0.6,
            'bar_alpha': 0.85,
            'colors': plt.cm.Set1(np.linspace(0, 1, 9))
        }

        self.use_chart_data_extraction = getattr(souren_config, 'USE_CHART_DATA_EXTRACTION', False)
        if self.use_chart_data_extraction:
            self.data_extraction_config = getattr(souren_config, 'CHART_DATA_EXTRACTION_CONFIG', {})
            self.chart_steps_filter = []
            print(f"📊 使用命令模式提取图表数据")
            print(f"   配置内容: {self.data_extraction_config}")
        else:
            self.data_extraction_config = {}
            self.chart_steps_filter = getattr(souren_config, 'CHART_STEPS_FILTER', [])
            print(f"📊 使用步骤编号筛选配置")
            print(f"   配置内容: {self.chart_steps_filter}")

    def _extract_all_execution_details(self, data, path="root"):
        all_details = []
        if isinstance(data, dict):
            if 'execution_details' in data and isinstance(data['execution_details'], list) and data['execution_details']:
                config_name = self._get_config_name(data)
                for detail in data['execution_details']:
                    all_details.append((config_name, detail))
            if 'loop_results' in data and isinstance(data['loop_results'], list):
                for idx, loop in enumerate(data['loop_results']):
                    all_details.extend(self._extract_all_execution_details(loop, f"{path}.loop_results[{idx}]"))
            if 'result' in data and isinstance(data['result'], dict):
                all_details.extend(self._extract_all_execution_details(data['result'], f"{path}.result"))
            if 'execution_data' in data and isinstance(data['execution_data'], dict):
                all_details.extend(self._extract_all_execution_details(data['execution_data'], f"{path}.execution_data"))
        elif isinstance(data, list):
            for idx, item in enumerate(data):
                all_details.extend(self._extract_all_execution_details(item, f"{path}[{idx}]"))
        return all_details

    def _get_config_name(self, result):
        script_file = result.get('script_file', result.get('file', result.get('script_name', '')))
        if isinstance(script_file, str):
            if os.path.sep in script_file:
                script_file = os.path.basename(script_file)
            if script_file.endswith('.py'):
                script_file = script_file[:-3]

        params = {}
        if 'parameters' in result and isinstance(result['parameters'], dict):
            params = result['parameters']
        elif 'result' in result and isinstance(result['result'], dict) and 'parameters' in result['result']:
            params = result['result']['parameters']
        elif 'result' in result and isinstance(result['result'], dict) and 'result' in result['result']:
            inner = result['result']['result']
            if isinstance(inner, dict) and 'parameters' in inner:
                params = inner['parameters']

        param_str = ''
        if params:
            param_parts = []
            if 'lineLoss' in params: param_parts.append(f"线损:{params['lineLoss']}")
            if 'band' in params: param_parts.append(f"band:{params['band']}")
            if 'bw' in params: param_parts.append(f"bw:{params['bw']}")
            if 'scs' in params: param_parts.append(f"scs:{params['scs']}")
            if 'range' in params: param_parts.append(f"range:{params['range']}")
            param_str = f" ({', '.join(param_parts)})"

        base_name = script_file if script_file else '未知脚本'
        return f"{base_name}{param_str}"

    def _get_config_params(self, result):
        params = result.get('parameters', {})
        if not params and 'result' in result and isinstance(result['result'], dict):
            params = result['result'].get('parameters', {})
        return params

    def find_latest_json_result(self) -> Optional[str]:
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
        try:
            if json_file is None:
                json_file = self.find_latest_json_result()
                if json_file is None:
                    print("❌ 未找到JSON结果文件")
                    return False

            json_name = os.path.basename(json_file)
            excel_name = json_name.replace('.json', '.xlsx')
            excel_file = os.path.join(os.path.dirname(json_file), excel_name)

            if os.path.exists(excel_file):
                try:
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
                    backup_file = excel_file.replace('.xlsx', f'_{timestamp}.xlsx')
                    os.rename(excel_file, backup_file)
                    print(f"⚠️ 原文件已存在，已重命名为: {os.path.basename(backup_file)}")
                except:
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    excel_name = f"souren_results_{timestamp}.xlsx"
                    excel_file = os.path.join(os.path.dirname(json_file), excel_name)
                    print(f"⚠️ 文件被锁定，使用新文件名: {excel_name}")

            results = self.load_json_results(json_file)
            if not results:
                print("❌ 未加载到有效的JSON结果数据")
                return False

            print(f"📊 开始创建Excel文件: {excel_name}")

            with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
                self._create_summary_sheet(writer, results)
                self._create_execution_details_sheet(writer, results)
                self._create_charts_sheet(writer, results)

            self._apply_styles(excel_file, results)
            print(f"✅ Excel文件已生成: {excel_file}")
            return True
        except Exception as e:
            print(f"❌ Excel文件导出失败: {e}")
            import traceback
            traceback.print_exc()
            return False

    def _create_summary_sheet(self, writer: pd.ExcelWriter, results: List[Dict]):
        try:
            summary_data = []
            config_groups = {}

            all_test_results = []
            for result in results:
                if 'loop_results' in result:
                    all_test_results.extend(result['loop_results'])
                else:
                    all_test_results.append(result)

            for test in all_test_results:
                config_name = self._get_config_name(test)
                if config_name not in config_groups:
                    config_groups[config_name] = []
                config_groups[config_name].append(test)

            config_index = 1
            for config_name, config_results in config_groups.items():
                total_executed = 0
                total_passed = 0
                total_failed = 0
                total_duration = 0
                loop_count = len(config_results)
                first_result = config_results[0]

                for result in config_results:
                    exec_details = result.get('execution_details', [])
                    if not exec_details and 'result' in result and isinstance(result['result'], dict):
                        exec_details = result['result'].get('execution_details', [])

                    executed = len([d for d in exec_details if not d.get('is_skipped', False)])
                    passed = len([d for d in exec_details if d.get('status') == 'success'])
                    failed = len([d for d in exec_details if d.get('status') == 'failed'])
                    total_executed += executed
                    total_passed += passed
                    total_failed += failed
                    exec_time = result.get('execution_time', result.get('duration', 0))
                    if isinstance(exec_time, (int, float)):
                        total_duration += float(exec_time)

                success_rate = round((total_passed / total_executed) * 100, 2) if total_executed > 0 else 0
                device = first_result.get('device', 'N/A')
                if isinstance(device, dict):
                    device = device.get('address', 'N/A')
                mode = first_result.get('mode', 'normal')
                status = 'completed' if not any(r.get('interrupted', False) for r in config_results) else 'interrupted'
                total_steps = first_result.get('total_steps', total_executed)

                params = self._get_config_params(first_result)
                param_info = ""
                if params:
                    if 'lineLoss' in params: param_info += f"线损:{params['lineLoss']} "
                    if 'band' in params: param_info += f"band:{params['band']} "
                    if 'bw' in params: param_info += f"bw:{params['bw']} "
                    if 'scs' in params: param_info += f"scs:{params['scs']} "
                    if 'range' in params: param_info += f"range:{params['range']}"

                summary = {
                    "序号": config_index,
                    "执行时间": first_result.get('timestamp_readable', first_result.get('timestamp', 'N/A')),
                    "SCV文件": config_name.split(' (')[0] if ' (' in config_name else config_name,
                    "设备": device,
                    "参数信息": param_info.strip(),
                    "执行模式": mode,
                    "循环次数": loop_count,
                    "总步骤数": total_steps,
                    "已执行步骤": total_executed,
                    "通过步骤": total_passed,
                    "失败步骤": total_failed,
                    "成功率(%)": success_rate,
                    "总耗时(秒)": round(total_duration, 2),
                    "状态": status,
                    "状态消息": f"共执行{loop_count}个循环"
                }
                summary_data.append(summary)
                config_index += 1

            if summary_data:
                df_summary = pd.DataFrame(summary_data)
                columns_order = ["序号","执行时间","SCV文件","设备","参数信息","执行模式","循环次数",
                                 "总步骤数","已执行步骤","通过步骤","失败步骤","成功率(%)","总耗时(秒)","状态","状态消息"]
                df_summary = df_summary[[c for c in columns_order if c in df_summary.columns]]
                df_summary.to_excel(writer, sheet_name='执行汇总', index=False)
                print(f"✅ 汇总表创建成功，共 {len(summary_data)} 行数据")
            else:
                pd.DataFrame({"状态": ["没有有效的执行记录"]}).to_excel(writer, sheet_name='执行汇总', index=False)
        except Exception as e:
            print(f"❌ 创建汇总表失败: {e}")
            traceback.print_exc()

    def _create_execution_details_sheet(self, writer: pd.ExcelWriter, results: List[Dict]):
        try:
            details_data = []
            all_details = self._extract_all_execution_details(results)

            config_map = {}
            for config_name, detail in all_details:
                if config_name not in config_map:
                    config_map[config_name] = []
                config_map[config_name].append(detail)

            config_index = 1
            for config_name, details in config_map.items():
                details.sort(key=lambda d: (d.get('loop_iteration', 1), d.get('step', 0)))
                current_loop = None
                for detail in details:
                    loop_iter = detail.get('loop_iteration', 1)
                    if current_loop is not None and loop_iter != current_loop:
                        for _ in range(2):
                            details_data.append({col: "" for col in self._detail_columns()})
                    current_loop = loop_iter

                    timestamp = ''
                    if 'start_time' in detail and detail['start_time']:
                        try:
                            if isinstance(detail['start_time'], (int, float)):
                                timestamp = datetime.fromtimestamp(detail['start_time']).strftime('%Y-%m-%d %H:%M:%S')
                            else:
                                timestamp = str(detail['start_time'])
                        except:
                            timestamp = str(detail['start_time'])

                    result_str = str(detail.get('result', ''))
                    if len(result_str) > 32767:
                        result_str = result_str[:32767] + '...'

                    record = {
                        "配置": f"配置{config_index}",
                        "执行时间": timestamp,
                        "步骤序号": detail.get('step', ''),
                        "循环轮次": loop_iter,
                        "总循环数": detail.get('loop_count', 1),
                        "步骤内容": detail.get('content', ''),
                        "步骤类型": detail.get('type', 'Normal'),
                        "执行状态": detail.get('status', ''),
                        "执行结果": result_str,
                        "耗时(秒)": round(float(detail.get('duration', 0)), 3)
                    }
                    details_data.append(record)
                for _ in range(2):
                    details_data.append({col: "" for col in self._detail_columns()})
                config_index += 1

            if details_data:
                df_details = pd.DataFrame(details_data, columns=self._detail_columns())
                df_details.to_excel(writer, sheet_name='详细执行记录', index=False)
                print(f"✅ 详细执行记录表创建成功，共 {len(details_data)} 行数据")
            else:
                pd.DataFrame({"状态": ["没有详细的执行记录"]}).to_excel(writer, sheet_name='详细执行记录', index=False)
        except Exception as e:
            print(f"❌ 创建详细记录表失败: {e}")
            traceback.print_exc()

    def _detail_columns(self):
        return ["配置", "执行时间", "步骤序号", "循环轮次", "总循环数",
                "步骤内容", "步骤类型", "执行状态", "执行结果", "耗时(秒)"]

    def _get_colors_for_count(self, n: int) -> List:
        if n <= 20:
            cmap = plt.cm.tab20
            return [cmap(i / 20) for i in range(n)]
        elif n <= 40:
            colors = []
            for i in range(min(n, 20)):
                colors.append(plt.cm.tab20(i / 20))
            if n > 20:
                for i in range(n - 20):
                    colors.append(plt.cm.tab20b(i / 20))
            return colors
        else:
            return [plt.cm.rainbow(i / n) for i in range(n)]

    def _create_charts_sheet(self, writer: pd.ExcelWriter, results: List[Dict]):
        workbook = writer.book
        if '数据分析图表' in workbook.sheetnames:
            std = workbook['数据分析图表']
            workbook.remove(std)
        ws_chart = workbook.create_sheet('数据分析图表')

        all_extracted = self._extract_all_extracted_data(results)

        if not all_extracted:
            ws_chart.cell(row=1, column=1, value="无提取数据，无法生成图表")
            print("⚠️ 没有提取数据，图表工作表仅写入提示")
            return

        config_groups = defaultdict(lambda: defaultdict(list))
        for item in all_extracted:
            cfg = item['config_name']
            cmd = item['command']
            effective_title = item.get('chart_title')
            if not effective_title:
                effective_title = self._simplify_command(cmd)  
            config_groups[cfg][(cmd, effective_title)].append(item)

        row_offset = 1
        chart_count = 0
        dpi = self.chart_style['dpi']
        bar_alpha = self.chart_style['bar_alpha']

        for cfg_name, cmd_title_dict in config_groups.items():
            for (cmd, title), items in cmd_title_dict.items():
                items.sort(key=lambda x: (x['loop_iter'], x.get('seq_in_loop', 0)))

                chart_title = title

                all_have_xlabel = all(it.get('x_label') is not None for it in items)
                if all_have_xlabel:
                    x_labels = [it['x_label'] for it in items]
                    values = [it['value'] for it in items]
                    loops = [it['loop_iter'] for it in items]

                    fig_width = 6 + 0.3 * len(items)
                    fig_height = 5
                    fig, ax = plt.subplots(figsize=(fig_width, fig_height), dpi=dpi)

                    unique_loops = sorted(set(loops))
                    colors = self._get_colors_for_count(len(unique_loops))
                    color_map = {loop: colors[i] for i, loop in enumerate(unique_loops)}
                    bar_colors = [color_map[loop] for loop in loops]

                    x_pos = np.arange(len(items))
                    ax.bar(x_pos, values, color=bar_colors, alpha=bar_alpha, edgecolor='black', linewidth=0.5)

                    for i, (val, label) in enumerate(zip(values, x_labels)):
                        ax.text(i, val * 1.02, f"{val:.3f}\n({label})", ha='center', va='bottom',
                                fontsize=7, fontweight='bold', color='black', linespacing=0.8)

                    ax.set_xticks(x_pos)
                    ax.set_xticklabels(x_labels, rotation=45, ha='right', fontsize=8)

                    ax.set_title(f"{chart_title} - {cfg_name}", fontsize=10, fontweight='bold')
                    ax.set_ylabel('数值', fontsize=9)
                    ax.grid(axis='y', alpha=0.3)

                    from matplotlib.patches import Patch
                    legend_elements = [Patch(facecolor=color_map[loop], edgecolor='black', label=f'循环{loop}') for loop in unique_loops]
                    ax.legend(handles=legend_elements, loc='center left', bbox_to_anchor=(1, 0.5), fontsize=8, title='循环')

                    plt.tight_layout()

                    img_data = BytesIO()
                    plt.savefig(img_data, format='png', dpi=dpi, bbox_inches='tight', facecolor='white')
                    plt.close(fig)
                    img_data.seek(0)

                    img = ExcelImage(img_data)
                    img.width = int(fig_width * 80)
                    img.height = int(fig_height * 80)

                    cell_row = row_offset + (chart_count * 20)
                    cell_addr = f"A{cell_row}"
                    ws_chart.add_image(img, cell_addr)
                    chart_count += 1

                else:
                    loop_groups = defaultdict(list)
                    for item in items:
                        loop_groups[item['loop_iter']].append(item)

                    loops = sorted(loop_groups.keys())
                    if not loops:
                        continue

                    colors = self._get_colors_for_count(len(loops))
                    color_map = {loop: colors[i] for i, loop in enumerate(loops)}

                    group_spacing = 1.2
                    x_positions = []
                    group_labels = []
                    all_values = []
                    current_x = 0
                    for loop in loops:
                        items_in_loop = loop_groups[loop]
                        n = len(items_in_loop)
                        bar_width = 1.0 if n == 1 else 0.5
                        group_width = n * bar_width
                        for j, it in enumerate(items_in_loop):
                            x_positions.append(current_x + bar_width/2 + j*bar_width)
                            all_values.append({
                                'value': it['value'],
                                'loop': loop,
                                'seq': j+1,
                                'color': color_map[loop]
                            })
                        group_labels.append((current_x + group_width / 2, f"循环{loop}"))
                        current_x += group_width + group_spacing

                    fig_width = 6 + 0.5 * len(all_values) + len(cfg_name) * 0.05
                    fig_height = 5
                    fig, ax = plt.subplots(figsize=(fig_width, fig_height), dpi=dpi)

                    for pos, val_info in zip(x_positions, all_values):
                        ax.bar(pos, val_info['value'],
                            width=1.0/len(loop_groups[val_info['loop']]),
                            color=val_info['color'],
                            alpha=bar_alpha,
                            edgecolor='black', linewidth=0.5)

                    for pos, val_info in zip(x_positions, all_values):
                        ax.text(pos, val_info['value'] * 1.02, f"{val_info['value']:.3f}",
                                ha='center', va='bottom',
                                fontsize=7, fontweight='bold', color='black')

                    ax.set_xticks([pos for pos, _ in group_labels])
                    ax.set_xticklabels([label for _, label in group_labels], rotation=45, ha='right', fontsize=8)

                    ax.set_title(f"{chart_title} - {cfg_name}", fontsize=10, fontweight='bold')
                    ax.set_ylabel('数值', fontsize=9)
                    ax.grid(axis='y', alpha=0.3)

                    from matplotlib.patches import Patch
                    legend_elements = [Patch(facecolor=color_map[loop], edgecolor='black',
                                            label=f'循环{loop}') for loop in loops]
                    ax.legend(handles=legend_elements, loc='center left', bbox_to_anchor=(1, 0.5),
                            fontsize=8, title='循环')

                    plt.tight_layout()

                    img_data = BytesIO()
                    plt.savefig(img_data, format='png', dpi=dpi, bbox_inches='tight', facecolor='white')
                    plt.close(fig)
                    img_data.seek(0)

                    img = ExcelImage(img_data)
                    img.width = int(fig_width * 80)
                    img.height = int(fig_height * 80)

                    cell_row = row_offset + (chart_count * 20)
                    cell_addr = f"A{cell_row}"
                    ws_chart.add_image(img, cell_addr)
                    chart_count += 1

        print(f"✅ 数据分析图表已生成，共 {chart_count} 个柱状图")

    def _simplify_command(self, cmd: str) -> str:
        parts = cmd.split(':')
        if parts:
            last = parts[-1].replace('?', '').strip()
            return last if last else cmd[:20]
        return cmd[:20]

    def _extract_all_extracted_data(self, data, path="root"):
        extracted_items = []
        if isinstance(data, dict):
            if 'extracted_data' in data and isinstance(data['extracted_data'], list):
                config_name = self._get_config_name(data)
                seq_counter = defaultdict(int)
                for item in data['extracted_data']:
                    key = (config_name, item.get('command', ''), item.get('loop_iteration', 1))
                    seq_counter[key] += 1
                    extracted_items.append({
                        'command': item.get('command', ''),
                        'loop_iter': item.get('loop_iteration', 1),
                        'value': item.get('extracted_data', 'N/A'),
                        'config_name': config_name,
                        'seq_in_loop': seq_counter[key],
                        'chart_title': item.get('chart_title'),
                        'x_label': item.get('x_label')
                    })
            for key, value in data.items():
                if key not in ('extracted_data',) and isinstance(value, (dict, list)):
                    extracted_items.extend(self._extract_all_extracted_data(value, f"{path}.{key}"))
        elif isinstance(data, list):
            for idx, item in enumerate(data):
                extracted_items.extend(self._extract_all_extracted_data(item, f"{path}[{idx}]"))
        return extracted_items

    def _apply_styles(self, excel_file: str, results: List[Dict]):
        try:
            wb = load_workbook(excel_file)
            if '执行汇总' in wb.sheetnames:
                self._format_summary_sheet(wb['执行汇总'])
            if '详细执行记录' in wb.sheetnames:
                self._format_details_sheet(wb['详细执行记录'])
            self._adjust_dimensions(wb)
            wb.save(excel_file)
        except Exception as e:
            print(f"❌ 应用样式失败: {e}")
            traceback.print_exc()

    def _format_summary_sheet(self, ws):
        for cell in ws[1]:
            cell.fill = self.header_fill
            cell.font = self.header_font
            cell.alignment = self.center_align
            cell.border = self._create_border(self.border_colors.get("default", "000000"))
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                cell.border = self._create_border(self.border_colors.get("default", "000000"))
                cell.font = self.data_font
                col_letter = get_column_letter(cell.column)
                if col_letter in ["A","F","G","H","I","J","K","L","M","N"]:
                    cell.alignment = self.center_align
                else:
                    cell.alignment = self.left_align
            ws.row_dimensions[row[0].row].height = self.default_row_height

    def _format_details_sheet(self, ws):
        for cell in ws[1]:
            cell.fill = self.header_fill
            cell.font = self.header_font
            cell.alignment = self.center_align
            cell.border = self._create_border(self.border_colors.get("default", "000000"))
        ws.row_dimensions[1].height = self.header_row_height

        for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row), start=2):
            if not any(cell.value for cell in row):
                continue

            for cell in row:
                cell.border = self._create_border(self.border_colors.get("default", "000000"))
                cell.font = self.data_font
                col_idx = cell.column
                if col_idx in [1,3,4,5,7,8,10]:
                    cell.alignment = self.center_align
                elif col_idx == 2:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                else:
                    cell.alignment = self.left_align

            if len(row) > 7:
                status_cell = row[7]
                if status_cell.value and str(status_cell.value).strip().lower() == "failed":
                    status_cell.font = self.failed_font

            ws.row_dimensions[row_idx].height = self.default_row_height

    def _create_border(self, color):
        side = Side(style='thin', color=color)
        return Border(left=side, right=side, top=side, bottom=side)

    def _adjust_dimensions(self, wb):
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            if sheet_name == '_ChartData' or sheet_name.startswith('_'):
                continue

            for row in range(1, ws.max_row + 1):
                if row == 1:
                    ws.row_dimensions[row].height = self.header_row_height
                else:
                    if ws.row_dimensions[row].height is None:
                        ws.row_dimensions[row].height = self.default_row_height

            if sheet_name == '执行汇总':
                for col in range(1, ws.max_column + 1):
                    col_letter = get_column_letter(col)
                    ws.column_dimensions[col_letter].width = self.default_column_width
                for col_letter, width in self.summary_column_widths.items():
                    if col_letter in ws.column_dimensions:
                        ws.column_dimensions[col_letter].width = width

            elif sheet_name == '详细执行记录':
                for col in range(1, ws.max_column + 1):
                    col_letter = get_column_letter(col)
                    ws.column_dimensions[col_letter].width = self.default_column_width
                for col_letter, width in self.details_column_widths.items():
                    if col_letter in ws.column_dimensions:
                        ws.column_dimensions[col_letter].width = width

            elif sheet_name == '数据分析图表':
                for col in range(1, ws.max_column + 1):
                    col_letter = get_column_letter(col)
                    ws.column_dimensions[col_letter].width = self.default_column_width

            else:
                for col in range(1, ws.max_column + 1):
                    col_letter = get_column_letter(col)
                    ws.column_dimensions[col_letter].width = self.default_column_width

            if ws.max_row > 1 and ws.max_column > 0:
                try:
                    ws.auto_filter.ref = ws.dimensions
                except:
                    pass

        print(f"✅ 已应用所有Excel导出配置")


def create_summary_excel(main_execution_dir: str, all_results: List[Dict] = None) -> Optional[str]:
    try:
        print(f"\n{'='*60}")
        print("📊 开始创建图表汇总Excel文件")
        print(f"{'='*60}")

        if not main_execution_dir or not os.path.exists(main_execution_dir):
            print("❌ 主执行目录不存在，无法创建汇总文件")
            return None

        subdirs = []
        for item in os.listdir(main_execution_dir):
            item_path = os.path.join(main_execution_dir, item)
            if os.path.isdir(item_path) and not item.startswith('_') and not item.startswith('.'):
                subdirs.append(item_path)

        if not subdirs:
            print("❌ 主执行目录下没有配置子目录")
            return None

        print(f"📁 找到 {len(subdirs)} 个配置子目录")

        extracted_dict = {}

        for subdir in subdirs:
            config_name = os.path.basename(subdir)
            json_files = []
            for file in os.listdir(subdir):
                if file.endswith('.json') and not file.startswith('summary'):
                    json_files.append(os.path.join(subdir, file))
            if not json_files:
                print(f"⚠️  目录 {config_name} 中没有找到 JSON 结果文件，跳过")
                continue
            json_file = max(json_files, key=os.path.getmtime)
            try:
                with open(json_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                extract_list = _extract_extracted_data_from_json(data, config_name)
                for item in extract_list:
                    key = (item['config_name'], item['command'], item['loop_iter'], item.get('seq_in_loop', 0))
                    if key not in extracted_dict:
                        extracted_dict[key] = item
            except Exception as e:
                print(f"❌ 处理 {config_name} 的 JSON 文件失败: {e}")
                continue

        if not extracted_dict:
            print("❌ 没有提取到任何图表数据")
            return None

        all_extracted = list(extracted_dict.values())
        print(f"📊 共提取到 {len(all_extracted)} 条唯一的提取数据")

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        summary_file = os.path.join(main_execution_dir, f"summary_charts_{timestamp}.xlsx")
        wb = Workbook()

        ws_summary = wb.active
        ws_summary.title = "图表汇总"
        _populate_chart_summary_sheet(ws_summary, all_extracted)

        ws_comparison = wb.create_sheet("横向对比")
        _populate_comparison_sheet(ws_comparison, all_extracted)

        wb.save(summary_file)
        print(f"\n{'='*60}")
        print(f"✅ 图表汇总Excel文件创建成功!")
        print(f"📁 文件路径: {summary_file}")
        print(f"📊 工作表: 图表汇总, 横向对比")
        print(f"{'='*60}")

        return summary_file

    except Exception as e:
        print(f"❌ 创建图表汇总Excel文件失败: {e}")
        import traceback
        traceback.print_exc()
        return None


def _extract_extracted_data_from_json(data, config_name):
    extracted = []
    if isinstance(data, dict):
        if 'extracted_data' in data and isinstance(data['extracted_data'], list):
            seq_counter = defaultdict(int)
            for item in data['extracted_data']:
                key = (config_name, item.get('command', ''), item.get('loop_iteration', 1))
                seq_counter[key] += 1
                extracted.append({
                    'command': item.get('command', ''),
                    'loop_iter': item.get('loop_iteration', 1),
                    'value': item.get('extracted_data', None),
                    'config_name': config_name,
                    'seq_in_loop': seq_counter[key],
                    'chart_title': item.get('chart_title'),
                    'x_label': item.get('x_label')
                })
        for key, value in data.items():
            if key != 'extracted_data' and isinstance(value, (dict, list)):
                extracted.extend(_extract_extracted_data_from_json(value, config_name))
    elif isinstance(data, list):
        for item in data:
            extracted.extend(_extract_extracted_data_from_json(item, config_name))
    return extracted


def _populate_chart_summary_sheet(ws, all_extracted):
    config_groups = defaultdict(lambda: defaultdict(list))
    for item in all_extracted:
        cfg = item['config_name']
        cmd = item['command']
        effective_title = item.get('chart_title') or _simplify_command_name(cmd)
        config_groups[cfg][(cmd, effective_title)].append(item)

    dpi = 150
    bar_alpha = 0.85

    row = 1
    chart_count = 0

    for cfg_name, cmd_title_dict in config_groups.items():
        title_cell = ws.cell(row=row, column=1, value=f"配置: {cfg_name}")
        title_cell.font = Font(size=12, bold=True)
        row += 1

        for (cmd, title), items in cmd_title_dict.items():
            items.sort(key=lambda x: (x['loop_iter'], x.get('seq_in_loop', 0)))

            chart_title = title

            all_have_xlabel = all(it.get('x_label') is not None for it in items)
            if all_have_xlabel:
                x_labels = [it['x_label'] for it in items]
                values = [it['value'] for it in items]
                loops = [it['loop_iter'] for it in items]

                fig_width = 6 + 0.3 * len(items)
                fig_height = 5
                fig, ax = plt.subplots(figsize=(fig_width, fig_height), dpi=dpi)

                unique_loops = sorted(set(loops))
                colors = _get_colors_for_count(len(unique_loops))
                color_map = {loop: colors[i] for i, loop in enumerate(unique_loops)}
                bar_colors = [color_map[loop] for loop in loops]

                x_pos = np.arange(len(items))
                ax.bar(x_pos, values, color=bar_colors, alpha=bar_alpha, edgecolor='black', linewidth=0.5)

                for i, (val, label) in enumerate(zip(values, x_labels)):
                    ax.text(i, val * 1.02, f"{val:.3f}\n({label})", ha='center', va='bottom',
                            fontsize=7, fontweight='bold', color='black', linespacing=0.8)

                ax.set_xticks(x_pos)
                ax.set_xticklabels(x_labels, rotation=45, ha='right', fontsize=8)

                ax.set_title(f"{chart_title}\n{cfg_name[:30]}", fontsize=10, fontweight='bold')
                ax.set_ylabel('数值', fontsize=9)
                ax.grid(axis='y', alpha=0.3)

                from matplotlib.patches import Patch
                legend_elements = [Patch(facecolor=color_map[loop], edgecolor='black', label=f'循环{loop}') for loop in unique_loops]
                ax.legend(handles=legend_elements, loc='center left', bbox_to_anchor=(1, 0.5), fontsize=8, title='循环')

                plt.tight_layout()

                img_data = BytesIO()
                plt.savefig(img_data, format='png', dpi=dpi, bbox_inches='tight', facecolor='white')
                plt.close(fig)
                img_data.seek(0)

                img = ExcelImage(img_data)
                img.width = int(fig_width * 80)
                img.height = int(fig_height * 80)

                cell_addr = f"A{row}"
                ws.add_image(img, cell_addr)

                img_height_rows = int(fig_height * 5)
                row += img_height_rows + 4
                chart_count += 1

            else:
                loop_groups = defaultdict(list)
                for it in items:
                    loop_groups[it['loop_iter']].append(it)

                loops = sorted(loop_groups.keys())
                if not loops:
                    continue

                colors = _get_colors_for_count(len(loops))
                color_map = {loop: colors[i] for i, loop in enumerate(loops)}

                group_spacing = 1.2
                x_positions = []
                group_labels = []
                all_values = []
                current_x = 0
                for loop in loops:
                    items_in_loop = loop_groups[loop]
                    n = len(items_in_loop)
                    bar_width = 1.5  
                    group_width = n * bar_width
                    for j, it in enumerate(items_in_loop):
                        x_positions.append(current_x + bar_width/2 + j*bar_width)
                        all_values.append({
                            'value': it['value'],
                            'loop': loop,
                            'seq': j+1,
                            'color': color_map[loop],
                            'width': bar_width
                        })
                    group_labels.append((current_x + group_width / 2, f"循环{loop}"))
                    current_x += group_width + group_spacing

                fig_width = 6 + 0.3 * len(all_values)
                fig_height = 5
                fig, ax = plt.subplots(figsize=(fig_width, fig_height), dpi=dpi)

                for pos, val_info in zip(x_positions, all_values):
                    ax.bar(pos, val_info['value'],
                           width=val_info['width'],
                           color=val_info['color'],
                           alpha=bar_alpha,
                           edgecolor='black', linewidth=0.5)

                for pos, val_info in zip(x_positions, all_values):
                    ax.text(pos, val_info['value'] * 1.02, f"{val_info['value']:.3f}",
                            ha='center', va='bottom',
                            fontsize=7, fontweight='bold', color='black')

                ax.set_xticks([pos for pos, _ in group_labels])
                ax.set_xticklabels([label for _, label in group_labels], rotation=45, ha='right', fontsize=8)

                ax.set_title(f"{chart_title}\n{cfg_name[:30]}", fontsize=10, fontweight='bold')
                ax.set_ylabel('数值', fontsize=9)
                ax.grid(axis='y', alpha=0.3)

                from matplotlib.patches import Patch
                legend_elements = [Patch(facecolor=color_map[loop], edgecolor='black',
                                         label=f'循环{loop}') for loop in loops]
                ax.legend(handles=legend_elements, loc='center left', bbox_to_anchor=(1, 0.5),
                          fontsize=8, title='循环')

                plt.tight_layout()

                img_data = BytesIO()
                plt.savefig(img_data, format='png', dpi=dpi, bbox_inches='tight', facecolor='white')
                plt.close(fig)
                img_data.seek(0)

                img = ExcelImage(img_data)
                img.width = int(fig_width * 80)
                img.height = int(fig_height * 80)

                cell_addr = f"A{row}"
                ws.add_image(img, cell_addr)

                img_height_rows = int(fig_height * 5)
                row += img_height_rows + 4
                chart_count += 1

        row += 2
    ws.column_dimensions['A'].width = 30
    print(f"   ✅ 图表汇总工作表已生成，共 {chart_count} 个图表")


def _populate_comparison_sheet(ws, all_extracted):
    cmd_title_groups = defaultdict(lambda: defaultdict(list))
    for item in all_extracted:
        cmd = item['command']
        cfg = item['config_name']
        effective_title = item.get('chart_title') or _simplify_command_name(cmd)
        cmd_title_groups[(cmd, effective_title)][cfg].append(item)

    comparison_items = {key: cfg_dict for key, cfg_dict in cmd_title_groups.items() if len(cfg_dict) >= 2}

    if not comparison_items:
        ws.cell(row=1, column=1, value="无横向对比数据")
        return

    BASE_HEIGHT = 5.0
    HEIGHT_PER_LOOP = 0.4
    LOOP_SPACING = 0.2
    CONFIG_SPACING = 0.5
    FONT_SIZE = 6.0
    BASE_CONFIG_WIDTH = 2.5
    CHAR_WIDTH_FACTOR = 0.1
    TEXT_PADDING = 1.2
    MULTI_BAR_FACTOR = 0.6

    dpi = 150
    bar_alpha = 0.85

    row = 1
    chart_count = 0

    for (cmd, title), config_items in comparison_items.items():
        config_names = list(config_items.keys())
        config_data = {}
        all_loops = set()
        for cfg, items in config_items.items():
            items_sorted = sorted(items, key=lambda x: (x['loop_iter'], x.get('seq_in_loop', 0)))
            data_points = []
            for it in items_sorted:
                loop = it['loop_iter']
                seq = it.get('seq_in_loop', 1)
                data_points.append((loop, seq, it['value'], it.get('x_label')))
                all_loops.add(loop)
            config_data[cfg] = data_points

        if not config_data:
            continue

        loops = sorted(all_loops)
        colors = _get_colors_for_count(len(loops))
        color_map = {loop: colors[i] for i, loop in enumerate(loops)}

        cfg_loop_info = {}
        for cfg in config_names:
            loop_info = {}
            for loop in loops:
                points = [(l, seq, val, xl) for (l, seq, val, xl) in config_data[cfg] if l == loop]
                if not points:
                    continue
                points.sort(key=lambda x: x[1])
                n = len(points)
                texts = []
                values = []
                x_labels = []
                for (l, seq, val, xl) in points:
                    if n > 1:
                        label = f"{val:.3f}({seq})"
                    else:
                        label = f"{val:.3f}"
                    texts.append(label)
                    values.append(val)
                    x_labels.append(xl if xl else '')
                max_len = max(len(t) for t in texts)
                if n > 1:
                    min_bar_width = max_len * CHAR_WIDTH_FACTOR * TEXT_PADDING * MULTI_BAR_FACTOR
                else:
                    min_bar_width = max_len * CHAR_WIDTH_FACTOR * TEXT_PADDING
                loop_info[loop] = (n, texts, values, min_bar_width, x_labels)
            cfg_loop_info[cfg] = loop_info

        cfg_required_width = {}
        for cfg in config_names:
            present_loops = [loop for loop in loops if loop in cfg_loop_info[cfg]]
            if not present_loops:
                cfg_required_width[cfg] = 0
                continue
            total = 0
            for loop in present_loops:
                n, texts, values, min_bar_width, x_labels = cfg_loop_info[cfg][loop]
                total += n * min_bar_width
            total += (len(present_loops) - 1) * LOOP_SPACING
            cfg_required_width[cfg] = total

        cfg_widths = {}
        for cfg in config_names:
            cfg_widths[cfg] = max(BASE_CONFIG_WIDTH, cfg_required_width.get(cfg, 0))

        cfg_start_x = {}
        current_x = 0
        for cfg in config_names:
            cfg_start_x[cfg] = current_x
            current_x += cfg_widths[cfg] + CONFIG_SPACING

        total_fig_width = current_x - CONFIG_SPACING
        fig_height = BASE_HEIGHT + HEIGHT_PER_LOOP * (len(loops) - 1)
        fig_width = total_fig_width

        n_configs = len(config_names)
        if n_configs >= 8:
            label_fontsize = 7
            rotation = 60
        elif n_configs >= 5:
            label_fontsize = 8
            rotation = 45
        else:
            label_fontsize = 9
            rotation = 30

        fig, ax = plt.subplots(figsize=(fig_width, fig_height), dpi=dpi)

        for cfg in config_names:
            start_x = cfg_start_x[cfg]
            present_loops = [loop for loop in loops if loop in cfg_loop_info[cfg]]
            if not present_loops:
                continue

            m = len(present_loops)
            if m == 0:
                continue
            total_required_group = sum(n * min_bar_width for (n, _, _, min_bar_width, _) in [cfg_loop_info[cfg][loop] for loop in present_loops])
            available_group_width = cfg_widths[cfg] - (m - 1) * LOOP_SPACING
            if available_group_width <= 0:
                group_widths = {loop: n * min_bar_width for loop, (n, _, _, min_bar_width, _) in cfg_loop_info[cfg].items()}
            else:
                ratio = available_group_width / total_required_group if total_required_group > 0 else 1.0
                group_widths = {}
                for loop in present_loops:
                    n, texts, values, min_bar_width, x_labels = cfg_loop_info[cfg][loop]
                    group_widths[loop] = n * min_bar_width * ratio

            current_x_pos = start_x
            for loop in present_loops:
                n, texts, values, min_bar_width, x_labels = cfg_loop_info[cfg][loop]
                group_width = group_widths[loop]
                bar_width = group_width / n
                points = [(l, seq, val, xl) for (l, seq, val, xl) in config_data[cfg] if l == loop]
                points.sort(key=lambda x: x[1])
                for j, (l, seq, val, xl) in enumerate(points):
                    pos = current_x_pos + (j + 0.5) * bar_width
                    ax.bar(pos, val,
                           width=bar_width,
                           color=color_map[loop],
                           alpha=bar_alpha,
                           edgecolor='black', linewidth=0.5)

                    if xl:
                        label = f"{val:.3f}\n({xl})"
                    else:
                        label = f"{val:.3f}"
                    ax.text(pos, val/2, label,
                            ha='center', va='center',
                            fontsize=FONT_SIZE, fontweight='bold', color='black')
                current_x_pos += group_width + LOOP_SPACING

        config_centers = [cfg_start_x[cfg] + cfg_widths[cfg]/2 for cfg in config_names]
        ax.set_xticks(config_centers)
        ax.set_xticklabels(config_names, rotation=rotation, ha='right', fontsize=label_fontsize)

        ax.set_title(title, fontsize=12, fontweight='bold')   
        ax.set_xlabel('配置', fontsize=11)
        ax.set_ylabel('数值', fontsize=11)
        ax.grid(axis='y', alpha=0.3)

        from matplotlib.patches import Patch
        legend_elements = [Patch(facecolor=color_map[loop], edgecolor='black',
                                 label=f'循环{loop}') for loop in loops]
        ax.legend(handles=legend_elements, loc='center left', bbox_to_anchor=(1, 0.5),
                  fontsize=9, title='循环', title_fontsize=10)

        ax.margins(x=0.01)

        plt.tight_layout()

        img_data = BytesIO()
        plt.savefig(img_data, format='png', dpi=dpi,
                    bbox_inches='tight', facecolor='white')
        plt.close(fig)
        img_data.seek(0)

        img = ExcelImage(img_data)
        img.width = int(fig_width * 100)
        img.height = int(fig_height * 100)

        cell_addr = f"A{row}"
        ws.add_image(img, cell_addr)

        img_height_rows = int(fig_height * 5)
        info_row = row + img_height_rows
        ws.cell(row=info_row, column=1, value=f"命令: {title}")
        ws.cell(row=info_row + 1, column=1, value=f"配置数量: {n_configs}")
        ws.cell(row=info_row + 2, column=1, value=f"循环数: {len(loops)}")

        row = info_row + 3 + 4
        chart_count += 1

    ws.column_dimensions['A'].width = 40
    print(f"   ✅ 横向对比工作表已生成，共 {chart_count} 个增强对比图表")


def _get_colors_for_count(n: int) -> List:
    if n <= 20:
        return [plt.cm.tab20(i / 20) for i in range(n)]
    elif n <= 40:
        colors = []
        for i in range(min(n, 20)):
            colors.append(plt.cm.tab20(i / 20))
        if n > 20:
            for i in range(n - 20):
                colors.append(plt.cm.tab20b(i / 20))
        return colors
    else:
        return [plt.cm.rainbow(i / n) for i in range(n)]


def _simplify_command_name(cmd: str) -> str:
    parts = cmd.split(':')
    if parts:
        last = parts[-1].replace('?', '').strip()
        return last if last else cmd[:20]
    return cmd[:20]


if __name__ == "__main__":
    exporter = ResultExporter()
    success = exporter.convert_to_excel()
    if success:
        print("✅ Excel文件导出成功!")
    else:
        print("❌ Excel文件导出失败")