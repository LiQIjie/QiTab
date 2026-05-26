// QiTab.cs - 面板主控件，负责与 SheetTabList 交互并对外暴露事件供宿主使用
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Drawing;
/*文件编码：UTF-8 无BOM*/
namespace Qtab
{
    /// <summary>
    /// `QiTab` 是任务面板中的主控件，职责包括：
    /// - 展示搜索框与 `SheetTabList`（用于显示工作表和分组列表）；
    /// - 提供一个用于切换停靠位置（左/右）的按钮；
    /// - 将用户的操作（激活工作表、分组/取消分组、着色、重命名等）通过事件通知给 VSTO 宿主（`ThisAddIn`）。
    /// 控件本身不直接操作 Excel，对 Excel 的变更由宿主处理，本控件仅充当 UI 与宿主之间的桥梁。
    /// </summary>
    public class QTabNavi : UserControl
    {
        // ============ 事件（供宿主订阅）===========
        public event Action<string> OnSheetSelected;                    // 用户请求激活某个工作表时触发
        public event Action DockLeftRequested;                          // 请求将任务面板停靠到左侧
        public event Action DockRightRequested;                         // 请求将任务面板停靠到右侧
        public event Action<List<string>> GroupRequested;               // 请求将选中工作表分组
        public event Action<List<string>> UngroupRequested;             // 请求取消选中工作表的分组
        public event Action<List<string>> RenameGroupRequested;         // 请求对某个分组进行重命名（触发宿主显示输入框或内联编辑）
        public event Action<List<string>, Color> ColorRequested;        // 请求设置选中工作表的标签颜色
        public event Action<string> RenameSheetRequested;               // 请求重命名工作表（宿主执行实际重命名）
        public event Action<string, string> GroupRenameCommitted;       // 分组重命名已提交（参数：旧名，新名）
        public event Action<string, string> RenameSheetCommitted;       // 新增：当用户在列表内完成工作表的内联重命名后，向宿主发送旧名与新名
        public event Action<List<string>> DeleteSheetRequested;         // 新增：请求删除选中工作表
        public event Action<string, Point> ShowExcelContextMenu;        // 新增：Excel 原生菜单显示请求

        // 新增：请求切换 Excel 工具（Solver + DataAnalysis 合并）
        public event Action ToolsToggleRequested;

        // 拖拽事件
        public event Action<string, string, bool> SheetDragDropRequested;  // (draggedSheet, targetSheet, insertBefore)
        public event Action<string, string, bool> GroupDragDropRequested;  // (draggedGroup, targetItem, insertBefore)

        // 搜索输入框（顶部）
        private TextBox _search = new TextBox { Dock = DockStyle.Top };
        // 工作表/分组列表（中间区域，具体列表逻辑在 SheetTabList 中实现）
        private QtabList _list = new QtabList { Dock = DockStyle.Fill };
        // 当前所有工作表名（用于本地搜索与过滤）
        private List<string> _all = new List<string>();

        // 底部按钮区域（用于放置 Dock 切换按钮和统计信息）
        private FlowLayoutPanel _buttonPanel;
        private Button _btnDockToggle;           // 切换停靠位置的按钮：左停靠时显示 "->"，右停靠时显示 "<-"
        private Label _statsLabel;               // 显示统计信息（总数、隐藏数）
        private bool _isLeftDock = true;         // 当前停靠侧（默认左侧）
        
        // 统计信息（由 LoadSheets 传入）
        private int _totalSheets = 0;
        private int _hiddenSheets = 0;
        private int _matchSheets = 0;

        // 标志：程序性修改搜索框文本时，禁止触发过滤
        private bool _suppressFilter = false;

        // === 计算区控件 ===
        private RadioButton _rbT; // 计算扭矩
        private RadioButton _rbN; // 计算转速
        private RadioButton _rbP; // 计算功率
        private TextBox _txtT;    // 输入或显示 T (N·m)
        private TextBox _txtN;    // 输入或显示 N (rpm)
        private TextBox _txtP;    // 输入或显示 P (W / kW)
        // 移除按钮：改为自动计算
        // private Button _btnCalc;  // 计算按钮

        // 合并：Excel 工具快捷按钮（同时控制 Solver 与 Data Analysis）
        private Button _btnTools;

        // 防止自动计算时递归触发 TextChanged
        private bool _suppressAutoCalc = false;

        public QTabNavi()
        {
            // 使用表格布局：上 - 搜索，中 - 列表，下 - 按钮区
            var panel = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3, BackColor = Color.WhiteSmoke };
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // 搜索行
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));   // 列表行（填充剩余空间）
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // 按钮行

            // 初始化 Dock 切换按钮与统计标签
            _buttonPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Padding = new Padding(0) };
            _statsLabel = new Label { AutoSize = true, Text = "Total: 0 | Hidden: 0", Margin = new Padding(0, 4, 2, 0), ForeColor = Color.DimGray };
            _btnDockToggle = new Button { Width= 10, AutoSize = true, Text = "->", Margin = new Padding(2, 0, 0, 0) };
            _btnDockToggle.Click += (s, e) => { if (_isLeftDock) DockRightRequested?.Invoke(); else DockLeftRequested?.Invoke(); };

            // 计算区容器（置于统计与停靠按钮上方，以符合示意图）
            var calcPanel = new FlowLayoutPanel
            {
                AutoSize = true,
                FlowDirection = FlowDirection.LeftToRight,
                Margin = new Padding(4, 4, 4, 0)
            };

            _rbT = new RadioButton { Text = "", Checked = false, AutoSize = true, Margin = new Padding(0, 0, 0, 0) };
            _txtT = new TextBox { Width = 40, Margin = new Padding(0, 0, 0, 0) };
            var tUnit = new Label { Text = "N·m", AutoSize = true, ForeColor = Color.Firebrick, Margin = new Padding(0, 0, 0, 0) };

            _rbP = new RadioButton { Text = "", Checked = true, AutoSize = true, Margin = new Padding(0, 0, 0, 0) };
            _txtP = new TextBox { Width = 60, Margin = new Padding(0, 0, 0, 0) };
            var pUnit = new Label { Text = "W", AutoSize = true, ForeColor = Color.Firebrick, Margin = new Padding(0, 0, 0, 0) };

            // 合并工具按钮（默认显示为 OFF，由宿主刷新真实状态）
            _btnTools = new Button { AutoSize = true, Margin = new Padding(2, 0, 0, 0), Text = "ANSYS:OFF", BackColor = Color.LightCoral };
            _btnTools.Click += (s, e) => ToolsToggleRequested?.Invoke();

            _rbN = new RadioButton { Text = "", Checked = false, AutoSize = true, Margin = new Padding(4, 0, 0, 0) };
            _txtN = new TextBox { Width = 40, Margin = new Padding(0, 0, 0, 0) };
            var nUnit = new Label { Text = "rpm", AutoSize = true, ForeColor = Color.Firebrick, Margin = new Padding(0, 0, 0, 0) };

            // 组装计算区（顺序：T、N、P、合并按钮）
            calcPanel.Controls.Add(_rbT); calcPanel.Controls.Add(_txtT); calcPanel.Controls.Add(tUnit);
            calcPanel.Controls.Add(_rbN); calcPanel.Controls.Add(_txtN); calcPanel.Controls.Add(nUnit);
            calcPanel.Controls.Add(_rbP); calcPanel.Controls.Add(_txtP); calcPanel.Controls.Add(pUnit);
            calcPanel.Controls.Add(_btnTools);
            // 添加到底部区域：先计算区，再统计与停靠
            _buttonPanel.Controls.Add(calcPanel);
            _buttonPanel.Controls.Add(_statsLabel);
            _buttonPanel.Controls.Add(_btnDockToggle);

            // 搜索框占位与行为
            _search.ForeColor = System.Drawing.Color.Gray;
            _search.Text = "Search sheets...";    // 占位文本

            _search.GotFocus += (s, e) =>
            {
                if (_search.Text == "Search sheets...")
                {
                    // 程序性修改文本时不做过滤
                    _suppressFilter = true;
                    _search.Text = string.Empty;
                    _search.ForeColor = System.Drawing.Color.Black;
                    _suppressFilter = false;
                }
            };
            _search.LostFocus += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(_search.Text))
                {
                    _suppressFilter = true;
                    _search.Text = "Search sheets...";
                    _search.ForeColor = System.Drawing.Color.Gray;
                    _suppressFilter = false;
                }
            };
            _search.TextChanged += (s, e) =>
            {
                // 程序性修改或仍为占位文本时，不触发过滤
                if (_suppressFilter) return;
                if (_search.ForeColor == System.Drawing.Color.Gray) return;
                Filter(_search.Text);
            };
            _search.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Escape && _search.ForeColor != System.Drawing.Color.Gray)
                {
                    _suppressFilter = true;
                    _search.Text = string.Empty;
                    _suppressFilter = false;
                    Filter(string.Empty);
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                }
            };

            // 将 `SheetTabList` 的事件转发给宿主（通过本控件暴露的事件）
            _list.OnSheetSelected += s => OnSheetSelected?.Invoke(s);
            _list.GroupRequested += names => GroupRequested?.Invoke(names);
            _list.UngroupRequested += names => UngroupRequested?.Invoke(names);
            _list.RenameGroupRequested += names => RenameGroupRequested?.Invoke(names);
            _list.ColorRequested += (names, color) => ColorRequested?.Invoke(names, color);
            _list.RenameSheetRequested += s => RenameSheetRequested?.Invoke(s);
            _list.GroupRenameCommitted += (oldName, newName) => GroupRenameCommitted?.Invoke(oldName, newName);
            _list.RenameSheetCommitted += (oldName, newName) => RenameSheetCommitted?.Invoke(oldName, newName);
            _list.DeleteSheetRequested += names => DeleteSheetRequested?.Invoke(names);
            _list.ShowExcelContextMenu += (sheet, pos) => ShowExcelContextMenu?.Invoke(sheet, pos);
            
            // 拖拽物
            _list.SheetDragDropRequested += (dragged, target, before) => SheetDragDropRequested?.Invoke(dragged, target, before);
            _list.GroupDragDropRequested += (dragged, target, before) => GroupDragDropRequested?.Invoke(dragged, target, before);

            // 输入变化时自动计算（仅在被计算目标未被编辑的情况下）
            _txtT.TextChanged += (s, e) => AutoCalculateIfReady(changedField: "T");
            _txtN.TextChanged += (s, e) => AutoCalculateIfReady(changedField: "N");
            _txtP.TextChanged += (s, e) => AutoCalculateIfReady(changedField: "P");
            _rbT.CheckedChanged += (s, e) => AutoCalculateIfReady(changedField: null);
            _rbN.CheckedChanged += (s, e) => AutoCalculateIfReady(changedField: null);
            _rbP.CheckedChanged += (s, e) => AutoCalculateIfReady(changedField: null);

            // 装配布局
            panel.Controls.Add(_search, 0, 0);
            panel.Controls.Add(_list, 0, 1);
            panel.Controls.Add(_buttonPanel, 0, 2);
            Controls.Add(panel);

            // 初始化按钮文本与位置（靠近屏幕中心方向）
            SetDockIndicator(true);

            // 初始化计算默认值与默认目标：N=1000 rpm, T=5 N·m, 计算 P
            InitializeDefaultCalculationState();
        }

        /// <summary>
        /// 供宿主更新工具按钮的状态显示。
        /// </summary>
        public void SetToolStatus(bool solverEnabled, bool analysisEnabled)
        {
            try
            {
                if (solverEnabled && analysisEnabled)
                {
                    _btnTools.Text = "ANSYS:ON";
                    _btnTools.BackColor = Color.LightGreen;
                }
                else if (!solverEnabled && !analysisEnabled)
                {
                    _btnTools.Text = "ANSYS:OFF";
                    _btnTools.BackColor = Color.LightCoral; // 红色系
                }
                else
                {
                    // 其中一个未开启，显示具体未开启的那个
                    _btnTools.Text = solverEnabled ? "Data:OFF" : "Solver:OFF";
                    _btnTools.BackColor = Color.Khaki; // 黄色系
                }
            }
            catch { }
        }

        /// <summary>
        /// 由宿主调用以在 UI 上指示当前停靠侧（左侧显示 "->"；右侧显示 "<-")，并将按钮靠近屏幕中心显示。
        /// </summary>
        public void SetDockIndicator(bool isLeftDock)
        {
            _isLeftDock = isLeftDock;
            _btnDockToggle.Text = isLeftDock ? "->" : "<-";   // 指向中心方向的箭头
            
            // 根据停靠侧调整布局方向
            if (isLeftDock)
            {
                // 左侧停靠：统计在左，按钮在右
                _buttonPanel.FlowDirection = FlowDirection.LeftToRight;
            }
            else
            {
                // 右侧停靠：按钮在左，统计在右
                _buttonPanel.FlowDirection = FlowDirection.RightToLeft;
            }
            _buttonPanel.Padding = new Padding(0);
        }

        /// <summary>
        /// 由宿主传入工作表名列表、总数、隐藏数以初始化或刷新显示。
        /// </summary>
        public void LoadSheets(List<string> visibleNames, int totalCount, int hiddenCount)
        {
            _all = visibleNames ?? new List<string>();
            _totalSheets = totalCount;
            _hiddenSheets = hiddenCount;

            // 如果当前有有效的搜索关键字，则优先按关键字过滤；
            // 否则显示全部。这样在过滤状态下，刷新不会自动恢复为全部列表。
            var term = string.Empty;
            if (_search != null && _search.ForeColor != Color.Gray)
            {
                term = _search.Text;
            }

            var filtered = string.IsNullOrWhiteSpace(term)
                ? _all
                : _all.FindAll(x => x.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0);

            _matchSheets = filtered.Count;
            _list.LoadSheets(filtered);
            UpdateStatsLabel();
        }

        /// <summary>
        /// 更新统计标签显示
        /// </summary>
        private void UpdateStatsLabel()
        {
            if (_search != null && _search.ForeColor != System.Drawing.Color.Gray && !string.IsNullOrWhiteSpace(_search.Text))
            {
                _statsLabel.Text = string.Format("Total: {0} | Hidden: {1} | Match: {2}", _totalSheets, _hiddenSheets, _matchSheets);
            }
            else
            {
                _statsLabel.Text = string.Format("Total: {0} | Hidden: {1}", _totalSheets, _hiddenSheets);
            }
        }

        /// <summary>
        /// 高亮显示指定工作表名（用于与 Excel 的活动工作表同步）。
        /// </summary>
        public void Highlight(string name)
        {
            _list.Highlight(name);
        }

        /// <summary>
        /// 将分组信息应用到列表（包含颜色与成员信息）。
        /// </summary>
        public void ApplyGroups(Dictionary<string, (List<string> Members, System.Drawing.Color Color)> groups)
        {
            _list.ApplyGroups(groups);
        }

        /// <summary>
        /// 由宿主触发，开始对某个工作表进行内联重命名。
        /// </summary>
        public void BeginInlineRenameSheet(string sheetName)
        {
            _list.BeginInlineRenameSheet(sheetName);
        }

        /// <summary>
        /// 由宿主触发，开始对某个分组进行内联重命名。
        /// </summary>
        public void BeginInlineRenameGroup(string groupName)
        {
            _list.BeginInlineRenameGroup(groupName);
        }

        /// <summary>
        /// 根据搜索关键字过滤并显示匹配的工作表名（不区分大小写）。
        /// </summary>
        private void Filter(string term)
        {
            var filtered = string.IsNullOrWhiteSpace(term)
                ? _all
                : _all.FindAll(x => x.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0);
            _matchSheets = filtered.Count;
            _list.LoadSheets(filtered);
            UpdateStatsLabel();
        }

        // 宿主用于应用工作表颜色映射
        public void ApplySheetColors(Dictionary<string, Color> colors)
        {
            try { _list?.ApplySheetColors(colors); } catch { }
        }

        // 宿主获取当前选择
        public List<string> GetSelectedNames()
        {
            try { return _list?.GetSelectedNames() ?? new List<string>(); } catch { return new List<string>(); }
        }

        // === 计算逻辑 ===
        private void CalculateTNPorShowError()
        {
            try
            {
                // 解析输入：支持 P 写成 "200" 或 "0.2kW"；其余按数值解析
                bool wantT = _rbT.Checked;
                bool wantN = _rbN.Checked;
                bool wantP = _rbP.Checked;

                if (!wantT && !wantN && !wantP)
                {
                    MessageBox.Show("Please select which value to calculate: T (torque) / N (speed) / P (power).","Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 读取并规范单位
                double? tNm = ParseNumber(_txtT.Text);          // N·m
                double? nRpm = ParseNumber(_txtN.Text);         // rpm
                double? pW = ParsePowerW(_txtP.Text);           // W

                // 根据选择的目标量，要求其余两个必须有值
                if (wantP)
                {
                    if (tNm == null || nRpm == null) { ShowInputError("Power P is calculated from torque T (N·m) and speed N (rpm)."); return; }
                    double p = (tNm.Value * 2.0 * Math.PI * nRpm.Value) / 60.0; // W
                    _txtP.Text = Math.Round(p, 6).ToString();
                }
                else if (wantN)
                {
                    if (tNm == null || pW == null) { ShowInputError("Speed N is calculated from torque T (N·m)  and power P (W/kW)"); return; }
                    if (tNm.Value == 0) { ShowInputError("T must not be 0."); return; }
                    double n = (60.0 * pW.Value) / (2.0 * Math.PI * tNm.Value); // rpm
                    _txtN.Text = Math.Round(n, 6).ToString();
                }
                else if (wantT)
                {
                    if (nRpm == null || pW == null) { ShowInputError("Torque T is calculated from speed N (rpm) and power P (W/kW)"); return; }
                    if (nRpm.Value == 0) { ShowInputError("N must not be 0."); return; }
                    double t = (60.0 * pW.Value) / (2.0 * Math.PI * nRpm.Value); // N·m
                    _txtT.Text = Math.Round(t, 6).ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Calculation error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 自动计算：在输入变化时，根据当前选择的目标量进行计算；若参与计算的输入为空，则将目标输出为 "Err."。
        private void AutoCalculateIfReady(string changedField)
        {
            if (_suppressAutoCalc) return;

            bool wantT = _rbT.Checked;
            bool wantN = _rbN.Checked;
            bool wantP = _rbP.Checked;

            // 被计算的值发生改变时，不触发计算
            if (changedField == "T" && wantT) return;
            if (changedField == "N" && wantN) return;
            if (changedField == "P" && wantP) return;

            try
            {
                _suppressAutoCalc = true;

                double? tNm = ParseNumber(_txtT.Text);
                double? nRpm = ParseNumber(_txtN.Text);
                double? pW = ParsePowerW(_txtP.Text);

                if (wantP)
                {
                    if (IsEmpty(_txtT.Text) || IsEmpty(_txtN.Text) || tNm == null || nRpm == null)
                    {
                        _txtP.Text = "Err.";
                    }
                    else
                    {
                        double p = (tNm.Value * 2.0 * Math.PI * nRpm.Value) / 60.0; // W
                        _txtP.Text = Math.Round(p, 6).ToString();
                    }
                }
                else if (wantN)
                {
                    if (IsEmpty(_txtT.Text) || IsEmpty(_txtP.Text) || tNm == null || pW == null)
                    {
                        _txtN.Text = "Err.";
                    }
                    else if (tNm.Value == 0)
                    {
                        _txtN.Text = "Err.";
                    }
                    else
                    {
                        double n = (60.0 * pW.Value) / (2.0 * Math.PI * tNm.Value); // rpm
                        _txtN.Text = Math.Round(n, 6).ToString();
                    }
                }
                else if (wantT)
                {
                    if (IsEmpty(_txtN.Text) || IsEmpty(_txtP.Text) || nRpm == null || pW == null)
                    {
                        _txtT.Text = "Err.";
                    }
                    else if (nRpm.Value == 0)
                    {
                        _txtT.Text = "Err.";
                    }
                    else
                    {
                        double t = (60.0 * pW.Value) / (2.0 * Math.PI * nRpm.Value); // N·m
                        _txtT.Text = Math.Round(t, 6).ToString();
                    }
                }
            }
            finally
            {
                _suppressAutoCalc = false;
            }
        }

        private static bool IsEmpty(string s)
        {
            return string.IsNullOrWhiteSpace(s);
        }

        private void InitializeDefaultCalculationState()
        {
            // 默认：N=1000 rpm, T=5 N·m，计算 P
            _suppressAutoCalc = true;
            _rbP.Checked = true;
            _rbN.Checked = false;
            _rbT.Checked = false;
            _txtN.Text = "1000"; // rpm
            _txtT.Text = "5";    // N·m
            _txtP.Text = string.Empty;
            _suppressAutoCalc = false;
            // 触发一次计算，得到初始 P
            AutoCalculateIfReady(changedField: null);
        }

        private static double? ParseNumber(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return null;
            var original = s.Trim();
            var lower = original.ToLowerInvariant();
            // 去掉空格与中点便于匹配
            var compact = lower.Replace(" ", string.Empty).Replace("·", string.Empty);

            double multiplier = 1.0;

            // 扭矩：mNm (milli-Newton meter) -> N·m
            if (compact.Contains("mnm") || compact.Contains("mnm"))
            {
                multiplier *= 1e-3; // mNm -> N·m
            }

            // 转速：krpm -> rpm（支持“2krpm”或“2k rpm”或“2kr/min”）
            if (compact.Contains("krpm") || compact.Contains("kr/min") || (compact.Contains("k") && (compact.Contains("rpm") || compact.Contains("r/min") || compact.Contains("rmin"))))
            {
                multiplier *= 1000.0;
            }

            // 清理单位标识
            string cleaned = lower
                .Replace("mn·m", string.Empty)
                .Replace("m n·m", string.Empty)
                .Replace("mnm", string.Empty)
                .Replace("n·m", string.Empty)
                .Replace("nm", string.Empty)
                .Replace("krpm", string.Empty)
                .Replace("k rpm", string.Empty)
                .Replace("k r/min", string.Empty)
                .Replace("k rmin", string.Empty)
                .Replace("rpm", string.Empty)
                .Replace("r/min", string.Empty)
                .Replace("rmin", string.Empty)
                .Replace("·", string.Empty)
                .Replace(" ", string.Empty)
                .Trim();

            // 如果还有单独的 k（例如 "2k"），在用于转速时也当作 *1000 处理
            if (cleaned.EndsWith("k"))
            {
                multiplier *= 1000.0;
                cleaned = cleaned.Substring(0, cleaned.Length - 1);
            }

            if (double.TryParse(cleaned, out var v))
                return v * multiplier;
            return null;
        }
        private static double? ParsePowerW(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return null;
            s = s.Trim();
            bool isKW = s.IndexOf("kw", StringComparison.OrdinalIgnoreCase) >= 0;
            var cleaned = s.Replace("kW", "").Replace("KW", "").Replace("kw", "").Replace("W", "").Trim();
            if (!double.TryParse(cleaned, out var v)) return null;
            return isKW ? v * 1000.0 : v; // 统一换算为瓦特
        }
        private static void ShowInputError(string msg)
        {
            MessageBox.Show(msg, "ErrorInput", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
}
