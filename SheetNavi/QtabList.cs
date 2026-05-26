using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

/*文件编码：UTF-8 无BOM*/
namespace Qtab
{
    /// <summary>
    /// `SheetTabList` 是一个自定义的 `UserControl`，用于在任务面板中显示 Excel 工作表列表与分组，
    /// 支持选择、分组/取消分组、设置颜色、以及工作表与分组的内联重命名。
    /// 本控件只负责 UI 展示与交互，具体对 Excel 的操作通过事件回调交给宿主（`ThisAddIn`）处理。
    /// </summary>
    public class QtabList : UserControl
    {
        // ============ 宿主（ThisAddIn / QiTab）订阅的事件 ============
        public event Action<string> OnSheetSelected;                 // 当用户请求激活某个工作表时触发，参数为工作表名
        public event Action<List<string>> GroupRequested;            // 请求将当前选中工作表分组
        public event Action<List<string>> UngroupRequested;          // 请求将当前选中工作表取消分组
        public event Action<List<string>> RenameGroupRequested;      // 请求重命名分组（通常由右键菜单触发，宿主可弹出输入或启用内联）
        public event Action<List<string>, Color> ColorRequested;     // 请求设置所选工作表的标签颜色
        public event Action<string> RenameSheetRequested;            // 请求开始工作表重命名（例如双击），宿主负责实际改名或启动内联
        public event Action<string, string> GroupRenameCommitted;    // 分组内联重命名已提交（参数：旧名，新名）
        public event Action<string, string> RenameSheetCommitted;    // 工作表内联重命名已提交（参数：旧名，新名）
        public event Action<List<string>> DeleteSheetRequested;      // 新增：请求删除选中工作表
        public event Action<string, Point> ShowExcelContextMenu;     // 新增：请求显示 Excel 原生右键菜单（工作表名，屏幕坐标）

        /// <summary>
        /// 分组数据模型：名称、颜色、是否折叠、成员工作表列表
        /// </summary>
        private class GroupModel
        {
            public string Name;
            public Color Color;
            public bool Collapsed;
            public List<string> Members = new List<string>();
        }

        // 所有分组（按名称索引，忽略大小写）
        private Dictionary<string, GroupModel> _groups = new Dictionary<string, GroupModel>(StringComparer.OrdinalIgnoreCase);
        // 分组的显示顺序（只保存分组名）
        private List<string> _groupOrder = new List<string>();
        // 所有工作表名（由宿主传入）
        private List<string> _allSheets = new List<string>();

        // 当前被选中的工作表集合（支持多选），用于高亮显示与命令作用范围
        private HashSet<string> _selected = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        // 最近一次选择的工作表名（便于 Shift 多选）
        private string _lastSelected;
        // 悬停状态（不改变选择，仅用于绘制反馈）
        private string _hoverSheet;
        private string _hoverGroup;

        // 绘制用字体（工作表行、分组行）
        private Font _sheetFont;
        private Font _groupFont;

        // 点击计时器，用于区分单/双击或快速连续点击的交互（例如触发分组重命名）
        private Timer _clickTimer = new Timer();
        private int _clickCount = 0;
        private Rectangle _lastClickRect;
        private string _lastClickGroup;

        // 记录最近一次右键命中的目标（用于右键菜单操作）
        private HitInfo _lastContextHit;

        // 内联编辑相关控件与状态
        private TextBox _editBox;           // 覆盖在文本上的编辑框
        private Rectangle _editRect;        // 编辑框位置/大小
        private string _editOriginal;       // 编辑前的原始名称
        private bool _editIsGroup;          // 编辑目标是否为分组（否则为工作表）

        // 拖拽相关字段
        private bool _isDragging = false;
        private HitInfo _dragStartHit;      // 拖拽开始时命中的项
        private Point _dragStartPoint;      // 拖拽开始位置
        private HitInfo _dragOverHit;       // 当前拖拽悬停的项
        private const int DragThreshold = 5; // 拖拽触发阈值（像素）
        
        // 拖拽事件
        public event Action<string, string, bool> SheetDragDropRequested;  // (draggedSheet, targetSheet, insertBefore)
        public event Action<string, string, bool> GroupDragDropRequested;  // (draggedGroup, targetItem, insertBefore)

        // 每个工作表的标签颜色（未分组时显示；有分组时由分组颜色覆盖）
        private Dictionary<string, Color> _sheetColors = new Dictionary<string, Color>(StringComparer.OrdinalIgnoreCase);

        public QtabList()
        {
            // 开启双缓冲以避免闪烁
            DoubleBuffered = true;
            BackColor = SystemColors.Window;
            _sheetFont = new Font(SystemFonts.MessageBoxFont, FontStyle.Regular);
            _groupFont = new Font(SystemFonts.MessageBoxFont, FontStyle.Bold);
            
            // 启用拖放
            AllowDrop = true;

            // ============ 构建右键菜单（移除，改由宿主处理）============
            // （保留接口：GroupRequested、UngroupRequested、RenameGroupRequested、ColorRequested、DeleteSheetRequested、ShowExcelContextMenu）

            // 点击计时器：用于检测快速连续点击（用于分组内联重命名的触发逻辑）
            _clickTimer.Interval = 300;
            _clickTimer.Tick += (s, e) => { _clickCount = 0; _clickTimer.Stop(); _lastClickGroup = null; };

            // 订阅鼠标事件与重绘
            MouseDown += OnMouseDown;
            MouseMove += OnMouseMove;
            MouseUp += OnMouseUp;
            MouseDoubleClick += OnMouseDoubleClick;
            MouseLeave += (s, e) =>
            {
                if (!string.IsNullOrEmpty(_hoverSheet) || !string.IsNullOrEmpty(_hoverGroup))
                {
                    _hoverSheet = null;
                    _hoverGroup = null;
                    Cursor = Cursors.Default;
                    Invalidate();
                }
            };
            Resize += (s, e) => Invalidate();
            
            // 拖拽事件
            DragEnter += (s, e) => Invalidate();
            DragOver += OnDragOver;
            DragDrop += OnDragDrop;

            // 内联编辑框：失焦或按 Enter 提交，Esc 取消
            _editBox = new TextBox { Visible = false };
            _editBox.Leave += (s, e) => CommitEdit(true);
            _editBox.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter) { CommitEdit(true); e.Handled = true; }
                else if (e.KeyCode == Keys.Escape) { CommitEdit(false); e.Handled = true; }
            };
            Controls.Add(_editBox);
        }

        /// <summary>
        /// 由宿主传入工作表名列表，控件保存并重绘。
        /// </summary>
        public void LoadSheets(List<string> names)
        {
            _allSheets = names ?? new List<string>();
            Invalidate();
        }

        /// <summary>
        /// 由宿主传入分组信息（字典：分组名 -> 成员与颜色），控件更新内部模型并重绘。
        /// </summary>
        public void ApplyGroups(Dictionary<string, (List<string> Members, Color Color)> groups)
        {
            _groups.Clear();
            _groupOrder.Clear();
            if (groups != null)
            {
                foreach (var kv in groups)
                {
                    var gm = new GroupModel
                    {
                        Name = kv.Key,
                        Color = kv.Value.Color,
                        Collapsed = false,
                        Members = kv.Value.Members?.ToList() ?? new List<string>()
                    };
                    _groups[kv.Key] = gm;
                    _groupOrder.Add(kv.Key);
                }
            }
            Invalidate();
        }

        /// <summary>
        /// 宿主传入每个工作表的颜色映射。
        /// </summary>
        public void ApplySheetColors(Dictionary<string, Color> colors)
        {
            _sheetColors = colors ?? new Dictionary<string, Color>(StringComparer.OrdinalIgnoreCase);
            Invalidate();
        }

        /// <summary>
        /// 高亮显示指定工作表（单选），用于与 Excel 活动工作表同步显示。
        /// </summary>
        public void Highlight(string name)
        {
            _selected.Clear();
            if (!string.IsNullOrEmpty(name)) _selected.Add(name);
            _lastSelected = name;
            Invalidate();
        }

        /// <summary>
        /// 获取当前被选中的工作表名列表（用于分组/取消分组/上移/下移/着色等操作）。
        /// </summary>
        public List<string> GetSelectedNames() => _selected.ToList();

        /// <summary>
        /// 自定义绘制：绘制分组头、分组成员与未分组工作表。包含选中高亮与分组颜色提示。
        /// </summary>
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            var g = e.Graphics;
            g.Clear(BackColor);

            int y = 4;                         // 当前绘制的垂直偏移
            int arrowSize = 10;                // 分组折叠箭头大小
            var sheetHeight = Math.Max(_sheetFont.Height + 8, 22);   // 工作表行高度
            var groupHeight = Math.Max(_groupFont.Height + 8, 24);   // 分组头行高度

            // 新逻辑：按工作表在 Excel 中的原始顺序绘制，当遇到某分组的第一个成员时插入分组头及其所有成员
            // 1. 构建每个工作表所属分组的映射
            var sheetToGroup = new Dictionary<string, GroupModel>(StringComparer.OrdinalIgnoreCase);
            foreach (var gm in _groups.Values)
            {
                foreach (var member in gm.Members)
                {
                    sheetToGroup[member] = gm;
                }
            }

            // 2. 记录已绘制的分组，避免重复绘制
            var drawnGroups = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // 3. 按 _allSheets 顺序遍历（这是 Excel 工作簿中的原始顺序）
            foreach (var sheetName in _allSheets)
            {
                // 检查该工作表是否属于某个分组
                if (sheetToGroup.TryGetValue(sheetName, out var gm))
                {
                    // 如果该分组尚未绘制，则在此处绘制分组头及其所有成员
                    if (!drawnGroups.Contains(gm.Name))
                    {
                        drawnGroups.Add(gm.Name);

                        // 绘制分组头
                        var headerRect = new Rectangle(0, y, Width, groupHeight);
                        DrawGroupHeader(g, gm, headerRect, arrowSize);
                        y += groupHeight;

                        // 如果未折叠，绘制该组所有成员（按 Excel 原始顺序排序）
                        if (!gm.Collapsed)
                        {
                            // **修复**：按 _allSheets 的顺序绘制分组成员，而不是按 gm.Members 的顺序
                            // 这样可以确保分组内的工作表按照它们在 Excel 中的实际位置顺序显示
                            foreach (var member in _allSheets)
                            {
                                if (!gm.Members.Contains(member)) continue; // 不是该组成员则跳过
                                var rect = new Rectangle(0, y, Width, sheetHeight);
                                DrawSheetItem(g, member, rect, gm.Color);
                                y += sheetHeight;
                            }
                        }
                    }
                    // 如果该分组已绘制，则当前工作表已在分组内绘制过，跳过
                }
                else
                {
                    // 该工作表未分组，直接绘制
                    var rect = new Rectangle(0, y, Width, sheetHeight);
                    Color own = Color.Empty;
                    if (_sheetColors != null && _sheetColors.TryGetValue(sheetName, out var c)) own = c;
                    DrawSheetItem(g, sheetName, rect, own);
                    y += sheetHeight;
                }
            }

            // 绘制拖拽插入位置指示线
            if (_dragOverHit != null && _isDragging)
            {
                DrawDropIndicator(e.Graphics, _dragOverHit);
            }
        }

        /// <summary>
        /// 绘制分组头：可显示分组颜色背景，左侧折叠/展开标识，右侧绘制分组名称文本。
        /// </summary>
        private void DrawGroupHeader(Graphics g, GroupModel gm, Rectangle rect, int arrowSize)
        {
            // 分组颜色背景（半透明）
            if (gm.Color != Color.Empty)
            {
                using (var b = new SolidBrush(Color.FromArgb(200, gm.Color))) g.FillRectangle(b, rect);
            }

            // 悬停高亮（不覆盖分组色，仅做轻量叠加）
            if (!string.IsNullOrEmpty(_hoverGroup) && string.Equals(_hoverGroup, gm.Name, StringComparison.OrdinalIgnoreCase))
            {
                using (var b = new SolidBrush(Color.FromArgb(26, SystemColors.HotTrack))) g.FillRectangle(b, rect);
            }

            // 折叠/展开标识区域
            var arrowRect = new Rectangle(rect.X + 6, rect.Y + (rect.Height - arrowSize) / 2, arrowSize, arrowSize);
            using (var p = new Pen(SystemColors.ControlDark, 1))
            {
                g.DrawRectangle(p, arrowRect);
                using (var p2 = new Pen(SystemColors.ControlDarkDark, 2))
                {
                    g.DrawLine(p2, arrowRect.Left + 2, arrowRect.Top + arrowRect.Height / 2, arrowRect.Right - 2, arrowRect.Top + arrowRect.Height / 2);
                    if (gm.Collapsed)
                    {
                        g.DrawLine(p2, arrowRect.Left + arrowRect.Width / 2, arrowRect.Top + 2, arrowRect.Left + arrowRect.Width / 2, arrowRect.Bottom - 2);
                    }
                }
            }
            var text = gm.Name;
            var textRect = new Rectangle(arrowRect.Right + 6, rect.Y, rect.Width - (arrowRect.Right + 6), rect.Height);
            TextRenderer.DrawText(g, text, _groupFont, textRect, SystemColors.ControlText, TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.EndEllipsis);
        }

        /// <summary>
        /// 绘制工作表条目：包含选中高亮、分组颜色的淡背景与左侧颜色条。
        /// </summary>
        private void DrawSheetItem(Graphics g, string name, Rectangle rect, Color color)
        {
            bool sel = _selected.Contains(name);
            if (sel)
            {
                using (var b = new SolidBrush(SystemColors.ActiveCaptionText)) g.FillRectangle(b, rect);
            }
            else if (!string.IsNullOrEmpty(_hoverSheet) && string.Equals(_hoverSheet, name, StringComparison.OrdinalIgnoreCase))
            {
                using (var b = new SolidBrush(Color.FromArgb(24, SystemColors.Highlight))) g.FillRectangle(b, rect);
            }
            else if (color != Color.Empty)
            {
                using (var b = new SolidBrush(Color.FromArgb(24, color))) g.FillRectangle(b, rect);
                using (var b = new SolidBrush(color)) g.FillRectangle(b, new Rectangle(rect.X + 8, rect.Y + 4, 4, rect.Height - 8));
            }
            var fore = sel ? SystemColors.HighlightText : SystemColors.ControlText;
            var textRect = new Rectangle(rect.X + 24, rect.Y, rect.Width - 24, rect.Height);
            TextRenderer.DrawText(g, name, _sheetFont, textRect, fore, TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.EndEllipsis);
        }

        /// <summary>
        /// 鼠标按下：处理分组的折叠/展开切换，并记录拖拽起始信息。
        /// </summary>
        private void OnMouseDown(object sender, MouseEventArgs e)
        {
            Focus();

            // If an inline edit is active and the user clicked outside the edit box, end editing.
            if (_editBox?.Visible == true)
            {
                // If click is inside the edit rectangle, forward focus to the edit box and do not process as normal click.
                if (_editRect.Contains(e.Location))
                {
                    try { _editBox.Focus(); } catch { }
                    return;
                }
                else
                {
                    // Commit edit when clicking elsewhere
                    try { CommitEdit(true); } catch { }
                    Invalidate();
                    return;
                }
            }

            var hit = HitTest(e.Location);
            if (hit == null) return;
            
            if (hit.Type == HitType.GroupArrow)
            {
                var gm = hit.Group;
                gm.Collapsed = !gm.Collapsed;
                Invalidate();
                return;
            }

            // 左键：记录拖拽起始信息（用于工作表和分组头）
            if (e.Button == MouseButtons.Left && (hit.Type == HitType.Sheet || hit.Type == HitType.GroupHeader))
            {
                _dragStartHit = hit;
                _dragStartPoint = e.Location;
                _isDragging = false;
            }
        }

        /// <summary>
        /// 鼠标松开：处理右键菜单显示，处理选择行为以及分组的快速点击（触发重命名）。
        /// </summary>
        private void OnMouseUp(object sender, MouseEventArgs e)
        {
            var hit = HitTest(e.Location);
            _lastContextHit = hit; // 供右键菜单操作使用
            if (hit == null) return;

            if (e.Button == MouseButtons.Right)
            {
                // 右键：如果命中工作表则将其加入选中；命中分组则选中整个分组成员
                if (hit.Type == HitType.Sheet)
                {
                    if (!_selected.Contains(hit.Sheet)) _selected.Add(hit.Sheet);
                }
                else if (hit.Type == HitType.GroupHeader)
                {
                    if (hit.Group.Members.Count > 0)
                    {
                        _selected.Clear(); // 右键分组时仅选择该组成员，避免误触导致编辑分组名
                        foreach (var s in hit.Group.Members) _selected.Add(s);
                    }
                }
                Invalidate();

                try
                {
                    var screenPos = System.Windows.Forms.Cursor.Position;
                    var targetSheet = hit.Type == HitType.Sheet ? hit.Sheet : _selected.FirstOrDefault();
                    if (!string.IsNullOrEmpty(targetSheet))
                    {
                        ShowExcelContextMenu?.Invoke(targetSheet, screenPos);
                    }
                }
                catch { }
                return;
            }

            // 左键：支持 Ctrl 多选、Shift 范围选；否则单选并触发 OnSheetSelected。
            bool ctrl = (ModifierKeys & Keys.Control) != 0;
            bool shift = (ModifierKeys & Keys.Shift) != 0;
            if (hit.Type == HitType.Sheet)
            {
                if (ctrl)
                {
                    if (_selected.Contains(hit.Sheet)) _selected.Remove(hit.Sheet); else _selected.Add(hit.Sheet);
                }
                else if (shift && !string.IsNullOrEmpty(_lastSelected))
                {
                    SelectRange(_lastSelected, hit.Sheet);
                }
                else
                {
                    _selected.Clear();
                    _selected.Add(hit.Sheet);
                    _lastSelected = hit.Sheet;
                    OnSheetSelected?.Invoke(hit.Sheet);
                }
                Invalidate();
            }
            else if (hit.Type == HitType.GroupHeader)
            {
                // 点击分组头：选中整个分组成员，便于后续的分组/移动操作
                _selected.Clear();
                foreach (var s in hit.Group.Members) _selected.Add(s);
                _lastSelected = hit.Group.Members.FirstOrDefault();
                if (!string.IsNullOrEmpty(_lastSelected)) OnSheetSelected?.Invoke(_lastSelected);
                Invalidate();

                // 如果在相同位置快速连续点击分组头，则触发一次“重命名”意图（由上层处理内联重命名）
                if (_lastClickGroup == hit.Group.Name && _clickCount == 1 && _lastClickRect.Contains(e.Location))
                {
                    GroupRenameCommitted?.Invoke(hit.Group.Name, hit.Group.Name); // 触发一次事件，宿主将启用内联重命名
                    _clickCount = 0;
                }
                else
                {
                    _clickCount = 1;
                    _lastClickRect = new Rectangle(hit.Bounds.Location, hit.Bounds.Size);
                    _lastClickGroup = hit.Group.Name;
                    _clickTimer.Stop();
                    _clickTimer.Start();
                }
    

            }
        }

        /// <summary>
        /// 鼠标双击：双击触发重命名请求（由宿主处理或开始内联编辑）。
        /// </summary>
        private void OnMouseDoubleClick(object sender, MouseEventArgs e)
        {
            var hit = HitTest(e.Location);
            if (hit == null) return;
            // 双击组名或表名直接开始内联重命名（由控件自身管理编辑框）
            if (hit.Type == HitType.GroupHeader)
            {
                // 开始分组内联重命名
                BeginInlineRenameGroup(hit.Group.Name);
            }
            else if (hit.Type == HitType.Sheet)
            {
                // 开始工作表内联重命名
                BeginInlineRenameSheet(hit.Sheet);
            }
        }

        /// <summary>
        /// 在可见顺序中选择范围（用于 Shift 多选）。
        /// </summary>
        private void SelectRange(string a, string b)
        {
            var order = GetVisibleSheetsInOrder();
            int ia = order.IndexOf(a);
            int ib = order.IndexOf(b);
            if (ia < 0 || ib < 0) return;
            if (ia > ib) { var t = ia; ia = ib; ib = t; }
            _selected.Clear();
            for (int i = ia; i <= ib; i++) _selected.Add(order[i]);
        }

        /// <summary>
        /// 计算当前界面上工作表的可见顺序（按分组展开顺序 + 未分组）。
        /// </summary>
        private List<string> GetVisibleSheetsInOrder()
        {
            var list = new List<string>();
            foreach (var groupName in _groupOrder)
            {
                var gm = _groups[groupName];
                if (!gm.Collapsed)
                {
                    foreach (var s in gm.Members)
                    {
                        if (_allSheets.Contains(s)) list.Add(s);
                    }
                }
            }
            foreach (var s in _allSheets)
            {
                if (!_groups.Values.Any(g => g.Members.Contains(s))) list.Add(s);
            }
            return list;
        }

        // ============ 命中测试（用于交互判定）===========
        private enum HitType { None, GroupHeader, GroupArrow, Sheet }
        private class HitInfo
        {
            public HitType Type;
            public GroupModel Group;
            public string Sheet;
            public Rectangle Bounds;
        }

        /// <summary>
        /// 根据鼠标位置判断命中元素（分组头、折叠箭头、工作表行）。
        /// </summary>
        private HitInfo HitTest(Point pt)
        {
            int y = 4;
            int arrowSize = 10;
            var sheetHeight = Math.Max(_sheetFont.Height + 8, 22);
            var groupHeight = Math.Max(_groupFont.Height + 8, 24);

            // 使用与 OnPaint 相同的逻辑计算布局
            var sheetToGroup = new Dictionary<string, GroupModel>(StringComparer.OrdinalIgnoreCase);
            foreach (var gm in _groups.Values)
            {
                foreach (var member in gm.Members)
                {
                    sheetToGroup[member] = gm;
                }
            }

            var drawnGroups = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var sheetName in _allSheets)
            {
                if (sheetToGroup.TryGetValue(sheetName, out var gm))
                {
                    if (!drawnGroups.Contains(gm.Name))
                    {
                        drawnGroups.Add(gm.Name);

                        var headerRect = new Rectangle(0, y, Width, groupHeight);
                        var arrowRect = new Rectangle(headerRect.X + 6, headerRect.Y + (headerRect.Height - arrowSize) / 2, arrowSize, arrowSize);
                        if (arrowRect.Contains(pt)) return new HitInfo { Type = HitType.GroupArrow, Group = gm, Bounds = arrowRect };
                        if (headerRect.Contains(pt)) return new HitInfo { Type = HitType.GroupHeader, Group = gm, Bounds = headerRect };
                        y += groupHeight;

                        if (!gm.Collapsed)
                        {
                            // **修复**：按 _allSheets 的顺序遍历分组成员，与 OnPaint 保持一致
                            foreach (var member in _allSheets)
                            {
                                if (!gm.Members.Contains(member)) continue;
                                var rect = new Rectangle(0, y, Width, sheetHeight);
                                if (rect.Contains(pt)) return new HitInfo { Type = HitType.Sheet, Group = gm, Sheet = member, Bounds = rect };
                                y += sheetHeight;
                            }
                        }
                    }
                }
                else
                {
                    var rect = new Rectangle(0, y, Width, sheetHeight);
                    if (rect.Contains(pt)) return new HitInfo { Type = HitType.Sheet, Sheet = sheetName, Bounds = rect };
                    y += sheetHeight;
                }
            }

            return null;
        }

        // ============ 内联重命名：由上层触发，控件显示编辑框并在提交后通知宿主 ============

        /// <summary>
        /// 开始工作表的内联重命名：在对应条目上覆盖一个 `TextBox` 供用户编辑。
        /// </summary>
        public void BeginInlineRenameSheet(string sheetName)
        {
            if (string.IsNullOrEmpty(sheetName)) return;
            var hit = FindSheetBounds(sheetName);
            if (hit == Rectangle.Empty) return;
            BeginEdit(hit, sheetName, isGroup: false);
        }

        /// <summary>
        /// 开始分组的内联重命名：在分组头上覆盖一个 `TextBox` 供用户编辑。
        /// </summary>
        public void BeginInlineRenameGroup(string groupName)
        {
            if (string.IsNullOrEmpty(groupName)) return;
            var hit = FindGroupHeaderBounds(groupName);
            if (hit == Rectangle.Empty) return;
            BeginEdit(hit, groupName, isGroup: true);
        }

        /// <summary>
        /// 内部方法：显示编辑框并初始化状态。
        /// </summary>
        private void BeginEdit(Rectangle rect, string original, bool isGroup)
        {
            _editIsGroup = isGroup;
            _editOriginal = original;

            // 计算文本显示区域，保持与 OnPaint 中的 textRect 对齐，
            // 使编辑框覆盖在名称文本上（而非整个行）以实现 overlay 效果。
            Rectangle textRect;
            if (isGroup)
            {
                // 与 DrawGroupHeader 中的布局保持一致
                int arrowSize = 10;
                var arrowRect = new Rectangle(rect.Left + 6, rect.Top + (rect.Height - arrowSize) / 2, arrowSize, arrowSize);
                textRect = new Rectangle(arrowRect.Right + 6, rect.Top, rect.Width - (arrowRect.Right + 6), rect.Height);
                _editBox.Font = _groupFont;
            }
            else
            {
                // 与 DrawSheetItem 中的布局保持一致
                textRect = new Rectangle(rect.Left + 24, rect.Top, rect.Width - 24, rect.Height);
                _editBox.Font = _sheetFont;
            }

            // 小范围收缩以避免编辑框与行边界重叠
            var bounds = new Rectangle(textRect.Left + 2, textRect.Top + 1, Math.Max(60, textRect.Width - 4), Math.Max(18, textRect.Height - 4));
            _editRect = bounds;
            _editBox.SetBounds(bounds.Left, bounds.Top, bounds.Width, bounds.Height);
            _editBox.Text = original;
            _editBox.Visible = true;
            _editBox.BringToFront();
            _editBox.Focus();
            _editBox.SelectAll();
        }

        /// <summary>
        /// 结束编辑：根据 `commit` 决定是否提交更名。提交后通过事件把新名称交给宿主处理。
        /// </summary>
        private void CommitEdit(bool commit)
        {
            if (!_editBox.Visible) return;
            var newName = _editBox.Text?.Trim() ?? string.Empty;
            var old = _editOriginal;
            _editBox.Visible = false;
            if (!commit || string.IsNullOrEmpty(newName) || newName == old) return;
            if (_editIsGroup)
                GroupRenameCommitted?.Invoke(old, newName);
            else
                RenameSheetCommitted?.Invoke(old, newName);
        }

        /// <summary>
        /// 查找指定工作表在控件中的近似矩形区域（用于定位编辑框）。
        /// </summary>
        private Rectangle FindSheetBounds(string sheetName)
        {
            // 使用与 OnPaint 和 HitTest 相同的布局算法来计算工作表条目的位置，
            // 确保内联编辑框覆盖在正确的行上。
            int y = 4;
            var sheetHeight = Math.Max(_sheetFont.Height + 8, 22);
            var groupHeight = Math.Max(_groupFont.Height + 8, 24);

            // 构建 sheet -> group 显示名称映射
            var sheetToGroup = new Dictionary<string, GroupModel>(StringComparer.OrdinalIgnoreCase);
            foreach (var gm in _groups.Values)
            {
                foreach (var member in gm.Members)
                {
                    sheetToGroup[member] = gm;
                }
            }

            var drawnGroups = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var s in _allSheets)
            {
                if (sheetToGroup.TryGetValue(s, out var gm))
                {
                    if (!drawnGroups.Contains(gm.Name))
                    {
                        // 绘制分组头
                        y += groupHeight;

                        if (!gm.Collapsed)
                        {
                            // 按 _allSheets 顺序绘制该组的成员，并检查目标工作表
                            foreach (var member in _allSheets)
                            {
                                if (!gm.Members.Contains(member)) continue;
                                var rect = new Rectangle(0, y, this.Width, sheetHeight);
                                if (string.Equals(member, sheetName, StringComparison.OrdinalIgnoreCase)) return rect;
                                y += sheetHeight;
                            }
                        }

                        drawnGroups.Add(gm.Name);
                    }
                    // 如果该分组已绘制，则当前成员已在上面处理过，跳过
                }
                else
                {
                    // 未分组的单行项
                    var rect = new Rectangle(0, y, this.Width, sheetHeight);
                    if (string.Equals(s, sheetName, StringComparison.OrdinalIgnoreCase)) return rect;
                    y += sheetHeight;
                }
            }

            // 回退：按旧的 groupOrder 计算（用于极端情况）
            y = 4;
            foreach (var g in _groupOrder)
            {
                var rect = new Rectangle(0, y, this.Width, groupHeight);
                y += groupHeight;
                if (_groups.TryGetValue(g, out var gm) && !gm.Collapsed)
                {
                    foreach (var member in gm.Members)
                    {
                        var mrect = new Rectangle(0, y, this.Width, sheetHeight);
                        if (string.Equals(member, sheetName, StringComparison.OrdinalIgnoreCase)) return mrect;
                        y += sheetHeight;
                    }
                }
            }

            return Rectangle.Empty;
        }

        /// <summary>
        /// 查找指定分组头在控件中的近似矩形区域（用于定位编辑框）。
        /// </summary>
        private Rectangle FindGroupHeaderBounds(string groupName)
        {
            // 使用与 OnPaint 和 HitTest 相同的布局算法来计算分组头的位置，
            // 这样编辑框可以准确覆盖在绘制的分组名称文本上。
            int y = 4;
            var sheetHeight = Math.Max(_sheetFont.Height + 8, 22);
            var groupHeight = Math.Max(_groupFont.Height + 8, 24);

            // 构建 sheet -> group 显示名称映射
            var sheetToGroup = new Dictionary<string, GroupModel>(StringComparer.OrdinalIgnoreCase);
            foreach (var gm in _groups.Values)
            {
                foreach (var member in gm.Members)
                {
                    sheetToGroup[member] = gm;
                }
            }

            var drawnGroups = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var sheetName in _allSheets)
            {
                if (sheetToGroup.TryGetValue(sheetName, out var gm))
                {
                    if (!drawnGroups.Contains(gm.Name))
                    {
                        // 这里即将绘制分组头，若是目标分组则返回其矩形
                        var headerRect = new Rectangle(0, y, this.Width, groupHeight);
                        if (string.Equals(gm.Name, groupName, StringComparison.OrdinalIgnoreCase))
                        {
                            return headerRect;
                        }

                        drawnGroups.Add(gm.Name);
                        y += groupHeight;

                        if (!gm.Collapsed)
                        {
                            // 绘制分组成员（按 _allSheets 顺序）
                            foreach (var member in _allSheets)
                            {
                                if (!gm.Members.Contains(member)) continue;
                                y += sheetHeight;
                            }
                        }
                    }
                }
                else
                {
                    // 未分组项直接占用一行
                    y += sheetHeight;
                }
            }

            // 如果分组在绘制逻辑中未找到（例如分组存在但没有成员显示在 _allSheets 中），
            // 回退到按 groupOrder 计算的位置（保持兼容旧逻辑）
            y = 4;
            foreach (var g in _groupOrder)
            {
                var rect = new Rectangle(0, y, this.Width, groupHeight);
                if (string.Equals(g, groupName, StringComparison.OrdinalIgnoreCase)) return rect;
                y += groupHeight;
                if (_groups.TryGetValue(g, out var gm) && !gm.Collapsed)
                {
                    y += sheetHeight * (gm.Members?.Count ?? 0);
                }
            }

            return Rectangle.Empty;
        }

        /// <summary>
        /// 鼠标移动：处理拖拽启动。
        /// </summary>
        private void OnMouseMove(object sender, MouseEventArgs e)
        {
            if (!_isDragging)
            {
                var hover = HitTest(e.Location);
                var newHoverSheet = hover != null && hover.Type == HitType.Sheet ? hover.Sheet : null;
                var newHoverGroup = hover != null && hover.Type == HitType.GroupHeader ? hover.Group?.Name : null;

                bool changed = !string.Equals(_hoverSheet, newHoverSheet, StringComparison.OrdinalIgnoreCase)
                               || !string.Equals(_hoverGroup, newHoverGroup, StringComparison.OrdinalIgnoreCase);
                if (changed)
                {
                    _hoverSheet = newHoverSheet;
                    _hoverGroup = newHoverGroup;
                    Invalidate();
                }
                Cursor = hover == null ? Cursors.Default : Cursors.Hand;
            }

            if (e.Button != MouseButtons.Left || _dragStartHit == null) return;
            if (_isDragging) return;

            // 检查是否超过拖拽阈值
            int dx = Math.Abs(e.X - _dragStartPoint.X);
            int dy = Math.Abs(e.Y - _dragStartPoint.Y);
            if (dx > DragThreshold || dy > DragThreshold)
            {
                _isDragging = true;

                // 根据拖拽的是分组头还是工作表，准备不同的数据
                object dragData = null;
                if (_dragStartHit.Type == HitType.GroupHeader)
                {
                    dragData = "GROUP:" + _dragStartHit.Group.Name;
                }
                else if (_dragStartHit.Type == HitType.Sheet)
                {
                    dragData = "SHEET:" + _dragStartHit.Sheet;
                }

                if (dragData != null)
                {
                    try
                    {
                        DoDragDrop(dragData, DragDropEffects.Move);
                    }
                    catch { }
                    finally
                    {
                        _isDragging = false;
                        _dragStartHit = null;
                        _dragOverHit = null;
                        Cursor = Cursors.Default;
                        Invalidate();
                    }
                }
            }
        }

        /// <summary>
        /// 拖拽悬停：显示拖放位置提示。
        /// </summary>
        private void OnDragOver(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(typeof(string)))
            {
                e.Effect = DragDropEffects.None;
                return;
            }

            var dragData = e.Data.GetData(typeof(string)) as string;
            if (string.IsNullOrEmpty(dragData))
            {
                e.Effect = DragDropEffects.None;
                return;
            }

            // 转换屏幕坐标到控件坐标
            var clientPoint = PointToClient(new Point(e.X, e.Y));
            var hit = HitTest(clientPoint);

            _dragOverHit = hit;

            if (hit != null && (hit.Type == HitType.Sheet || hit.Type == HitType.GroupHeader))
            {
                e.Effect = DragDropEffects.Move;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }

            Invalidate(); // 重绘以显示插入位置提示
        }

        /// <summary>
        /// 拖放完成：处理工作表或分组的重新排序。
        /// </summary>
        private void OnDragDrop(object sender, DragEventArgs e)
        {
            try
            {
                if (!e.Data.GetDataPresent(typeof(string))) return;

                var dragData = e.Data.GetData(typeof(string)) as string;
                if (string.IsNullOrEmpty(dragData)) return;

                var clientPoint = PointToClient(new Point(e.X, e.Y));
                var targetHit = HitTest(clientPoint);
                if (targetHit == null) return;

                // 解析拖拽数据
                bool isDraggingGroup = dragData.StartsWith("GROUP:");
                bool isDraggingSheet = dragData.StartsWith("SHEET:");
                var draggedName = dragData.Substring(dragData.IndexOf(':') + 1);

                // 计算插入位置（在目标之前还是之后）
                bool insertBefore = clientPoint.Y < (targetHit.Bounds.Y + targetHit.Bounds.Height / 2);

                if (isDraggingGroup && targetHit.Type == HitType.GroupHeader)
                {
                    // 分组拖拽到分组
                    if (draggedName != targetHit.Group.Name)
                    {
                        GroupDragDropRequested?.Invoke(draggedName, targetHit.Group.Name, insertBefore);
                    }
                }
                else if (isDraggingGroup && targetHit.Type == HitType.Sheet)
                {
                    // 分组拖拽到工作表（移动到该工作表位置）
                    GroupDragDropRequested?.Invoke(draggedName, targetHit.Sheet, insertBefore);
                }
                else if (isDraggingSheet && targetHit.Type == HitType.Sheet)
                {
                    // 工作表拖拽到工作表
                    if (draggedName != targetHit.Sheet)
                    {
                        SheetDragDropRequested?.Invoke(draggedName, targetHit.Sheet, insertBefore);
                    }
                }
                else if (isDraggingSheet && targetHit.Type == HitType.GroupHeader)
                {
                    // 工作表拖拽到分组头（移动到分组第一个位置）
                    SheetDragDropRequested?.Invoke(draggedName, targetHit.Group.Name, insertBefore);
                }
            }
            catch { }
            finally
            {
                _dragOverHit = null;
                Invalidate();
            }
        }

        /// <summary>
        /// 绘制拖放插入位置指示线。
        /// </summary>
        private void DrawDropIndicator(Graphics g, HitInfo hit)
        {
            if (hit == null || hit.Bounds.IsEmpty) return;

            // 根据鼠标位置决定在上方还是下方绘制指示线
            int lineY = hit.Bounds.Top; // 默认在上方
            
            using (var pen = new Pen(Color.DodgerBlue, 2))
            {
                g.DrawLine(pen, 0, lineY, Width, lineY);
                // 绘制左侧箭头
                using (var brush = new SolidBrush(Color.DodgerBlue))
                {
                    g.FillPolygon(brush, new Point[] {
                        new Point(0, lineY - 3),
                        new Point(0, lineY + 3),
                        new Point(6, lineY)
                    });
                    // 绘制右侧箭头
                    g.FillPolygon(brush, new Point[] {
                        new Point(Width, lineY - 3),
                        new Point(Width, lineY + 3),
                        new Point(Width - 6, lineY)
                    });
                }
            }
        }
    }
}


