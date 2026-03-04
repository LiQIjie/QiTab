using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using SheetNavi;
/*文件编码：UTF-8 无BOM*/
namespace Qtab
{
    /// <summary>
    /// VSTO `ThisAddIn`：Excel 加载项的主类。
    /// - 负责创建与维护自定义任务面板（`CustomTaskPane`）以及其承载的 UI 控件 `Qtab`；
    /// - 订阅 Excel 应用的各种事件（工作簿/窗口/工作表），并在恰当的时机刷新 UI；
    /// - 管理每个 Excel 窗口对应的任务面板（避免重复创建与丢失关联）；
    /// - 实现分组、取消分组、颜色、重命名、移动等对工作表/分组的业务逻辑，并持久化到工作簿的 CustomXML。
    /// 注意：为提高稳定性，某些刷新与显示操作采用计时器延迟（例如新建工作簿后 200ms 再刷新）。
    /// </summary>
    public partial class ThisAddIn
    {
        private CustomTaskPane _sheetPane;               // 当前活动窗口对应的任务面板引用（便于 Dock/Visible 操作）
        private QTabNavi _control;  
        private bool _userClosedPane = false;
        private int PanelWidth = 210;
        // 当前活动窗口的主 UI 控件（指向 _controlsByWindow 中的实例）

        // **关键改动**：每窗口独立的控件实例缓存（键：Window.Caption）
        private readonly Dictionary<string, QTabNavi> _controlsByWindow = new Dictionary<string, QTabNavi>(StringComparer.OrdinalIgnoreCase);
        // 每窗口的任务面板缓存：键使用 `ActiveWindow.Caption`，避免同一窗口重复创建面板
        private readonly Dictionary<string, CustomTaskPane> _panesByWindow = new Dictionary<string, CustomTaskPane>(StringComparer.OrdinalIgnoreCase);

        // 每个工作簿的分组信息（键：工作簿 Key -> 值：分组字典（分组名 -> 成员列表与颜色））
        private Dictionary<string, Dictionary<string, (List<string> Members, Color Color)>> _groupsByWorkbook = new Dictionary<string, Dictionary<string, (List<string>, Color)>>();

        // **新增**：记录每个工作簿的可见工作表计数，用于识别未触发事件的复制（Ctrl+拖拽）
        private readonly Dictionary<string, int> _sheetCountsByWorkbook = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        // **新增**：记录每个工作簿当前可见工作表的顺序，用于识别顺序变化（拖拽移动等）
        private readonly Dictionary<string, List<string>> _sheetOrderByWorkbook = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

        // 记录添加到 Excel "Ply" 右键菜单的按钮（便于关闭时删除）
        private List<Office.CommandBarButton> _addedButtons = new List<Office.CommandBarButton>();
        private List<Office.CommandBarControl> _hiddenPlyControls = new List<Office.CommandBarControl>();

        // 刷新用计时器：用于对频繁事件（如新建工作表）进行防抖，确保 Excel 新对象稳定后再刷新
        private System.Windows.Forms.Timer _refreshTimer;
        // **新增**：轻量轮询计时器，用于检测在 Excel 中拖拽移动/复制标签后未触发事件的变化
        private System.Windows.Forms.Timer _pollTimer;

        // 生产环境日志：此处保留空实现，避免磁盘 IO
        private void Log(string fmt, params object[] args) { /* no-operate in release */ }

        /// <summary>
        /// 加载项启动：创建 UI 控件与任务面板，配置计时器与订阅 Excel 事件。
        /// </summary>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Log("Startup called");
            CreateControlAndPane();
            Log("After CreateControlAndPane: sheetPanePresent={0}, controlPresent={1}", _sheetPane != null, _control != null);

            // 启动时刷新工具按钮状态
            try { UpdateToolButtonsStatus(); } catch { }

            // 创建刷新计时器（200ms 防抖）
            _refreshTimer = new System.Windows.Forms.Timer();
            _refreshTimer.Interval = 200; // ms - 稍等一会再刷新，保证 Excel 新对象稳定
            _refreshTimer.Tick += (s, ev) =>
            {
                try
                {
                    _refreshTimer.Stop();
                    RefreshSheetList();
                }
                catch (Exception ex) { Log("Refresh timer exception: {0}", ex.Message); }
            };

            // **新增**：创建轻量轮询计时器以检测未触发事件的移动/复制
            _pollTimer = new System.Windows.Forms.Timer();
            _pollTimer.Interval = 500; // 500ms，足够轻量且响应及时
            _pollTimer.Tick += (s, ev) =>
            {
                try
                {
                    var wb = this.Application.ActiveWorkbook;
                    if (wb == null) return;
                    var key = GetWorkbookKey(wb);

                    var visibleSheets = wb.Worksheets.Cast<Excel.Worksheet>()
                        .Where(w => { try { return w.Visible == Excel.XlSheetVisibility.xlSheetVisible; } catch { return true; } })
                        .OrderBy(w => w.Index)
                        .Select(w => w.Name)
                        .ToList();
                    var visibleCount = visibleSheets.Count;

                    bool needRefresh = false;
                    if (_sheetCountsByWorkbook.TryGetValue(key, out var last) && last != visibleCount) needRefresh = true;
                    if (_sheetOrderByWorkbook.TryGetValue(key, out var lastOrder))
                    {
                        if (lastOrder == null || lastOrder.Count != visibleSheets.Count || !lastOrder.SequenceEqual(visibleSheets)) needRefresh = true;
                    }
                    else
                    {
                        _sheetOrderByWorkbook[key] = new List<string>(visibleSheets);
                    }

                    if (needRefresh)
                    {
                        _sheetCountsByWorkbook[key] = visibleCount;
                        _sheetOrderByWorkbook[key] = new List<string>(visibleSheets);
                        RefreshSheetList();
                    }
                }
                catch { }
            };

            // 订阅 Excel 应用事件：窗口/工作簿/工作表级，统一使用防抖计时器与延迟显示
            var app = this.Application;
            app.WorkbookActivate += wb => { PostShowPane(); try { UpdateToolButtonsStatus(); } catch { } };
           
            ((Excel.AppEvents_Event)app).NewWorkbook += wb =>
            {
                try
                {
                    _refreshTimer.Stop();
                    _refreshTimer.Interval = 200;
                    _refreshTimer.Start();
                    EnsurePaneExists();
                    TryAddPlyContextItems();
                    PostShowPane();
                    UpdateToolButtonsStatus();
                }
                catch (Exception ex) { Log("NewWorkbook handler exception: {0}", ex.Message); }
            };
            //((Excel.AppEvents_Event)app).WorkbookOpen += wb =>
            //{
            //    try
            //    {
            //        _refreshTimer.Stop();
            //        _refreshTimer.Interval = 200;
            //        _refreshTimer.Start();
            //        EnsurePaneExists();
            //        TryAddPlyContextItems();
            //        PostShowPane();
            //        UpdateToolButtonsStatus();
            //    }
            //    catch (Exception ex) { Log("WorkbookOpen handler exception: {0}", ex.Message); }
            //};
            ((Excel.AppEvents_Event)app).WindowActivate += (wb, wn) =>
            {
                try
                {
                    EnsurePaneExists();
                    TryAddPlyContextItems();

                    // 开启轮询以捕捉未触发事件的移动/复制
                    try { _pollTimer?.Start(); } catch { }

                    // **新增**：窗口激活时也进行一次工作表数量与顺序检查，兜底刷新
                    try
                    {
                        var key = GetWorkbookKey(wb);
                        var visibleSheets = wb.Worksheets.Cast<Excel.Worksheet>()
                            .Where(w => { try { return w.Visible == Excel.XlSheetVisibility.xlSheetVisible; } catch { return true; } })
                            .OrderBy(w => w.Index)
                            .Select(w => w.Name)
                            .ToList();
                        var visibleCount = visibleSheets.Count;

                        bool needRefresh = false;
                        if (_sheetCountsByWorkbook.TryGetValue(key, out var last) && last != visibleCount) needRefresh = true;
                        if (_sheetOrderByWorkbook.TryGetValue(key, out var lastOrder))
                        {
                            if (lastOrder == null || lastOrder.Count != visibleSheets.Count || !lastOrder.SequenceEqual(visibleSheets)) needRefresh = true;
                        }
                        else
                        {
                            _sheetOrderByWorkbook[key] = new List<string>(visibleSheets);
                        }

                        if (needRefresh)
                        {
                            _sheetCountsByWorkbook[key] = visibleCount;
                            _sheetOrderByWorkbook[key] = new List<string>(visibleSheets);
                            RefreshSheetList();
                        }
                        else if (!_sheetCountsByWorkbook.ContainsKey(key))
                        {
                            _sheetCountsByWorkbook[key] = visibleCount;
                        }
                    }
                    catch { }

                    PostShowPane();
                    UpdateToolButtonsStatus();
                }
                catch (Exception ex) { Log("WindowActivate handler exception: {0}", ex.Message); }
            };
            app.SheetActivate += sh =>
            {
                if (sh is Excel.Worksheet ws)
                {
                    try
                    {
                        // 高亮并刷新（确保颜色在激活时同步到侧边栏）
                        var ctrl = GetControlForActiveWindow();
                        ctrl?.Highlight(ws.Name);
                        RefreshSheetList();
                    }
                    catch (Exception ex) { Log("SheetActivate highlight/refresh error: {0}", ex.Message); }

                    // **新增**：在激活时检查工作表数量与顺序是否变化（用于识别移动/复制）
                    try
                    {
                        var wb = this.Application.ActiveWorkbook;
                        if (wb != null)
                        {
                            var key = GetWorkbookKey(wb);
                            var visibleSheets = wb.Worksheets.Cast<Excel.Worksheet>()
                                .Where(w => { try { return w.Visible == Excel.XlSheetVisibility.xlSheetVisible; } catch { return true; } })
                                .OrderBy(w => w.Index)
                                .Select(w => w.Name)
                                .ToList();
                            var visibleCount = visibleSheets.Count;

                            bool needRefresh = false;
                            if (_sheetCountsByWorkbook.TryGetValue(key, out var last) && last != visibleCount) needRefresh = true;
                            if (_sheetOrderByWorkbook.TryGetValue(key, out var lastOrder))
                            {
                                if (lastOrder == null || lastOrder.Count != visibleSheets.Count || !lastOrder.SequenceEqual(visibleSheets)) needRefresh = true;
                            }
                            else
                            {
                                _sheetOrderByWorkbook[key] = new List<string>(visibleSheets);
                            }

                            if (needRefresh)
                            {
                                _sheetCountsByWorkbook[key] = visibleCount;
                                _sheetOrderByWorkbook[key] = new List<string>(visibleSheets);
                                RefreshSheetList();
                            }
                            else if (!_sheetCountsByWorkbook.ContainsKey(key))
                            {
                                _sheetCountsByWorkbook[key] = visibleCount;
                            }
                        }
                    }
                    catch (Exception ex) { Log("SheetActivate order/count check error: {0}", ex.Message); }
                }
            };

            // 订阅工作表被删除事件：当用户在 Excel 中删除工作表时，更新分组数据并刷新列表
            try
            {
                app.SheetBeforeDelete += (object sh) =>
                {
                    try
                    {
                        var ws = sh as Excel.Worksheet;
                        var wb = ws?.Parent as Excel.Workbook ?? this.Application.ActiveWorkbook;
                        var name = ws?.Name;
                        if (wb != null && !string.IsNullOrEmpty(name))
                        {
                            try
                            {
                                var groups = GetGroupsForWorkbook(wb, createIfMissing: false);
                                if (groups != null)
                                {
                                    bool changed = false;
                                    var keysToRemove = new List<string>();
                                    foreach (var kv in groups.ToList())
                                    {
                                        var members = kv.Value.Item1;
                                        if (members == null) continue;
                                        if (members.Remove(name)) changed = true;
                                        if (members.Count == 0) keysToRemove.Add(kv.Key);
                                    }
                                    foreach (var k in keysToRemove) groups.Remove(k);
                                    if (changed) SaveGroupsForWorkbook(wb);

                                    var ctrl = GetControlForActiveWindow();
                                    ctrl?.ApplyGroups(groups);
                                }
                            }
                            catch (Exception ex) { Log("SheetBeforeDelete groups update error: {0}", ex.Message); }
                        }
                    }
                    catch (Exception ex) { Log("SheetBeforeDelete handler error: {0}", ex.Message); }
                    finally
                    {
                        try { _refreshTimer.Stop(); _refreshTimer.Interval = 200; _refreshTimer.Start(); } catch { }
                        try { EnsurePaneExists(); } catch { }
                        try { PostShowPane(); } catch { }
                    }
                };
            }
            catch (Exception ex)
            {
                // 有些 Excel 版本/宿主可能不支持此事件，安全忽略
                Log("Subscribe SheetBeforeDelete error: {0}", ex.Message);
            };




            app.WorkbookNewSheet += (Excel.Workbook wb, object sh) =>
            {
                try
                {
                    // 新增/复制工作表：立即刷新列表（同时保留防抖以覆盖快速连发情况）
                    try { RefreshSheetList(); } catch { }

                    _refreshTimer.Stop();
                    _refreshTimer.Interval = 200;
                    _refreshTimer.Start();

                    EnsurePaneExists();
                    TryAddPlyContextItems();
                    PostShowPane();
                    UpdateToolButtonsStatus();
                }
                catch (Exception ex) { Log("WorkbookNewSheet handler exception: {0}", ex.Message); }
            };
            app.SheetChange += (object sh, Excel.Range target) => { /* 可选：响应单元格更改 */ };
            app.SheetBeforeRightClick += (object sh, Excel.Range target, ref bool cancel) =>
            {
                try { _refreshTimer.Stop(); _refreshTimer.Interval = 200; _refreshTimer.Start(); } catch (Exception ex) { Log("SheetBeforeRightClick Refresh error: {0}", ex.Message); }
                try { TryAddPlyContextItems(); } catch (Exception ex) { Log("TryAddPlyContextItems error: {0}", ex.Message); }
            };

            // 初次尝试向工作表标签（Ply）菜单添加入口
            TryAddPlyContextItems();
        }

        /// <summary>
        /// 打印当前任务面板集合状态（用于调试）。
        /// </summary>
        private void LogCustomTaskPanesState(string when)
        {
            try
            {
                var list = this.CustomTaskPanes.Cast<CustomTaskPane>().Select(p => p.Title + "(" + (p.Control?.GetType().Name ?? "null") + ")").ToList();
                Log("CustomTaskPanes [{0}] at {1}: {2}", list.Count, when, string.Join(", ", list));
            }
            catch (Exception ex) { Log("LogCustomTaskPanesState error: {0}", ex.Message); }
        }

        private string GetActiveWindowKey()
        {
            try { return this.Application?.ActiveWindow?.Caption ?? string.Empty; } catch { return string.Empty; }
        }

        /// <summary>获取当前活动窗口的控制器实例。</summary>
        private QTabNavi GetControlForActiveWindow()
        {
            var key = GetActiveWindowKey();
            if (string.IsNullOrEmpty(key)) return _control;
            if (_controlsByWindow.TryGetValue(key, out var ctrl)) return ctrl;
            return _control;
        }

        private CustomTaskPane GetPaneForActiveWindow()
        {
            try
            {
                var key = GetActiveWindowKey();
                if (string.IsNullOrEmpty(key)) return null;
                if (_panesByWindow.TryGetValue(key, out var pane) && pane != null) return pane;

                foreach (CustomTaskPane p in this.CustomTaskPanes)
                {
                    try
                    {
                        var wcap = (p.Window as Excel.Window)?.Caption;
                        if (!string.IsNullOrEmpty(wcap) && string.Equals(wcap, key, StringComparison.OrdinalIgnoreCase))
                        {
                            _panesByWindow[key] = p;
                            return p;
                        }
                    }
                    catch { }
                }
            }
            catch { }
            return null;
        }

        // 返回任务面板标题（在“Qtab”前加上一个 Unicode 图标）。
        private string GetPaneTitle()
        {
            // 使用书签符号 📑 (U+1F4D1) 作为前缀；任务窗格标题不支持位图，只能用文本。
            try { return char.ConvertFromUtf32(0x1F4D1) + " Qtab"; } catch { return "Qtab"; }
        }

        /// <summary>
        /// 为当前活动窗口确保一个专属的 QTabNavi 控件和 CustomTaskPane。
        /// **关键改动**：为每个窗口创建独立的控件实例。
        /// </summary>
        private CustomTaskPane EnsurePaneForActiveWindow()
        {
            var key = GetActiveWindowKey();
            if (string.IsNullOrEmpty(key)) return null;

            var pane = GetPaneForActiveWindow();
            if (pane != null)
            {
                _sheetPane = pane;
                _control = pane.Control as QTabNavi;
                return pane;
            }

            try
            {
                var control = new QTabNavi();
                control.OnSheetSelected += name => ActivateSheet(name);
                control.DockLeftRequested += () => SetDock(Office.MsoCTPDockPosition.msoCTPDockPositionLeft);
                control.DockRightRequested += () => SetDock(Office.MsoCTPDockPosition.msoCTPDockPositionRight);
                control.GroupRequested += names => GroupSheetsForActiveWorkbook(names);
                control.UngroupRequested += names => UngroupSheetsForActiveWorkbook(names);
                control.RenameGroupRequested += names => RenameGroupForActiveWorkbook(names);
                control.ColorRequested += (names, color) => ColorTabsForActiveWorkbook(names, color);
                control.RenameSheetRequested += name => HandleRenameRequest(name);
                control.GroupRenameCommitted += (oldName, newName) => RenameGroupKeyForActiveWorkbook(oldName, newName);
                control.DeleteSheetRequested += names => DeleteSheetsForActiveWorkbook(names);
                control.ShowExcelContextMenu += (sheet, screenPos) => ShowMergedPlyMenu(sheet, screenPos);
                control.RenameSheetCommitted += (oldName, newName) =>
                {
                    try
                    {
                        var wb = this.Application.ActiveWorkbook;
                        if (wb == null) return;
                        try
                        {
                            var ws = wb.Worksheets.Cast<Excel.Worksheet>().FirstOrDefault(x => x.Name == oldName);
                            if (ws != null && !string.IsNullOrWhiteSpace(newName) && newName != oldName)
                            {
                                try { ws.Name = newName; } catch { }
                            }
                        }
                        catch { }

                        try
                        {
                            var groups = GetGroupsForWorkbook(wb, createIfMissing: false);
                            if (groups != null)
                            {
                                bool changed = false;
                                foreach (var kv in groups.ToList())
                                {
                                    var members = kv.Value.Item1;
                                    if (members == null) continue;
                                    for (int i = 0; i < members.Count; i++)
                                    {
                                        if (members[i] == oldName)
                                        {
                                            members[i] = newName;
                                            changed = true;
                                        }
                                    }
                                }
                                if (changed)
                                {
                                    var currentControl = GetControlForActiveWindow();
                                    currentControl?.ApplyGroups(groups);
                                    SaveGroupsForWorkbook(wb);
                                }
                            }
                        }
                        catch { }

                        // **优化**：重命名后刷新是必要的（结构变化）
                        try { RefreshSheetList(); } catch { }
                    }
                    catch { }
                };
                control.SheetDragDropRequested += (dragged, target, before) => HandleSheetDragDrop(dragged, target, before);
                control.GroupDragDropRequested += (dragged, target, before) => HandleGroupDragDrop(dragged, target, before);

                // 仅工具切换
                control.ToolsToggleRequested += () => ToggleBothTools();

                var newPane = this.CustomTaskPanes.Add(control, GetPaneTitle(), this.Application.ActiveWindow);
                newPane.Width = PanelWidth;
                newPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;

                try
                {
                    newPane.VisibleChanged += (s, e) =>
                    {
                        try { var p = s as CustomTaskPane; if (p != null) _userClosedPane = !p.Visible; } catch { }
                    };
                }
                catch { }

                _controlsByWindow[key] = control;
                _panesByWindow[key] = newPane;
                _sheetPane = newPane;
                _control = control;

                // 初始化工具按钮状态
                try { UpdateToolButtonsStatus(control); } catch { }
                return newPane;
            }
            catch { }
            return null;
        }

        /// <summary>
        /// 创建 UI 控件与任务面板（如果需要），并将其显示。
        /// </summary>
        private void CreateControlAndPane()
        {
            try
            {
                Log("CreateControlAndPane start");
                var pane = EnsurePaneForActiveWindow();
                if (pane != null)
                {
                    // 标题已在 Add(...) 时设置，Title 为只读，这里只需显示窗格
                    try { pane.Visible = true; } catch { }
                    _sheetPane = pane;
                    Log("CustomTaskPane added/ensured and made visible");
                }
            }
            catch (Exception ex) { Log("CreateControlAndPane exception: {0}", ex.Message); }
        }

        /// <summary>
        /// 确保当前活动窗口的任务面板存在；如缺失则创建。
        /// </summary>
        private void EnsurePaneExists()
        {
            try
            {
                Log("EnsurePaneExists called");
                var pane = GetPaneForActiveWindow();
                var ctrl = GetControlForActiveWindow();
                if (pane == null || ctrl == null)
                {
                    Log("Pane not found or control null, recreating for active window");
                    CreateControlAndPane();
                    pane = GetPaneForActiveWindow();
                }
                if (pane != null)
                {
                    _sheetPane = pane;
                    _control = pane.Control as QTabNavi;
                }
            }
            catch (Exception ex) { Log("EnsurePaneExists exception: {0}", ex.Message); }
        }

        /// <summary>
        /// 刷新任务面板中的工作表列表与分组显示。
        /// **改动**：使用 GetControlForActiveWindow() 获取当前窗口的控件实例。
        /// **优化**：减少不必要的刷新调用，仅在结构变化时使用。
        /// </summary>
        private void RefreshSheetList()
        {
            try
            {
                Log("RefreshSheetList start");
                EnsurePaneExists();

                var wb = this.Application.ActiveWorkbook;
                var control = GetControlForActiveWindow();

                if (control == null)
                {
                    Log("RefreshSheetList: no control for active window");
                    return;
                }

                if (wb == null)
                {
                    try { control.LoadSheets(new List<string>(), 0, 0); } catch (Exception ex) { Log("LoadSheets empty error: {0}", ex.Message); }
                    try { if (_sheetPane != null && !_userClosedPane) _sheetPane.Visible = true; } catch (Exception ex) { Log("Set pane visible error: {0}", ex.Message); }
                    Log("RefreshSheetList: no active workbook");
                    return;
                }

                var allSheets = wb.Worksheets.Cast<Excel.Worksheet>().OrderBy(s => s.Index).ToList();
                var totalCount = allSheets.Count;
                var visibleSheets = allSheets.Where(ws =>
                {
                    try { return ws.Visible == Excel.XlSheetVisibility.xlSheetVisible; }
                    catch { return true; }
                }).ToList();
                var visibleNames = visibleSheets.Select(ws => ws.Name).ToList();
                var hiddenCount = totalCount - visibleSheets.Count;

                // **更新缓存：计数与顺序**
                try
                {
                    var key = GetWorkbookKey(wb);
                    _sheetCountsByWorkbook[key] = visibleSheets.Count;
                    _sheetOrderByWorkbook[key] = new List<string>(visibleNames);
                }
                catch { }

                try { control.LoadSheets(visibleNames, totalCount, hiddenCount); } catch (Exception ex) { Log("LoadSheets error: {0}", ex.Message); }
                Log("RefreshSheetList loaded {0} visible sheets, {1} hidden, {2} total", visibleNames.Count, hiddenCount, totalCount);

                // 构建每个工作表的颜色映射：有分组则用分组颜色覆盖，否则使用工作表自身标签颜色
                try
                {
                    var colors = new Dictionary<string, Color>(StringComparer.OrdinalIgnoreCase);
                    foreach (var ws in visibleSheets)
                    {
                        Color color = Color.Empty;
                        try
                        {
                            var ole = ws.Tab.Color;
                            if (ole != null)
                            {
                                // Excel Interop 可能返回 double（OLE 色值）或 int，甚至是其他可 Convertible 的类型
                                if (ole is double)
                                {
                                    int oleInt = Convert.ToInt32((double)ole);
                                    // OLE 颜色 0 代表“无色”，不要显示为黑色
                                    color = oleInt == 0 ? Color.Empty : ColorTranslator.FromOle(oleInt);
                                }
                                else if (ole is int)
                                {
                                    int oleInt = (int)ole;
                                    color = oleInt == 0 ? Color.Empty : ColorTranslator.FromOle(oleInt);
                                }
                                else
                                {
                                    try
                                    {
                                        int oleInt = Convert.ToInt32(ole);
                                        color = oleInt == 0 ? Color.Empty : ColorTranslator.FromOle(oleInt);
                                    }
                                    catch { }
                                }
                            }
                        }
                        catch { }

                        colors[ws.Name] = color;
                    }

                    var groupsForColors = GetGroupsForWorkbook(wb, createIfMissing: false);
                    if (groupsForColors != null)
                    {
                        foreach (var kv in groupsForColors)
                        {
                            var gcolor = kv.Value.Item2;
                            if (gcolor != Color.Empty)
                            {
                                foreach (var m in kv.Value.Item1)
                                {
                                    if (colors.ContainsKey(m)) colors[m] = gcolor;
                                }
                            }
                        }
                    }

                    control.ApplySheetColors(colors);
                }
                catch (Exception ex) { Log("ApplySheetColors error: {0}", ex.Message); }

                var groups = GetGroupsForWorkbook(wb, createIfMissing: false);
                if (groups != null) try { control.ApplyGroups(groups); } catch (Exception ex) { Log("ApplyGroups error: {0}", ex.Message); }

                try { if (_sheetPane != null && !_userClosedPane) _sheetPane.Visible = true; } catch (Exception ex) { Log("Set pane visible error2: {0}", ex.Message); }
                Log("RefreshSheetList end: pane visible={0}", _sheetPane != null && _sheetPane.Visible);

                // 刷新工具按钮状态
                try { UpdateToolButtonsStatus(control); } catch { }
            }
            catch (Exception ex)
            {
                Log("RefreshSheetList exception: {0}", ex.Message);
                try { if (_sheetPane != null && !_userClosedPane) _sheetPane.Visible = true; } catch (Exception) { }
            }
        }

        /// <summary>
        /// 尝试向工作表标签（Ply）菜单添加 "Show Qtab" 按钮；若菜单不可用或已存在则忽略。
        /// </summary>
        private void TryAddPlyContextItems()
        {
            try
            {
                Office.CommandBar cb = null;
                try { cb = this.Application.CommandBars["Ply"]; } catch { cb = null; }

                // 回退：尝试在 CommandBars 中查找类似工作表标签上下文菜单的弹出 CommandBar
                if (cb == null)
                {
                    try
                    {
                        // 已知关键字：用于在多个语言环境中识别 Ply 菜单的重命名/分组
                        var renameKeywords = new[] { "rename", "名前の変更", "名前の変更(&N)", "重命名", "Renommer", "Umbenennen", "グループ", "グループ化", "グループ解除", "タブの色", "タブの色の変更" };
                        foreach (Office.CommandBar candidate in this.Application.CommandBars)
                        {
                            try
                            {
                                if (candidate.Type != Office.MsoBarType.msoBarTypePopup) continue;
                                // 检查控件的标题以匹配重命名相关关键词
                                foreach (Office.CommandBarControl ctrl in candidate.Controls)
                                {
                                    try
                                    {
                                        var cap = (ctrl.Caption ?? string.Empty).ToLowerInvariant();
                                        foreach (var kw in renameKeywords)
                                        {
                                            if (!string.IsNullOrEmpty(kw) && cap.Contains(kw.ToLowerInvariant()))
                                            {
                                                cb = candidate;
                                                break;
                                            }
                                        }
                                        if (cb != null) break;
                                    }
                                    catch { }
                                }
                                if (cb != null) break;
                            }
                            catch { }
                        }
                    }
                    catch { cb = null; }
                }

                if (cb == null) return;

                // 确保不会在该 CommandBar 上重复添加 Show 按钮（通过 Tag 判断）
                bool alreadyAdded = false;
                try
                {
                    foreach (Office.CommandBarControl c in cb.Controls)
                    {
                        try { if ((c.Tag as string) == "SheetNavi_ShowPane") { alreadyAdded = true; break; } } catch { }
                    }
                }
                catch { }
                if (alreadyAdded) return;

                // 添加一个菜单项用于显示/隐藏 Qtab 面板
                try
                {
                    var btnShow = (Office.CommandBarButton)cb.Controls.Add(Office.MsoControlType.msoControlButton, Type.Missing, Type.Missing, 1, true);
                    btnShow.Caption = "Show Qtab";
                    btnShow.Tag = "SheetNavi_ShowPane";
                    btnShow.Click += new Office._CommandBarButtonEvents_ClickEventHandler((Office.CommandBarButton ctrl, ref bool cancelDefault) =>
                    {
                        try
                        {
                            // 用户显式打开面板，清除关闭标记
                            _userClosedPane = false;
                            var pane = EnsurePaneForActiveWindow();
                            if (pane != null)
                            {
                                try { pane.Width = PanelWidth; } catch { }
                                try { pane.Visible = true; } catch { }
                                _sheetPane = pane;
                                try { _control?.Focus(); } catch { }
                            }
                        }
                        catch { }
                    });
                    _addedButtons.Add(btnShow);
                }
                catch (Exception ex) { Log("Add Show button error: {0}", ex.Message); }
            }
            catch (Exception ex)
            {
                Log("Add Show button error: {0}", ex.Message);
            }
        }

        /// <summary>返回当前窗口中被选中的工作表名列表。</summary>
        private List<string> GetSelectedSheetNames()
        {
            try
            {
                var sel = this.Application.ActiveWindow.SelectedSheets;
                var list = new List<string>();
                foreach (Excel.Worksheet ws in sel)
                {
                    list.Add(ws.Name);
                }
                return list;
            }
            catch
            {
                return new List<string>();
            }
        }

        /// <summary>根据工作簿对象生成唯一键（优先 FullName，退化为 Name@HashCode）。</summary>
        private string GetWorkbookKey(Excel.Workbook wb)
        {
            if (wb == null) return string.Empty;
            try
            {
                var key = wb.FullName;
                if (string.IsNullOrEmpty(key)) key = wb.Name + "@" + wb.GetHashCode();
                return key;
            }
            catch
            {
                return wb.Name + "@" + wb.GetHashCode();
            }
        }

        /// <summary>获取指定工作簿的分组字典；可选择不存在时创建。</summary>
        private Dictionary<string, (List<string> Members, Color)> GetGroupsForWorkbook(Excel.Workbook wb, bool createIfMissing = true)
        {
            var key = GetWorkbookKey(wb);
            if (string.IsNullOrEmpty(key)) return new Dictionary<string, (List<string>, Color)>();
            if (!_groupsByWorkbook.TryGetValue(key, out var groups))
            {
                // 尝试从工作簿存储加载
                var loaded = LoadGroupsFromWorkbook(wb);
                if (loaded != null)
                {
                    _groupsByWorkbook[key] = loaded;
                    return loaded;
                }

                if (createIfMissing)
                {
                    groups = new Dictionary<string, (List<string>, Color)>();
                    _groupsByWorkbook[key] = groups;
                }
            }
            return groups ?? new Dictionary<string, (List<string>, Color)>();
        }

        /// <summary>从工作簿的 CustomXMLParts 读取并构建分组信息，必要时恢复标签颜色。</summary>
        private Dictionary<string, (List<string> Members, Color)> LoadGroupsFromWorkbook(Excel.Workbook wb)
        {
            try
            {
                var parts = wb.CustomXMLParts;
                for (int i = 1; i <= parts.Count; i++)
                {
                    var part = parts[i];
                    if (part == null) continue;
                    var xml = part.XML as string;
                    if (string.IsNullOrEmpty(xml)) continue;
                    if (xml.IndexOf("<SheetNaviGroups", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        try
                        {
                            var doc = XElement.Parse(xml);
                            var dict = new Dictionary<string, (List<string>, Color)>();
                            foreach (var g in doc.Elements("Group"))
                            {
                                var gname = (string)g.Attribute("name") ?? "Group";
                                var colorAttr = (string)g.Attribute("color");
                                Color c = System.Drawing.Color.Empty;
                                if (!string.IsNullOrEmpty(colorAttr))
                                {
                                    try { c = ColorTranslator.FromHtml(colorAttr); } catch { c = Color.Empty; }
                                }
                                var sheets = g.Elements("Sheet").Select(x => (string)x).Where(s => !string.IsNullOrEmpty(s)).ToList();
                                dict[gname] = (sheets, c);
                            }

                            // 将加载的颜色应用回工作表标签
                            try
                            {
                                foreach (var kv in dict)
                                {
                                    var members = kv.Value.Item1;
                                    var color = kv.Value.Item2;
                                    if (color == Color.Empty) continue;
                                    var wbSheets = wb.Worksheets.Cast<Excel.Worksheet>().ToList();
                                    foreach (var name in members)
                                    {
                                        var ws = wbSheets.FirstOrDefault(x => x.Name == name);
                                        if (ws != null)
                                        {
                                            try { ws.Tab.Color = ColorTranslator.ToOle(color); } catch { }
                                        }
                                    }
                                }
                            }
                            catch { }

                            return dict;
                        }
                        catch { }
                    }
                }
            }
            catch
            {
                // 忽略
            }
            return null;
        }

        /// <summary>将分组信息保存到工作簿的 CustomXMLParts（覆盖旧数据）。</summary>
        private void SaveGroupsForWorkbook(Excel.Workbook wb)
        {
            try
            {
                var key = GetWorkbookKey(wb);
                if (string.IsNullOrEmpty(key)) return;
                if (!_groupsByWorkbook.TryGetValue(key, out var groups)) return;

                // 移除现有的 SheetNavi XML 部分
                try
                {
                    var parts = wb.CustomXMLParts;
                    for (int i = parts.Count; i >= 1; i--)
                    {
                        try
                        {
                            var p = parts[i];
                            if (p != null)
                            {
                                var xml = p.XML as string;
                                if (!string.IsNullOrEmpty(xml) && xml.IndexOf("<SheetNaviGroups", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    p.Delete();
                                }
                            }
                        }
                        catch { }
                    }
                }
                catch { }

                // 构建存储用的 XML
                var root = new XElement("SheetNaviGroups",
                    groups.Select(kv => new XElement("Group",
                        new XAttribute("name", kv.Key),
                        kv.Value.Item2 != Color.Empty ? new XAttribute("color", ColorTranslator.ToHtml(kv.Value.Item2)) : null,
                        kv.Value.Item1.Select(s => new XElement("Sheet", s))
                    ))
                );
                wb.CustomXMLParts.Add(root.ToString());
            }
            catch
            {
                // 忽略
            }
        }

        /// <summary>将选中工作表分组、设置默认颜色，并让成员在工作簿中连续排列。</summary>
        private void GroupSheetsForActiveWorkbook(List<string> names)
        {
            try
            {
                var wb = this.Application.ActiveWorkbook;
                if (wb == null || names == null || names.Count == 0) return;
                var groups = GetGroupsForWorkbook(wb);

                // 将这些工作表从任何已有分组中移除，以便它们加入到新的分组
                var keysToRemove = new List<string>();
                foreach (var kv in groups.ToList())
                {
                    var members = kv.Value.Item1;
                    if (members == null) continue;
                    foreach (var n in names)
                    {
                        if (members.Contains(n)) members.Remove(n);
                    }
                    if (members.Count == 0) keysToRemove.Add(kv.Key);
                }
                foreach (var k in keysToRemove) groups.Remove(k);

                // 找到一个未使用的分组名
                int idx = 1;
                string gname;
                do
                {
                    gname = "Group " + idx;
                    idx++;
                } while (groups.ContainsKey(gname));

                groups[gname] = (new List<string>(names), Color.LightBlue);

                try
                {
                    var allSheets = wb.Worksheets.Cast<Excel.Worksheet>().OrderBy(s => s.Index).ToList();
                    var selectedInOrder = allSheets.Where(s => names.Contains(s.Name)).OrderBy(s => s.Index).ToList();
                    
                    if (selectedInOrder.Count > 1)
                    {
                        Excel.Worksheet anchor = selectedInOrder[0];
                        Excel.Worksheet prev = anchor;
                        
                        for (int i = 1; i < selectedInOrder.Count; i++)
                        {
                            var current = selectedInOrder[i];
                            try
                            {
                                current.Move(Type.Missing, prev);
                                prev = current;
                            }
                            catch { }
                        }
                    }
                }
                catch { }

                // 为分组的标签设置默认颜色
                foreach (var n in names)
                {
                    var ws = wb.Worksheets.Cast<Excel.Worksheet>().FirstOrDefault(x => x.Name == n);
                    if (ws != null)
                    {
                        try
                        {
                            ws.Tab.Color = System.Drawing.ColorTranslator.ToOle(Color.LightBlue);
                        }
                        catch { }
                    }
                }

                // **优化**：直接更新当前窗口控件
                var control = GetControlForActiveWindow();
                control?.ApplyGroups(groups);

                SaveGroupsForWorkbook(wb);
                
                // **优化**：分组后需要刷新（结构变化）
                RefreshSheetList();
            }
            catch
            {
                // 忽略错误
            }
        }

        /// <summary>取消分组：清空成员颜色，移除空分组，刷新并持久化。</summary>
        private void UngroupSheetsForActiveWorkbook(List<string> names)
        {
            try
            {
                var wb = this.Application.ActiveWorkbook;
                if (wb == null || names == null || names.Count == 0) return;
                var groups = GetGroupsForWorkbook(wb, createIfMissing: false);
                if (groups == null) return;

                var keysToRemove = new List<string>();
                foreach (var kv in groups)
                {
                    var members = kv.Value.Item1;
                    if (members == null) continue;
                    foreach (var n in names)
                    {
                        if (members.Contains(n))
                        {
                            members.Remove(n);
                            // 重置标签颜色
                            var ws = wb.Worksheets.Cast<Excel.Worksheet>().FirstOrDefault(x => x.Name == n);
                            if (ws != null)
                            {
                                try
                                {
                                    ws.Tab.ColorIndex = Excel.XlColorIndex.xlColorIndexNone;
                                }
                                catch { }
                            }
                        }
                    }
                    if (members.Count == 0) keysToRemove.Add(kv.Key);
                }

                foreach (var k in keysToRemove) groups.Remove(k);

                // **优化**：直接更新当前窗口控件
                var control = GetControlForActiveWindow();
                control?.ApplyGroups(groups);

                SaveGroupsForWorkbook(wb);
                
                // **优化**：取消分组后需要刷新（结构变化）
                RefreshSheetList();
            }
            catch
            {
            }
        }

        /// <summary>分组重命名：唯一化新名称、更新显示并持久化。</summary>
        private void RenameGroupForActiveWorkbook(List<string> names)
        {
            try
            {
                var wb = this.Application.ActiveWorkbook;
                if (wb == null || names == null || names.Count == 0) return;
                var groups = GetGroupsForWorkbook(wb, createIfMissing: false);
                if (groups == null) return;

                var target = groups.FirstOrDefault(kv => kv.Value.Item1.Contains(names[0]));
                if (string.IsNullOrEmpty(target.Key)) return;
                var oldName = target.Key;
                var newName = ShowInputBox("Rename group:", "Rename Group", oldName);
                if (string.IsNullOrWhiteSpace(newName) || newName == oldName) return;

                if (groups.ContainsKey(newName))
                {
                    int i = 1;
                    var baseName = newName;
                    while (groups.ContainsKey(newName))
                    {
                        newName = baseName + " (" + i + ")";
                        i++;
                    }
                }

                var members = target.Value.Item1;
                var color = target.Value.Item2;
                groups.Remove(oldName);
                groups[newName] = (members, color);
                
                // **优化**：直接更新当前窗口控件
                var control = GetControlForActiveWindow();
                control?.ApplyGroups(groups);

                SaveGroupsForWorkbook(wb);
            }
            catch (Exception ex)
            {
                // 忽略
            }
        }

        /// <summary>设置选中工作表标签颜色，并在完全匹配的分组上应用颜色。</summary>
        private void ColorTabsForActiveWorkbook(List<string> names, Color color)
        {
            try
            {
                var wb = this.Application.ActiveWorkbook;
                if (wb == null || names == null || names.Count == 0) return;
                foreach (var n in names)
                {
                    var ws = wb.Worksheets.Cast<Excel.Worksheet>().FirstOrDefault(x => x.Name == n);
                    if (ws != null)
                    {
                        try
                        {
                            ws.Tab.Color = System.Drawing.ColorTranslator.ToOle(color);
                        }
                        catch { }
                    }
                }

                var groups = GetGroupsForWorkbook(wb, createIfMissing: false);
                if (groups != null)
                {
                    foreach (var kv in groups.ToList())
                    {
                        var members = kv.Value.Item1;
                        if (members != null && members.Count > 0 && names.All(n => members.Contains(n)) && members.All(m => names.Contains(m)))
                        {
                            groups[kv.Key] = (members, color);
                        }
                    }
                    
                    // **优化**：直接更新当前窗口控件
                    var control = GetControlForActiveWindow();
                    control?.ApplyGroups(groups);
                    
                    SaveGroupsForWorkbook(wb);
                }
            }
            catch
            {
                // 忽略
            }
        }

        /// <summary>
        /// 处理重命名请求：如果点击的是分组，开始分组内联重命名；否则开始工作表内联重命名。
        /// </summary>
        private void HandleRenameRequest(string sheetName)
        {
            try
            {
                var wb = this.Application.ActiveWorkbook;
                if (wb == null || string.IsNullOrEmpty(sheetName)) return;

                var control = GetControlForActiveWindow();
                if (control == null) return;

                var groups = GetGroupsForWorkbook(wb, createIfMissing: false);
                if (groups != null)
                {
                    var grp = groups.FirstOrDefault(kv => kv.Value.Item1.Contains(sheetName));
                    if (!string.IsNullOrEmpty(grp.Key))
                    {
                        try { control.BeginInlineRenameGroup(grp.Key); } catch { }
                        return;
                    }
                }

                try { control.BeginInlineRenameSheet(sheetName); } catch { }
            }
            catch { }
        }

        /// <summary>将分组字典中的键从旧名替换为新名，并持久化。</summary>
        private void RenameGroupKeyForActiveWorkbook(string oldName, string newName)
        {
            try
            {
                var wb = this.Application.ActiveWorkbook;
                if (wb == null) return;
                var groups = GetGroupsForWorkbook(wb, createIfMissing: false);
                if (groups == null || !groups.ContainsKey(oldName)) return;
                var val = groups[oldName];
                groups.Remove(oldName);
                groups[newName] = val;
                
                // **优化**：直接更新当前窗口控件
                var control = GetControlForActiveWindow();
                control?.ApplyGroups(groups);
                
                SaveGroupsForWorkbook(wb);
            }
            catch { }
        }

        /// <summary>
        /// 设置任务面板的停靠位置（左/右），并保持可见，同时更新箭头指示。
        /// </summary>
        private void SetDock(Office.MsoCTPDockPosition pos)
        {
            if (_sheetPane == null) return;
            try
            {
                _sheetPane.DockPosition = pos;
                _sheetPane.Visible = true;
                
                bool isLeft = pos == Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
                
                // **优化**：使用当前窗口控件
                var control = GetControlForActiveWindow();
                try { control?.SetDockIndicator(isLeft); } catch { }
            }
            catch
            {
            }
        }

        /// <summary>在 Excel 中激活指定名称的工作表。</summary>
        private void ActivateSheet(string name)
        {
            var wb = this.Application.ActiveWorkbook;
            if (wb == null) return;
            var ws = wb.Worksheets.Cast<Excel.Worksheet>().FirstOrDefault(x => x.Name == name);
            if (ws != null)
            {
                try
                {
                    ws.Activate();
                    // **优化**：移除这里的 RefreshSheetList 调用
                    // 激活操作由 SheetActivate 事件处理高亮即可
                }
                catch { }
            }
        }

        /// <summary>
        /// 启用/禁用 Solver 插件。
        /// </summary>
        private void ToggleSolver()
        {
            try
            {
                var addin = FindAddInByKeywords(new[] { "solver", "ソルバー", "求解" });
                if (addin == null)
                {
                    MessageBox.Show("Solver add-in was not found.",
                "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                bool enable = !addin.Installed;
                try { addin.Installed = enable; } catch { }
                UpdateToolButtonsStatus();
                MessageBox.Show(enable ? "Solver has been enabled." : "Solver has been disabled.",
            "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to toggle Solver: " + ex.Message,
            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 启用/禁用 分析工具库（Analysis Tool）。
        /// </summary>
        private void ToggleAnalysisToolpak()
        {
            try
            {
                // 优先选择主 Tool；如存在也可同时切换 VBA 版本
                var atp = FindAddInByKeywords(new[] { "analysis Tool", "データ分析", "分析ツール", "Analysis Tool" });
                if (atp == null)
                {
                    MessageBox.Show("Analysis Tool add-in was not found.",
        "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                bool enable = !atp.Installed;
                try { atp.Installed = enable; } catch { }

                // 尝试同步切换 VBA 版本，但忽略错误
                try
                {
                    var atpVba = FindAddInByKeywords(new[] { "analysis Tool - vba", "分析ツール - VBA" });
                    if (atpVba != null) atpVba.Installed = enable;
                }
                catch { }

                UpdateToolButtonsStatus();
                MessageBox.Show(enable ? "Analysis Tool has been enabled." : "Analysis Tool has been disabled.",
    "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to toggle Solver: Analysis Tool" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 切换两个插件：默认同时开同时关
        /// </summary>
        private void ToggleBothTools()
        {
            try
            {
                var (solverEnabled, analysisEnabled) = GetToolsStatus();
                bool enableBoth = !(solverEnabled && analysisEnabled); // 若两者已全开，则关闭；否则全部打开

                var solver = FindAddInByKeywords(new[] { "solver", "ソルバー", "求解" });
                var atp = FindAddInByKeywords(new[] { "analysis Tool", "データ分析", "分析ツール", "Analysis Tool" });
                var atpVba = FindAddInByKeywords(new[] { "analysis Tool - vba", "分析ツール - VBA" });

                bool changed = false;
                if (solver != null)
                {
                    try { solver.Installed = enableBoth; changed = true; } catch { }
                }
                if (atp != null)
                {
                    try { atp.Installed = enableBoth; changed = true; } catch { }
                }
                if (atpVba != null)
                {
                    try { atpVba.Installed = enableBoth; } catch { }
                }

                UpdateToolButtonsStatus();

                if (changed)
                {
                    if (enableBoth)
                        MessageBox.Show("データ -＞ 分析: Solver And DataAnalysis ON", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    else
                        MessageBox.Show("データ -＞ 分析: Solver And DataAnalysis OFF", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Addin Not Found。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Change Failed: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private (bool solverEnabled, bool analysisEnabled) GetToolsStatus()
        {
            bool solver = false, analysis = false;
            try
            {
                var addin = FindAddInByKeywords(new[] { "solver", "ソルバー", "求解" });
                if (addin != null) solver = SafeGetInstalled(addin);
            }
            catch { }
            try
            {
                var atp = FindAddInByKeywords(new[] { "analysis Tool", "データ分析", "分析ツール", "Analysis Tool" });
                if (atp != null) analysis = SafeGetInstalled(atp);
            }
            catch { }
            return (solver, analysis);
        }

        private bool SafeGetInstalled(Excel.AddIn ai)
        {
            try { return ai != null && ai.Installed; } catch { return false; }
        }

        private Excel.AddIn FindAddInByKeywords(string[] keywords)
        {
            try
            {
                var app = this.Application;
                // 先在 AddIns，再在 AddIns2 中查找
                var col1 = app.AddIns;
                for (int i = 1; i <= col1.Count; i++)
                {
                    Excel.AddIn ai = null; try { ai = col1[i]; } catch { ai = null; }
                    if (ai == null) continue;
                    if (MatchesAddIn(ai, keywords)) return ai;
                }
                try
                {
                    var col2 = app.AddIns2 as Excel.AddIns2;
                    if (col2 != null)
                    {
                        for (int i = 1; i <= col2.Count; i++)
                        {
                            Excel.AddIn ai = null; try { ai = col2[i]; } catch { ai = null; }
                            if (ai == null) continue;
                            if (MatchesAddIn(ai, keywords)) return ai;
                        }
                    }
                }
                catch { }
            }
            catch { }
            return null;
        }

        private bool MatchesAddIn(Excel.AddIn ai, string[] keywords)
        {
            try
            {
                var name = (ai.Name ?? string.Empty).ToLowerInvariant();
                var title = string.Empty;
                try { title = (ai.Title ?? string.Empty).ToLowerInvariant(); } catch { }
                var fullname = string.Empty;
                try { fullname = (ai.FullName ?? string.Empty).ToLowerInvariant(); } catch { }
                foreach (var kw in keywords)
                {
                    var k = (kw ?? string.Empty).ToLowerInvariant();
                    if (name.Contains(k) || title.Contains(k) || fullname.Contains(k)) return true;
                }
            }
            catch { }
            return false;
        }

        /// <summary>加载项关闭：清理添加的菜单按钮与恢复被隐藏的菜单项。</summary>
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // remove added commandbar buttons
            try
            {
                foreach (var btn in _addedButtons)
                {
                    try { btn.Delete(true); } catch { }
                }
                _addedButtons.Clear();
            }
            catch { }

            // 停止轮询计时器
            try { _pollTimer?.Stop(); _pollTimer?.Dispose(); } catch { }

            // restore hidden Ply controls
            RestorePlyControls();
        }

        /// <summary>恢复曾隐藏的 Ply 菜单项（如果有）。</summary>
        private void RestorePlyControls()
        {
            try
            {
                foreach (var ctrl in _hiddenPlyControls)
                {
                    try { ctrl.Visible = true; } catch { }
                    try { ctrl.Enabled = true; } catch { }
                }
                _hiddenPlyControls.Clear();
            }
            catch { }
        }

        /// <summary>
        /// 延迟显示任务面板（默认 200ms）：在窗口/工作簿初始化完成后执行，避免早期调用导致异常或 UI 不更新。
        /// </summary>
        private void PostShowPane()
        {
            try
            {
                try { RefreshSheetList(); } catch { }
                if (!_userClosedPane)
                {
                    ShowPane();
                }
            }
            catch { }
        }

        #region VSTO 生成代码（保持不变）
        private void ThisAddIn_InitializeComponent() { }
        private void ThisAddIn_ShutdownComponent() { }
        private void InternalStartup() { this.Startup += new System.EventHandler(ThisAddIn_Startup); this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown); }
        #endregion

        /// <summary>
        /// 简易输入框工具（已被内联编辑替换，仍保留以便备用）。
        /// 返回用户输入的字符串；当前实现为占位实现，返回空字符串以避免在 Release 环境弹窗。
        /// 如果需要启用交互，可在此处实现简单 WinForms 弹窗输入框。
        /// </summary>
        private string ShowInputBox(string prompt, string title, string defaultValue) { /* 原实现不变 */ return string.Empty; }

        /// <summary>
        /// 显示当前活动窗口的任务面板（保持宽度与可见，不改变用户选择的 Dock 侧），同时更新箭头指示。
        /// </summary>
        private void ShowPane()
        {
            try
            {
                var pane = EnsurePaneForActiveWindow();
                if (pane == null) return;
                try { pane.Width = PanelWidth; } catch { }
                if (!_userClosedPane)
                {
                    try { pane.Visible = true; } catch { }
                }
                _sheetPane = pane;
                bool isLeft = true;
                try { isLeft = pane.DockPosition == Office.MsoCTPDockPosition.msoCTPDockPositionLeft; } catch { }
                
                // **优化**：使用当前窗口控件
                var control = GetControlForActiveWindow();
                try { control?.SetDockIndicator(isLeft); } catch { }
                try { control?.Focus(); } catch { }
                try { UpdateToolButtonsStatus(control); } catch { }
            }
            catch { }
        }

        /// <summary>
        /// 按当前顺序上/下移动选中的工作表，并刷新显示。
        /// 注意：原实现省略，此处保留占位以便项目集中管理实现；调用方（Qtab）会订阅并触发该方法。
        /// 实际实现应在工作簿中按可见顺序移动指定工作表，并在完成后调用 `RefreshSheetList()`。
        /// </summary>
        private void MoveSheets(List<string> names, bool up) { /* 原实现不变 */ }
        /// <summary>
        /// 移动分组顺序（重建字典以保持顺序），并保存。
        /// 注意：原实现省略，此处保留占位以便项目集中管理实现；实际实现应调整 _groupsByWorkbook 内对应工作簿的分组顺序并调用 Save/Refresh。
        /// </summary>
        private void MoveGroup(string gname, bool up) { /* 原实现不变 */ }

        /// <summary>
        /// 处理工作表的拖拽排序。
        /// </summary>
        private void HandleSheetDragDrop(string draggedSheet, string target, bool insertBefore)
        {
            try
            {
                var wb = this.Application.ActiveWorkbook;
                if (wb == null || string.IsNullOrEmpty(draggedSheet) || string.IsNullOrEmpty(target)) return;

                var groups = GetGroupsForWorkbook(wb, createIfMissing: false);
                if (groups == null) groups = new Dictionary<string, (List<string>, Color)>();

                // 找到源工作表所属的分组
                string sourceGroup = null;
                foreach (var kv in groups)
                {
                    if (kv.Value.Item1.Contains(draggedSheet))
                    {
                        sourceGroup = kv.Key;
                        break;
                    }
                }

                // 找到目标所属的分组
                string targetGroup = null;
                foreach (var kv in groups)
                {
                    if (kv.Value.Item1.Contains(target))
                    {
                        targetGroup = kv.Key;
                        break;
                    }
                }

                // 在 Excel 中移动工作表
                var wsToMove = wb.Worksheets.Cast<Excel.Worksheet>().FirstOrDefault(s => s.Name == draggedSheet);
                var wsTarget = wb.Worksheets.Cast<Excel.Worksheet>().FirstOrDefault(s => s.Name == target);

                if (wsToMove != null && wsTarget != null)
                {
                    try
                    {
                        if (insertBefore)
                            wsToMove.Move(wsTarget, Type.Missing);
                        else
                            wsToMove.Move(Type.Missing, wsTarget);
                    }
                    catch (Exception ex) { Log("Move sheet in Excel error: {0}", ex.Message); }
                }

                // 更新分组数据
                if (sourceGroup != null && sourceGroup == targetGroup)
                {
                    // 同组内移动：更新 members 列表顺序
                    var members = groups[sourceGroup].Item1;
                    members.Remove(draggedSheet);
                    int targetIndex = members.IndexOf(target);
                    if (targetIndex >= 0)
                    {
                        if (insertBefore)
                            members.Insert(targetIndex, draggedSheet);
                        else
                            members.Insert(targetIndex + 1, draggedSheet);
                    }
                    else
                    {
                        members.Add(draggedSheet);
                    }
                }
                else
                {
                    // 跨组或无组移动
                    if (sourceGroup != null)
                    {
                        // 从源分组移除
                        groups[sourceGroup].Item1.Remove(draggedSheet);
                        if (groups[sourceGroup].Item1.Count == 0)
                        {
                            groups.Remove(sourceGroup);
                        }
                    }

                    if (targetGroup != null)
                    {
                        // 添加到目标分组
                        var members = groups[targetGroup].Item1;
                        int targetIndex = members.IndexOf(target);
                        if (targetIndex >= 0)
                        {
                            if (insertBefore)
                                members.Insert(targetIndex, draggedSheet);
                            else
                                members.Insert(targetIndex + 1, draggedSheet);
                        }
                        else
                        {
                            members.Add(draggedSheet);
                        }

                        // 应用目标分组的颜色
                        var color = groups[targetGroup].Item2;
                        if (color != Color.Empty && wsToMove != null)
                        {
                            try { wsToMove.Tab.Color = ColorTranslator.ToOle(color); } catch { }
                        }
                    }
                }

                SaveGroupsForWorkbook(wb);
                RefreshSheetList();
            }
            catch (Exception ex)
            {
                Log("HandleSheetDragDrop error: {0}", ex.Message);
            }
        }

        /// <summary>
        /// 处理分组的拖拽排序。
        /// </summary>
        private void HandleGroupDragDrop(string draggedGroup, string target, bool insertBefore)
        {
            try
            {
                var wb = this.Application.ActiveWorkbook;
                if (wb == null || string.IsNullOrEmpty(draggedGroup) || string.IsNullOrEmpty(target)) return;

                var groups = GetGroupsForWorkbook(wb, createIfMissing: false);
                if (groups == null || !groups.ContainsKey(draggedGroup)) return;

                var groupData = groups[draggedGroup];
                var members = groupData.Item1;
                if (members == null || members.Count == 0) return;

                // 获取所有工作表按 Excel 顺序
                var allSheets = wb.Worksheets.Cast<Excel.Worksheet>().OrderBy(s => s.Index).ToList();

                // 找到目标工作表或分组的第一个成员
                Excel.Worksheet targetSheet = null;
                if (groups.ContainsKey(target))
                {
                    // 目标是分组
                    var targetMembers = groups[target].Item1;
                    if (targetMembers.Count > 0)
                    {
                        targetSheet = allSheets.FirstOrDefault(s => targetMembers.Contains(s.Name));
                    }
                }
                else
                {
                    // 目标是工作表
                    targetSheet = allSheets.FirstOrDefault(s => s.Name == target);
                }

                if (targetSheet == null) return;

                // 移动分组的所有成员
                try
                {
                    var firstMember = allSheets.FirstOrDefault(s => members.Contains(s.Name));
                    if (firstMember != null)
                    {
                        if (insertBefore)
                            firstMember.Move(targetSheet, Type.Missing);
                        else
                            firstMember.Move(Type.Missing, targetSheet);

                        // 将其他成员依次移动到第一个成员之后
                        Excel.Worksheet prev = firstMember;
                        foreach (var member in members)
                        {
                            if (member == firstMember.Name) continue;
                            var ws = allSheets.FirstOrDefault(s => s.Name == member);
                            if (ws != null)
                            {
                                ws.Move(Type.Missing, prev);
                                prev = ws;
                            }
                        }
                    }
                }
                catch (Exception ex) { Log("Move group in Excel error: {0}", ex.Message); }

                RefreshSheetList();
            }
            catch (Exception ex)
            {
                Log("HandleGroupDragDrop error: {0}", ex.Message);
            }
        }

        /// <summary>
        /// 選択されたシートを削除します（Excel操作）。
        /// 削除後、グループ情報を更新し、UIをリフレッシュします。
        /// </summary>
        private void DeleteSheetsForActiveWorkbook(List<string> names)
        {
            try
            {
                var wb = this.Application.ActiveWorkbook;
                if (wb == null || names == null || names.Count == 0) return;

                // Excel でシート削除
                foreach (var name in names)
                {
                    try
                    {
                        var ws = wb.Worksheets.Cast<Excel.Worksheet>().FirstOrDefault(x => x.Name == name);
                        if (ws != null)
                        {
                            this.Application.DisplayAlerts = false;
                            try
                            {
                                ws.Delete();
                            }
                            finally
                            {
                                this.Application.DisplayAlerts = true;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Log("Delete sheet '{0}' error: {1}", name, ex.Message);
                    }
                }

                // グループ情報から削除されたシートを除去
                var groups = GetGroupsForWorkbook(wb, createIfMissing: false);
                if (groups != null)
                {
                    var keysToRemove = new List<string>();
                    foreach (var kv in groups.ToList())
                    {
                        var members = kv.Value.Item1;
                        if (members == null) continue;

                        foreach (var name in names)
                        {
                            members.Remove(name);
                        }

                        if (members.Count == 0)
                        {
                            keysToRemove.Add(kv.Key);
                        }
                    }
                    foreach (var k in keysToRemove)
                    {
                        groups.Remove(k);
                    }

                    SaveGroupsForWorkbook(wb);
                    
                    // **优化**：直接更新当前窗口控件
                    var control = GetControlForActiveWindow();
                    control?.ApplyGroups(groups);
                }

                // **优化**：删除后需要刷新（结构变化）
                RefreshSheetList();
            }
            catch (Exception ex)
            {
                Log("DeleteSheetsForActiveWorkbook error: {0}", ex.Message);
            }
        }

        /// <summary>
        /// Excel の Ply（シートタブ）メニューを指定位置に表示します。
        /// </summary>
        private void ShowExcelSheetContextMenu(string sheetName, Point screenPos)
        {
            try
            {
                var wb = this.Application.ActiveWorkbook;
                if (wb == null || string.IsNullOrEmpty(sheetName)) return;

                // 対象シートをアクティブ化
                var ws = wb.Worksheets.Cast<Excel.Worksheet>().FirstOrDefault(x => x.Name == sheetName);
                if (ws == null) return;

                try
                {
                    ws.Activate();
                }
                catch (Exception ex)
                {
                    Log("Activate sheet for context menu error: {0}", ex.Message);
                    return;
                }

                // Ply メニューを取得
                Office.CommandBar plyMenu = null;
                try
                {
                    plyMenu = this.Application.CommandBars["Ply"];
                }
                catch
                {
                    // Ply メニューが見つからない場合のフォールバック処理
                    Log("Ply menu not found");
                    return;
                }

                if (plyMenu != null)
                {
                    try
                    {
                        // コンテキストメニューを表示（スクリーン座標）
                        plyMenu.ShowPopup(screenPos.X, screenPos.Y);
                    }
                    catch (Exception ex)
                    {
                        Log("Show Ply menu error: {0}", ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                Log("ShowExcelSheetContextMenu error: {0}", ex.Message);
            }
        }

        private void ShowMergedPlyMenu(string sheetName, Point screenPos)
        {
            try
            {
                var app = this.Application;
                Bitmap bmp = Properties.Resources.SanDen;
                if (app == null) return;

                // 获取侧边栏当前选择，作为合并菜单操作的作用对象
                var ctrl = GetControlForActiveWindow();
                var currentSelection = ctrl?.GetSelectedNames() ?? new List<string>();
                if ((currentSelection == null || currentSelection.Count == 0) && !string.IsNullOrEmpty(sheetName))
                {
                    currentSelection = new List<string> { sheetName };
                }

                // 删除旧的合并菜单避免重复
                try { app.CommandBars["QtabMergedPly"].Delete(); } catch { }

                // 创建自定义弹出菜单
                object missing = Type.Missing;
                var merged = app.CommandBars.Add(
                    "QtabMergedPly",
                    Office.MsoBarPosition.msoBarPopup,
                    missing,
                    true // Temporary
                );

                // Group
                var btnGroup = (Office.CommandBarButton)merged.Controls.Add(Office.MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
                btnGroup.Caption = "Group";
                try { if (bmp != null) { btnGroup.Picture = ToPictureDisp(bmp); btnGroup.Style = Office.MsoButtonStyle.msoButtonIconAndCaption; } } catch { }
                btnGroup.Click += new Office._CommandBarButtonEvents_ClickEventHandler((Office.CommandBarButton ctrlBtn, ref bool cancelDefault) =>
                {
                    try
                    {
                        var targets = ctrl?.GetSelectedNames() ?? currentSelection;
                        if (targets != null && targets.Count > 0) GroupSheetsForActiveWorkbook(targets);
                    }
                    catch { }
                    finally { try { RefreshSheetList(); } catch { } }
                });

                // Ungroup
                var btnUnGroup = (Office.CommandBarButton)merged.Controls.Add(Office.MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
                btnUnGroup.Caption = "Ungroup";
                try { if (bmp != null) { btnUnGroup.Picture = ToPictureDisp(bmp); btnUnGroup.Style = Office.MsoButtonStyle.msoButtonIconAndCaption; } } catch { }
                btnUnGroup.Click += new Office._CommandBarButtonEvents_ClickEventHandler((Office.CommandBarButton ctrlBtn, ref bool cancelDefault) =>
                {
                    try
                    {
                        var targets = ctrl?.GetSelectedNames() ?? currentSelection;
                        if (targets != null && targets.Count > 0) UngroupSheetsForActiveWorkbook(targets);
                    }
                    catch { }
                    finally { try { RefreshSheetList(); } catch { } }
                });

                // Rename
                var btnRename = (Office.CommandBarButton)merged.Controls.Add(Office.MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
                btnRename.Caption = "Rename";
                try { if (bmp != null) { btnRename.Picture = ToPictureDisp(bmp); btnRename.Style = Office.MsoButtonStyle.msoButtonIconAndCaption; } } catch { }
                btnRename.Click += new Office._CommandBarButtonEvents_ClickEventHandler((Office.CommandBarButton ctrlBtn, ref bool cancelDefault) =>
                {
                    try
                    {
                        var targets = ctrl?.GetSelectedNames() ?? currentSelection;
                        if (targets == null || targets.Count == 0) return;

                        var wb = this.Application.ActiveWorkbook;
                        var groups = wb != null ? GetGroupsForWorkbook(wb, createIfMissing: false) : null;

                        if (targets.Count == 1)
                        {
                            // 单个工作表：直接开始工作表内联重命名
                            try { ctrl?.BeginInlineRenameSheet(targets[0]); } catch { }
                        }
                        else
                        {
                            // 多选：如果完全属于同一分组，则开始分组内联重命名；否则不处理
                            string groupName = null;
                            if (groups != null)
                            {
                                foreach (var kv in groups)
                                {
                                    var members = kv.Value.Item1 ?? new List<string>();
                                    if (members.Count > 0 && targets.All(t => members.Contains(t)))
                                    {
                                        groupName = kv.Key;
                                        break;
                                    }
                                }
                            }

                            if (!string.IsNullOrEmpty(groupName))
                            {
                                try { ctrl?.BeginInlineRenameGroup(groupName); } catch { }
                            }
                        }
                    }
                    catch { }
                    finally { try { RefreshSheetList(); } catch { } }
                });

                // Delete
                var btnDelete = (Office.CommandBarButton)merged.Controls.Add(Office.MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
                btnDelete.Caption = "Delete";
                try { if (bmp != null) { btnDelete.Picture = ToPictureDisp(bmp); btnDelete.Style = Office.MsoButtonStyle.msoButtonIconAndCaption; } } catch { }
                btnDelete.Click += new Office._CommandBarButtonEvents_ClickEventHandler((Office.CommandBarButton ctrlBtn, ref bool cancelDefault) =>
                {
                    try
                    {
                        var targets = ctrl?.GetSelectedNames() ?? currentSelection;
                        if (targets == null || targets.Count == 0) return;
                        var msg = targets.Count == 1 ? $"シート'{targets[0]}'を削除しますか?" : $"{targets.Count}個シートを削除しますか?";
                        if (MessageBox.Show(msg, "削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                        {
                            DeleteSheetsForActiveWorkbook(targets);
                        }
                    }
                    catch { }
                    finally { try { RefreshSheetList(); } catch { } }
                });

                // Set Tab Color...
                var btnColor = (Office.CommandBarButton)merged.Controls.Add(Office.MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
                btnColor.Caption = "Set Tab Color...";
                try
                {
                
                    if (bmp != null) { btnColor.Picture = ToPictureDisp(bmp); btnColor.Style = Office.MsoButtonStyle.msoButtonIconAndCaption; }
                }
                catch { }
                btnColor.Click += new Office._CommandBarButtonEvents_ClickEventHandler((Office.CommandBarButton ctrlBtn, ref bool cancelDefault) =>
                {
                    try
                    {
                        var targets = ctrl?.GetSelectedNames() ?? currentSelection;
                        if (targets == null || targets.Count == 0) return;
                        using (var dlg = new ColorDialog())
                        {
                            if (dlg.ShowDialog() == DialogResult.OK)
                            {
                                ColorTabsForActiveWorkbook(targets, dlg.Color);
                            }
                        }
                    }
                    catch { }
                    finally { try { RefreshSheetList(); } catch { } }
                });

                // 分隔与拼接原生菜单
                var sep = merged.Controls.Add(Office.MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
                try { sep.BeginGroup = true; } catch { }
                try { sep.Visible = false; } catch { }

                try
                {
                    Office.CommandBar ply = null;
                    try { ply = app.CommandBars["Ply"]; } catch { ply = null; }
                    if (ply != null && ply.Controls != null && ply.Controls.Count > 0)
                    {
                        foreach (Office.CommandBarControl c in ply.Controls)
                        {
                            try { c.Copy(merged, Type.Missing); } catch { }
                        }
                    }
                }
                catch { }

                // 显示合并菜单
                try { merged.ShowPopup(screenPos.X, screenPos.Y); } catch { }
            }
            catch (Exception ex)
            {
                Log("ShowMergedPlyMenu error: {0}", ex.Message);
            }
        }

        // 将 Bitmap 转为 IPictureDisp 用于 CommandBarButton 图标
        private static stdole.IPictureDisp ToPictureDisp(Bitmap bmp)
        {
            return (stdole.IPictureDisp)PictureDispConverter.GetIPictureDispFromImage(bmp);
        }
        private class PictureDispConverter : System.Windows.Forms.AxHost
        {
            private PictureDispConverter() : base("{595DF1F1-1C66-11D2-8D3E-00C04F79EFC3}") { }
            public static object GetIPictureDispFromImage(Image image)
            {
                return GetIPictureDispFromPicture(image);
            }
        }

        // 添加：根据当前 Excel 插件状态更新面板按钮显示
        private void UpdateToolButtonsStatus(QTabNavi ctrl = null)
        {
            try
            {
                var status = GetToolsStatus();
                (ctrl ?? GetControlForActiveWindow())?.SetToolStatus(status.solverEnabled, status.analysisEnabled);
            }
            catch { }
        }
    }
}
