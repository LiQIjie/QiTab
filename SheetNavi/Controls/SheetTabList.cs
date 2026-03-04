using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace SheetNavi.Controls
{
    // A custom vertical sheet tab list control that mimics Excel tab style
    // Supports Ctrl/Shift multi-select, group headers with +/- for expand/collapse,
    // double-click to expand/collapse group, double-click sheet to rename via host,
    // and two quick clicks on group name to trigger rename.
    public class SheetTabList : UserControl
    {
        public event Action<string> OnSheetSelected;
        public event Action<List<string>> GroupRequested;
        public event Action<List<string>> UngroupRequested;
        public event Action<List<string>> RenameGroupRequested;
        public event Action<List<string>, Color> ColorRequested;
        public event Action<string> RenameSheetRequested;
        public event Action<string, string> GroupRenameCommitted;

        public event Action<List<string>> MoveSheetsUpRequested;
        public event Action<List<string>> MoveSheetsDownRequested;
        public event Action<string> MoveGroupUpRequested;
        public event Action<string> MoveGroupDownRequested;

        private ContextMenuStrip _ctx = new ContextMenuStrip();

        private class GroupModel
        {
            public string Name;
            public Color Color;
            public bool Collapsed;
            public List<string> Members = new List<string>();
        }

        private Dictionary<string, GroupModel> _groups = new Dictionary<string, GroupModel>(StringComparer.OrdinalIgnoreCase);
        private List<string> _groupOrder = new List<string>();
        private List<string> _allSheets = new List<string>();

        private HashSet<string> _selected = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private string _lastSelected;

        private Font _sheetFont;
        private Font _groupFont;

        private Timer _clickTimer = new Timer();
        private int _clickCount = 0;
        private Rectangle _lastClickRect;
        private string _lastClickGroup;

        public SheetTabList()
        {
            DoubleBuffered = true;
            BackColor = SystemColors.Window;
            _sheetFont = new Font(SystemFonts.MessageBoxFont, FontStyle.Regular);
            _groupFont = new Font(SystemFonts.MessageBoxFont, FontStyle.Bold);

            // context menu
            var miGroup = new ToolStripMenuItem("Group");
            var miUnGroup = new ToolStripMenuItem("Ungroup");
            var miRenameGroup = new ToolStripMenuItem("Rename Group");
            var miColor = new ToolStripMenuItem("Set Tab Color...");
            var miMoveUp = new ToolStripMenuItem("Move Up");
            var miMoveDown = new ToolStripMenuItem("Move Down");

            miGroup.Click += (s, e) => GroupRequested?.Invoke(GetSelectedNames());
            miUnGroup.Click += (s, e) => UngroupRequested?.Invoke(GetSelectedNames());
            miRenameGroup.Click += (s, e) => RenameGroupRequested?.Invoke(GetSelectedNames());
            miColor.Click += (s, e) =>
            {
                var sel = GetSelectedNames();
                if (sel.Count == 0) return;
                using (var dlg = new ColorDialog())
                {
                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                        ColorRequested?.Invoke(sel, dlg.Color);
                    }
                }
            };

            miMoveUp.Click += (s, e) =>
            {
                // if selection belongs to a single group and whole group selected, move group
                if (_lastContextHit != null && _lastContextHit.Type == HitType.GroupHeader)
                {
                    MoveGroupUpRequested?.Invoke(_lastContextHit.Group.Name);
                }
                else
                {
                    MoveSheetsUpRequested?.Invoke(GetSelectedNames());
                }
            };
            miMoveDown.Click += (s, e) =>
            {
                if (_lastContextHit != null && _lastContextHit.Type == HitType.GroupHeader)
                {
                    MoveGroupDownRequested?.Invoke(_lastContextHit.Group.Name);
                }
                else
                {
                    MoveSheetsDownRequested?.Invoke(GetSelectedNames());
                }
            };

            _ctx.Items.AddRange(new ToolStripItem[] { miGroup, miUnGroup, miRenameGroup, new ToolStripSeparator(), miColor, new ToolStripSeparator(), miMoveUp, miMoveDown });

            _clickTimer.Interval = 300; // double-click/quick two-click window
            _clickTimer.Tick += (s, e) => { _clickCount = 0; _clickTimer.Stop(); _lastClickGroup = null; };

            MouseDown += OnMouseDown;
            MouseUp += OnMouseUp;
            MouseDoubleClick += OnMouseDoubleClick;
            Resize += (s, e) => Invalidate();
        }

        public void LoadSheets(List<string> names)
        {
            _allSheets = names ?? new List<string>();
            Invalidate();
        }

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

        public void Highlight(string name)
        {
            // select single
            _selected.Clear();
            if (!string.IsNullOrEmpty(name)) _selected.Add(name);
            _lastSelected = name;
            Invalidate();
        }

        public List<string> GetSelectedNames() => _selected.ToList();

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            var g = e.Graphics;
            g.Clear(BackColor);

            int y = 4;
            int pad = 6;
            int arrowSize = 10;
            var sheetHeight = Math.Max(_sheetFont.Height + 8, 22);
            var groupHeight = Math.Max(_groupFont.Height + 8, 24);

            // draw groups headers and members, then ungrouped sheets
            var assigned = new HashSet<string>(_groups.SelectMany(x => x.Value.Members), StringComparer.OrdinalIgnoreCase);

            foreach (var groupName in _groupOrder)
            {
                var gm = _groups[groupName];
                // group header
                var headerRect = new Rectangle(0, y, Width, groupHeight);
                DrawGroupHeader(g, gm, headerRect, arrowSize);
                y += groupHeight;
                if (!gm.Collapsed)
                {
                    foreach (var s in gm.Members)
                    {
                        if (!_allSheets.Contains(s)) continue;
                        var rect = new Rectangle(0, y, Width, sheetHeight);
                        DrawSheetItem(g, s, rect, gm.Color);
                        y += sheetHeight;
                    }
                }
            }

            foreach (var s in _allSheets)
            {
                if (assigned.Contains(s)) continue;
                var rect = new Rectangle(0, y, Width, sheetHeight);
                DrawSheetItem(g, s, rect, Color.Empty);
                y += sheetHeight;
            }
        }

        private void DrawGroupHeader(Graphics g, GroupModel gm, Rectangle rect, int arrowSize)
        {
            // background tint
            if (gm.Color != Color.Empty)
            {
                using (var b = new SolidBrush(Color.FromArgb(200, gm.Color))) g.FillRectangle(b, rect);
            }
            // arrow +/-
            var arrowRect = new Rectangle(rect.X + 6, rect.Y + (rect.Height - arrowSize) / 2, arrowSize, arrowSize);
            using (var p = new Pen(SystemColors.ControlDark, 1))
            {
                g.DrawRectangle(p, arrowRect);
                // draw + or - inside
                using (var p2 = new Pen(SystemColors.ControlDarkDark, 2))
                {
                    // horizontal
                    g.DrawLine(p2, arrowRect.Left + 2, arrowRect.Top + arrowRect.Height / 2, arrowRect.Right - 2, arrowRect.Top + arrowRect.Height / 2);
                    if (gm.Collapsed)
                    {
                        // vertical (plus)
                        g.DrawLine(p2, arrowRect.Left + arrowRect.Width / 2, arrowRect.Top + 2, arrowRect.Left + arrowRect.Width / 2, arrowRect.Bottom - 2);
                    }
                }
            }
            // text "гл group"
            var text = "гл " + gm.Name;
            var textRect = new Rectangle(arrowRect.Right + 6, rect.Y, rect.Width - (arrowRect.Right + 6), rect.Height);
            TextRenderer.DrawText(g, text, _groupFont, textRect, SystemColors.ControlText, TextFormatFlags.VerticalCenter | TextFormatFlags.Left);
        }

        private void DrawSheetItem(Graphics g, string name, Rectangle rect, Color groupColor)
        {
            bool sel = _selected.Contains(name);
            if (sel)
            {
                using (var b = new SolidBrush(SystemColors.Highlight)) g.FillRectangle(b, rect);
            }
            else if (groupColor != Color.Empty)
            {
                using (var b = new SolidBrush(Color.FromArgb(24, groupColor))) g.FillRectangle(b, rect);
            }

            var fore = sel ? SystemColors.HighlightText : SystemColors.ControlText;
            var textRect = new Rectangle(rect.X + 24, rect.Y, rect.Width - 24, rect.Height);
            TextRenderer.DrawText(g, name, _sheetFont, textRect, fore, TextFormatFlags.VerticalCenter | TextFormatFlags.Left);

            // small color bar on left if group has color
            if (groupColor != Color.Empty)
            {
                using (var b = new SolidBrush(groupColor)) g.FillRectangle(b, new Rectangle(rect.X + 8, rect.Y + 4, 4, rect.Height - 8));
            }
        }

        private HitInfo _lastContextHit;

        private void OnMouseDown(object sender, MouseEventArgs e)
        {
            Focus();
            var hit = HitTest(e.Location);
            if (hit == null) return;

            if (hit.Type == HitType.GroupArrow)
            {
                // toggle collapse on arrow click
                var gm = hit.Group;
                gm.Collapsed = !gm.Collapsed;
                Invalidate();
                return;
            }
        }

        private void OnMouseUp(object sender, MouseEventArgs e)
        {
            var hit = HitTest(e.Location);
            _lastContextHit = hit;
            if (hit == null) return;

            if (e.Button == MouseButtons.Right)
            {
                // add to selection if not included
                if (hit.Type == HitType.Sheet)
                {
                    if (!_selected.Contains(hit.Sheet)) _selected.Add(hit.Sheet);
                }
                else if (hit.Type == HitType.GroupHeader)
                {
                    if (hit.Group.Members.Count > 0)
                    {
                        foreach (var s in hit.Group.Members) _selected.Add(s);
                    }
                }
                Invalidate();
                _ctx.Show(this, e.Location);
                return;
            }

            // left click selection logic
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
                // select entire group and activate first sheet
                _selected.Clear();
                foreach (var s in hit.Group.Members) _selected.Add(s);
                _lastSelected = hit.Group.Members.FirstOrDefault();
                if (!string.IsNullOrEmpty(_lastSelected)) OnSheetSelected?.Invoke(_lastSelected);
                Invalidate();

                // quick two-click detection for rename
                if (_lastClickGroup == hit.Group.Name && _clickCount == 1 && _lastClickRect.Contains(e.Location))
                {
                    // trigger rename
                    GroupRenameCommitted?.Invoke(hit.Group.Name, hit.Group.Name); // host will prompt and update
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

        private void OnMouseDoubleClick(object sender, MouseEventArgs e)
        {
            var hit = HitTest(e.Location);
            if (hit == null) return;
            if (hit.Type == HitType.GroupHeader)
            {
                hit.Group.Collapsed = !hit.Group.Collapsed;
                Invalidate();
            }
            else if (hit.Type == HitType.Sheet)
            {
                RenameSheetRequested?.Invoke(hit.Sheet);
            }
        }

        // selection range by visible order
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

        private enum HitType { None, GroupHeader, GroupArrow, Sheet }
        private class HitInfo
        {
            public HitType Type;
            public GroupModel Group;
            public string Sheet;
            public Rectangle Bounds;
        }

        private HitInfo HitTest(Point pt)
        {
            int y = 4;
            int arrowSize = 10;
            var sheetHeight = Math.Max(_sheetFont.Height + 8, 22);
            var groupHeight = Math.Max(_groupFont.Height + 8, 24);

            foreach (var groupName in _groupOrder)
            {
                var gm = _groups[groupName];
                var headerRect = new Rectangle(0, y, Width, groupHeight);
                var arrowRect = new Rectangle(headerRect.X + 6, headerRect.Y + (headerRect.Height - arrowSize) / 2, arrowSize, arrowSize);
                if (arrowRect.Contains(pt)) return new HitInfo { Type = HitType.GroupArrow, Group = gm, Bounds = arrowRect };
                if (headerRect.Contains(pt)) return new HitInfo { Type = HitType.GroupHeader, Group = gm, Bounds = headerRect };
                y += groupHeight;
                if (!gm.Collapsed)
                {
                    foreach (var s in gm.Members)
                    {
                        if (!_allSheets.Contains(s)) continue;
                        var rect = new Rectangle(0, y, Width, sheetHeight);
                        if (rect.Contains(pt)) return new HitInfo { Type = HitType.Sheet, Group = gm, Sheet = s, Bounds = rect };
                        y += sheetHeight;
                    }
                }
            }

            foreach (var s in _allSheets)
            {
                if (_groups.Values.Any(g => g.Members.Contains(s))) continue;
                var rect = new Rectangle(0, y, Width, sheetHeight);
                if (rect.Contains(pt)) return new HitInfo { Type = HitType.Sheet, Sheet = s, Bounds = rect };
                y += sheetHeight;
            }

            return null;
        }
    }
}
