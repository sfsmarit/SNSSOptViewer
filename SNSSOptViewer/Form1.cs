
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace SNSSOptViewer
{
    public partial class Form1 : Form
    {
        private string _scanRoot = "";

        private TableLayoutPanel _root;
        private TableLayoutPanel _topRows;

        // Row-1 (Refresh + Folder)
        private FlowLayoutPanel _barRow1;
        private Button _btnRefresh;
        private Label _lblFolderPath;
        private Label _lblDropHint;

        // Row-2 (Pair selector + Prev/Next + FileName)
        private FlowLayoutPanel _barRow2;
        private ComboBox _cmbPairs;
        private Button _btnPrevPair;   // older (newest-first list)
        private Button _btnNextPair;   // newer (newest-first list)
        private Label _lblPaths;

        private ToolTip _tip;

        private TabControl _tabs;
        private TabPage _tabVariable;
        private TabPage _tabGoal;
        private StateTabControl _stateCtrl;
        private CharTabControl _charCtrl;

        private string _currentStatePath = "";
        private string _currentCharPath = "";

        private bool _suppressPairChange = false;

        private class PairItem
        {
            public string BaseName;
            public string StatePath;
            public string CharPath;
            public string Short;
            public override string ToString() => Short;
        }

        public Form1() : this("") { }

        public Form1(string initialRoot)
        {
            var asm = Assembly.GetExecutingAssembly();
            using (var s = asm.GetManifestResourceStream("SNSSOptViewer.cockatiel.ico"))
            {
                this.Icon = new Icon(s);
            }

            Text = "SNSS Optimization Log Viewer";
            Width = 900;
            Height = 620;
            MinimumSize = new Size(820, 580);
            StartPosition = FormStartPosition.CenterScreen;

            if (!string.IsNullOrWhiteSpace(initialRoot) && Directory.Exists(initialRoot))
                _scanRoot = initialRoot;

            BuildUI();

            UpdateFolderLabel();
            if (!string.IsNullOrEmpty(_scanRoot))
                PopulatePairs();
        }

        private void BuildUI()
        {
            _root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2 };
            _root.RowStyles.Add(new RowStyle(SizeType.Absolute, 70));
            _root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            Controls.Add(_root);

            _topRows = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Padding = new Padding(8, 6, 8, 4)
            };
            _topRows.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));
            _topRows.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));
            _root.Controls.Add(_topRows, 0, 0);

            _tip = new ToolTip();

            // Row-1
            _barRow1 = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(0)
            };

            _barRow1.AllowDrop = true;
            _barRow1.DragEnter += TopBar_DragEnter;
            _barRow1.DragDrop += TopBar_DragDrop;

            this.AllowDrop = true;
            this.DragEnter += TopBar_DragEnter;
            this.DragDrop += TopBar_DragDrop;

            _btnRefresh = new Button
            {
                Text = "Refresh",
                AutoSize = true,
                Margin = new Padding(0, 2, 8, 0)
            };
            _btnRefresh.Click += (s, e) => PopulatePairs();

            _lblFolderPath = new Label
            {
                AutoSize = false,
                Width = 560,
                Height = 22,
                AutoEllipsis = true,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 4, 8, 0)
            };
            _tip.SetToolTip(_lblFolderPath, "Current folder (full path)");

            _lblDropHint = new Label
            {
                Text = "Drag & drop a folder here",
                AutoSize = true,
                ForeColor = Color.DimGray,
                Font = new Font(SystemFonts.DefaultFont, FontStyle.Italic),
                Margin = new Padding(0, 6, 0, 0),
                Visible = false
            };

            _barRow1.Controls.AddRange(new Control[] { _btnRefresh, _lblFolderPath, _lblDropHint });
            _topRows.Controls.Add(_barRow1, 0, 0);

            // Row-2 (ComboBox, then Prev/Next adjacent, then file label)
            _barRow2 = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(0)
            };

            _cmbPairs = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 260,
                FormattingEnabled = true,
                Margin = new Padding(0, 2, 8, 0)
            };
            _cmbPairs.SelectedIndexChanged += (s, e) =>
            {
                if (_suppressPairChange) return;
                LoadSelectedPair();
            };

            _btnPrevPair = new Button
            {
                Text = "◀ Prev",
                Width = 64,
                Margin = new Padding(0, 2, 2, 0),
                Enabled = false
            };
            // newest-first => Prev means older
            _btnPrevPair.Click += (s, e) =>
            {
                int i = _cmbPairs.SelectedIndex + 1;
                if (i < _cmbPairs.Items.Count) _cmbPairs.SelectedIndex = i;
            };

            _btnNextPair = new Button
            {
                Text = "Next ▶",
                Width = 64,
                Margin = new Padding(0, 2, 12, 0),
                Enabled = false
            };
            // newest-first => Next means newer
            _btnNextPair.Click += (s, e) =>
            {
                int i = _cmbPairs.SelectedIndex - 1;
                if (i >= 0) _cmbPairs.SelectedIndex = i;
            };

            _lblPaths = new Label
            {
                AutoSize = false,
                Width = 500,
                Height = 22,
                AutoEllipsis = true,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 4, 0, 0)
            };
            _tip.SetToolTip(_lblPaths, "Current file (full path)");

            _barRow2.Controls.AddRange(new Control[] { _cmbPairs, _btnPrevPair, _btnNextPair, _lblPaths });
            _topRows.Controls.Add(_barRow2, 0, 1);

            _topRows.Resize += (s, e) => UpdateTopBarsWidth();

            // Tabs
            _tabs = new TabControl { Dock = DockStyle.Fill };
            _tabGoal = new TabPage("Goal");
            _tabVariable = new TabPage("Variable");

            _stateCtrl = new StateTabControl { Dock = DockStyle.Fill };
            _charCtrl = new CharTabControl { Dock = DockStyle.Fill };

            _tabGoal.Controls.Add(_charCtrl);
            _tabVariable.Controls.Add(_stateCtrl);

            _tabs.TabPages.Add(_tabGoal);
            _tabs.TabPages.Add(_tabVariable);
            _root.Controls.Add(_tabs, 0, 1);

            _tabs.SelectedIndexChanged += (s, e) => UpdateCurrentFileLabel();
        }

        private void UpdateTopBarsWidth()
        {
            int margin = 24;

            int w1 = _barRow1.ClientSize.Width - _btnRefresh.Width - margin;
            if (w1 < 180) w1 = 180;

            if (_lblFolderPath.Visible) _lblFolderPath.Width = w1;
            if (_lblDropHint.Visible) _lblFolderPath.Width = 0;

            int w2 = _barRow2.ClientSize.Width - _cmbPairs.Width - _btnPrevPair.Width - _btnNextPair.Width - margin;
            if (w2 < 180) w2 = 180;
            _lblPaths.Width = w2;
        }

        private void UpdateFolderLabel()
        {
            string shownFolder = (Directory.Exists(_scanRoot) ? _scanRoot : "");
            bool hasFolder = !string.IsNullOrWhiteSpace(shownFolder);

            _lblFolderPath.Text = shownFolder;
            _tip.SetToolTip(_lblFolderPath, shownFolder);

            _lblFolderPath.Visible = hasFolder;
            _lblDropHint.Visible = !hasFolder;

            UpdateTopBarsWidth();
        }

        private void UpdatePairNavButtons()
        {
            bool ok = (_cmbPairs != null && _cmbPairs.Items.Count > 0 && _cmbPairs.SelectedIndex >= 0);
            // newest-first:
            // older exists when SelectedIndex < Count-1
            // newer exists when SelectedIndex > 0
            _btnPrevPair.Enabled = ok && _cmbPairs.SelectedIndex < _cmbPairs.Items.Count - 1;
            _btnNextPair.Enabled = ok && _cmbPairs.SelectedIndex > 0;
        }

        private void TopBar_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data != null && e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var paths = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (paths != null && paths.Length > 0)
                {
                    string p = paths[0];
                    if (Directory.Exists(p) || File.Exists(p))
                    {
                        e.Effect = DragDropEffects.Link;
                        return;
                    }
                }
            }
            e.Effect = DragDropEffects.None;
        }

        private void TopBar_DragDrop(object sender, DragEventArgs e)
        {
            try
            {
                if (e.Data == null || !e.Data.GetDataPresent(DataFormats.FileDrop)) return;
                var paths = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (paths == null || paths.Length == 0) return;

                string first = paths[0];
                string newRoot = null;
                if (Directory.Exists(first)) newRoot = first;
                else if (File.Exists(first)) newRoot = Path.GetDirectoryName(first);

                if (!string.IsNullOrEmpty(newRoot) && Directory.Exists(newRoot))
                {
                    _scanRoot = newRoot;
                    UpdateFolderLabel();
                    PopulatePairs();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "Drag & Drop error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static string ShortBeforeAtAt(string name)
        {
            if (string.IsNullOrEmpty(name)) return name;
            int i = name.IndexOf("@@", StringComparison.Ordinal);
            return (i >= 0) ? name.Substring(0, i) : name;
        }

        // Pairing: accept *_state(tag).xls and pair with *_char(tag).xls
        private void PopulatePairs()
        {
            try
            {
                if (!Directory.Exists(_scanRoot))
                {
                    _suppressPairChange = true;
                    _cmbPairs.DataSource = null;
                    _cmbPairs.Items.Clear();
                    _cmbPairs.Items.Add("(folder not found)");
                    _cmbPairs.SelectedIndex = 0;
                    _suppressPairChange = false;

                    _lblPaths.Text = "";
                    _tip.SetToolTip(_lblPaths, "");
                    UpdatePairNavButtons();
                    return;
                }

                var rxState = new Regex(@"^(?<base>.+)_state\((?<tag>[^)]+)\)\.xls$", RegexOptions.IgnoreCase);

                var stateFiles = Directory.EnumerateFiles(_scanRoot, "*_state(*).xls", SearchOption.TopDirectoryOnly)
                    .Select(p => new FileInfo(p))
                    .Select(fi => new { fi, name = (Path.GetFileName(fi.FullName) ?? "") })
                    .Select(x => new { x.fi, m = rxState.Match(x.name) })
                    .Where(x => x.m.Success)
                    .OrderByDescending(x => x.fi.LastWriteTime) // newest first
                    .ToList();

                var pairs = new List<PairItem>();
                foreach (var x in stateFiles)
                {
                    string baseNoSuffix = x.m.Groups["base"].Value;
                    string tag = x.m.Groups["tag"].Value;

                    string charFileName = baseNoSuffix + "_char(" + tag + ").xls";
                    string charPath = Path.Combine(x.fi.DirectoryName, charFileName);

                    if (File.Exists(charPath))
                    {
                        pairs.Add(new PairItem
                        {
                            BaseName = baseNoSuffix,
                            StatePath = x.fi.FullName,
                            CharPath = charPath,
                            Short = ShortBeforeAtAt(baseNoSuffix)
                        });
                    }
                }

                if (pairs.Count == 0)
                {
                    _suppressPairChange = true;
                    _cmbPairs.DataSource = null;
                    _cmbPairs.Items.Clear();
                    _cmbPairs.Items.Add("(no paired *_state(*).xls & *_char(*).xls found)");
                    _cmbPairs.SelectedIndex = 0;
                    _suppressPairChange = false;

                    _lblPaths.Text = "";
                    _tip.SetToolTip(_lblPaths, "");
                    UpdatePairNavButtons();
                    return;
                }

                _suppressPairChange = true;
                _cmbPairs.DataSource = pairs;
                _cmbPairs.DisplayMember = nameof(PairItem.Short);
                _cmbPairs.SelectedIndex = 0;
                _suppressPairChange = false;

                LoadSelectedPair();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "Scan error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadSelectedPair()
        {
            if (!(_cmbPairs.SelectedItem is PairItem it))
            {
                UpdatePairNavButtons();
                return;
            }

            _currentStatePath = it.StatePath;
            _currentCharPath = it.CharPath;

            _stateCtrl.LoadStateFile(_currentStatePath);
            _charCtrl.LoadCharFile(_currentCharPath);

            Text = "SNSS Optimization Log Viewer";
            UpdateCurrentFileLabel();
            UpdatePairNavButtons();
        }

        private void UpdateCurrentFileLabel()
        {
            bool goalActive = (_tabs.SelectedTab == _tabGoal);
            string shownPath = goalActive ? _currentCharPath : _currentStatePath;
            string fileName = (string.IsNullOrEmpty(shownPath) ? "" : Path.GetFileName(shownPath));

            _lblPaths.Text = fileName;
            _tip.SetToolTip(_lblPaths, shownPath);
            UpdateTopBarsWidth();
        }

        // =====================================================================
        // Shared read helpers (ReadOnly + FileShare.ReadWrite)
        // =====================================================================
        private static string[] ReadAllLinesShared(string path)
        {
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (var sr = new StreamReader(fs))
            {
                string text = sr.ReadToEnd();
                bool endsWithNewline = text.EndsWith("\n") || text.EndsWith("\r");

                text = text.Replace("\r\n", "\n").Replace("\r", "\n");
                var lines = text.Split(new[] { "\n" }, StringSplitOptions.None);

                if (!endsWithNewline && lines.Length > 0)
                    Array.Resize(ref lines, lines.Length - 1);

                return lines;
            }
        }

        private static string[] ReadAllLinesSharedRetry(string path, int retries = 3, int delayMs = 150)
        {
            Exception last = null;
            for (int i = 0; i < retries; i++)
            {
                try
                {
                    return ReadAllLinesShared(path);
                }
                catch (IOException ex)
                {
                    last = ex;
                    Thread.Sleep(delayMs);
                }
                catch (UnauthorizedAccessException ex)
                {
                    last = ex;
                    Thread.Sleep(delayMs);
                }
            }
            if (last != null) throw last;
            return Array.Empty<string>();
        }

        // =====================================================================
        // Variable tab (state viewer)
        // =====================================================================
        private class StateTabControl : UserControl
        {
            private const float TL_ChartRowPct = 64f;
            private const int TL_CtrlRowAbs = 34;
            private const float TL_RankRowPct = 36f;
            private const float CA_Left = 5f;
            private const float CA_Width = 90f;
            private const float CA_MainTop = 4f;
            private const float CA_MainHeight = 40f;
            private const float CA_VarsTop = 46f;
            private const float CA_VarsHeight = 50f;

            private TableLayoutPanel _rows;
            private Chart _chart;

            private TableLayoutPanel _ctrl;
            private Label _lblTopN;
            private NumericUpDown _nudTopN;

            private TableLayoutPanel _rank;
            private DataGridView _gridTop;
            private DataGridView _gridWorst;

            private DataTable _table;
            private string _ne = "Neval";
            private string _e0 = "ERR0";
            private string _tt = "T";

            private string _hoverSeriesName = null;

            public StateTabControl()
            {
                BuildUI();
            }

            private void BuildUI()
            {
                _rows = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3 };
                _rows.RowStyles.Add(new RowStyle(SizeType.Percent, TL_ChartRowPct));
                _rows.RowStyles.Add(new RowStyle(SizeType.Absolute, TL_CtrlRowAbs));
                _rows.RowStyles.Add(new RowStyle(SizeType.Percent, TL_RankRowPct));
                Controls.Add(_rows);

                _chart = new Chart { Dock = DockStyle.Fill };
                var main = new ChartArea("main");
                main.AxisX.Title = "";
                main.AxisY.Title = "Error";
                main.AxisY2.Title = "Temperature";
                main.AxisY2.Enabled = AxisEnabled.True;
                main.AxisX.MajorGrid.Enabled = true;
                main.AxisY.MajorGrid.Enabled = true;
                main.AxisX.MajorGrid.LineColor = Color.FromArgb(30, 0, 0, 0);
                main.AxisY.MajorGrid.LineColor = Color.FromArgb(30, 0, 0, 0);
                main.AxisX.MajorGrid.LineDashStyle = ChartDashStyle.Dot;
                main.AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dot;
                main.AxisX.MajorGrid.LineWidth = 1;
                main.AxisY.MajorGrid.LineWidth = 1;
                main.AxisX.MajorTickMark.Enabled = false;
                main.AxisY.MajorTickMark.Enabled = false;
                main.AxisX.Minimum = 0;
                main.AxisX.IsMarginVisible = false;
                main.Position = new ElementPosition(CA_Left, CA_MainTop, CA_Width, CA_MainHeight);

                var vars = new ChartArea("vars");
                vars.AxisX.Title = "Neval";
                vars.AxisY.Title = "Normalized";
                vars.AxisX.MajorGrid.LineColor = Color.Gainsboro;
                vars.AxisY.MajorGrid.LineColor = Color.Gainsboro;
                vars.AxisX.Minimum = 0;
                vars.AxisX.IsMarginVisible = false;
                vars.Position = new ElementPosition(CA_Left, CA_VarsTop, CA_Width, CA_VarsHeight);

                _chart.ChartAreas.Add(main);
                _chart.ChartAreas.Add(vars);

                var legend = new Legend
                {
                    DockedToChartArea = "main",
                    IsDockedInsideChartArea = true,
                    Docking = Docking.Bottom,
                    Alignment = StringAlignment.Near,
                    BackColor = Color.White,
                    BorderColor = Color.Silver,
                    Name = "legend"
                };
                _chart.Legends.Add(legend);

                _chart.MouseMove += Chart_MouseMove;
                _chart.MouseLeave += (s, e) => ApplyHoverHighlight(null);

                _rows.Controls.Add(_chart, 0, 0);

                _ctrl = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Padding = new Padding(8, 2, 8, 2) };
                _ctrl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
                _ctrl.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

                var flp = new FlowLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    FlowDirection = FlowDirection.LeftToRight,
                    WrapContents = false,
                    AutoSize = true,
                    AutoSizeMode = AutoSizeMode.GrowAndShrink,
                    Margin = new Padding(0)
                };
                _lblTopN = new Label { Text = "TopN:", AutoSize = true, Margin = new Padding(0, 6, 4, 0) };
                _nudTopN = new NumericUpDown { Minimum = 1, Maximum = 200, Value = 10, Width = 68, Margin = new Padding(0, 2, 0, 0) };
                flp.Controls.AddRange(new Control[] { _lblTopN, _nudTopN });

                _ctrl.Controls.Add(new Panel { Dock = DockStyle.Fill }, 0, 0);
                _ctrl.Controls.Add(flp, 1, 0);
                _rows.Controls.Add(_ctrl, 0, 1);

                _nudTopN.ValueChanged += (s, e) => Redraw();

                _rank = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
                _rank.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
                _rank.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
                _rank.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

                _gridTop = new DataGridView
                {
                    Dock = DockStyle.Fill,
                    ReadOnly = true,
                    AllowUserToAddRows = false,
                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
                    RowHeadersVisible = false
                };
                _gridWorst = new DataGridView
                {
                    Dock = DockStyle.Fill,
                    ReadOnly = true,
                    AllowUserToAddRows = false,
                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
                    RowHeadersVisible = false
                };

                _rank.Controls.Add(_gridTop, 0, 0);
                _rank.Controls.Add(_gridWorst, 1, 0);
                _rows.Controls.Add(_rank, 0, 2);
            }

            public void LoadStateFile(string path)
            {
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return;

                _table = LoadAsNumericTable(path);
                DetectCoreColumns();

                var varsAll = GetVariableColumns();
                int count = Math.Max(varsAll.Length, 1);

                _nudTopN.Maximum = Math.Max(200, count);
                _nudTopN.Value = Math.Min(_nudTopN.Value, count);

                Redraw();
            }

            private void Redraw()
            {
                if (_table == null) return;

                _chart.Series.Clear();
                PlotMainErr0T();
                PlotVars();
                UpdateRanking();

                var caMain = _chart.ChartAreas["main"];
                var caVars = _chart.ChartAreas["vars"];
                caMain.AxisX.Minimum = 0; caMain.AxisX.IsMarginVisible = false;
                caVars.AxisX.Minimum = 0; caVars.AxisX.IsMarginVisible = false;
            }

            private static DataTable LoadAsNumericTable(string path)
            {
                string[] lines = Form1.ReadAllLinesSharedRetry(path);
                if (lines.Length == 0) throw new InvalidOperationException("Empty file.");

                string header = lines.FirstOrDefault(l => !string.IsNullOrWhiteSpace(l));
                if (header == null) throw new InvalidOperationException("Empty file.");
                header = header.Trim();

                char delim = header.Contains("\t") ? '\t' :
                             header.Contains(",") ? ',' :
                             header.Contains(";") ? ';' : ' ';

                string[] rawHeaders = header.Split(delim).Select(h => h.Trim()).ToArray();
                string[] headers = MakeUnique(rawHeaders);

                DataTable t = new DataTable();
                foreach (string h in headers) t.Columns.Add(h, typeof(double));

                foreach (string raw in lines.Skip(1))
                {
                    if (string.IsNullOrWhiteSpace(raw)) continue;

                    string line = raw.Trim().Trim('\\');
                    string[] tokens = line.Split(delim);
                    if (tokens.Length != headers.Length) tokens = Regex.Split(line, "\\s+");
                    if (tokens.Length != headers.Length) continue;

                    DataRow r = t.NewRow();
                    for (int i = 0; i < headers.Length; i++)
                    {
                        double.TryParse(tokens[i], NumberStyles.Float, CultureInfo.InvariantCulture, out double v);
                        r[i] = v;
                    }
                    t.Rows.Add(r);
                }
                return t;
            }

            private static string[] MakeUnique(string[] names)
            {
                string[] result = new string[names.Length];
                var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                for (int i = 0; i < names.Length; i++)
                {
                    string baseName = string.IsNullOrWhiteSpace(names[i]) ? "Col" : names[i];
                    string name = baseName;
                    int k = 1;
                    while (used.Contains(name)) name = baseName + "_" + (++k).ToString();
                    result[i] = name;
                    used.Add(name);
                }
                return result;
            }

            private void DetectCoreColumns()
            {
                if (!_table.Columns.Contains(_ne))
                    _ne = FindIgnoreCase(_table, "Neval") ?? _table.Columns[0].ColumnName;

                if (!_table.Columns.Contains(_e0))
                    _e0 = FindIgnoreCase(_table, "ERR0")
                          ?? _table.Columns.Cast<DataColumn>().Select(c => c.ColumnName).First(n => n.ToUpper().Contains("ERR0"));

                if (!_table.Columns.Contains(_tt))
                    _tt = FindIgnoreCase(_table, "T") ?? _tt;
            }

            private static string FindIgnoreCase(DataTable t, string name)
            {
                foreach (DataColumn c in t.Columns)
                    if (c.ColumnName.Equals(name, StringComparison.OrdinalIgnoreCase)) return c.ColumnName;
                return null;
            }

            private void PlotMainErr0T()
            {
                double[] x = Col(_ne);
                double[] e0 = Col(_e0).Select(v => -v).ToArray();

                var sErr = new Series("Error")
                {
                    ChartType = SeriesChartType.Line,
                    BorderWidth = 2,
                    ChartArea = "main"
                };
                for (int i = 0; i < x.Length; i++) sErr.Points.AddXY(x[i], e0[i]);
                _chart.Series.Add(sErr);

                if (_table.Columns.Contains(_tt))
                {
                    double[] t = Col(_tt);
                    var sT = new Series("Temperature")
                    {
                        ChartType = SeriesChartType.Line,
                        BorderWidth = 2,
                        Color = Color.DarkOrange,
                        ChartArea = "main",
                        YAxisType = AxisType.Secondary
                    };
                    for (int i = 0; i < x.Length; i++) sT.Points.AddXY(x[i], t[i]);
                    _chart.Series.Add(sT);
                }
            }

            private string[] GetVariableColumns()
            {
                int neIdx = _table.Columns.IndexOf(_ne);
                if (neIdx < 0) return new string[0];

                int pipeIdx = _table.Columns.IndexOf("|");
                if (pipeIdx < 0) pipeIdx = _table.Columns.Count;

                int start = neIdx + 1;
                int end = pipeIdx - 1;
                if (start > end) return new string[0];

                return _table.Columns.Cast<DataColumn>()
                        .Where(c =>
                        {
                            int idx = _table.Columns.IndexOf(c.ColumnName);
                            return idx >= start && idx <= end;
                        })
                        .Select(c => c.ColumnName)
                        .ToArray();
            }

            private void PlotVars()
            {
                string[] candidates = GetVariableColumns();
                if (candidates.Length == 0) return;

                var ranked = candidates
                    .Select(name =>
                    {
                        double[] a = Col(name);
                        double[] f = a.Where(v => !double.IsNaN(v)).ToArray();
                        if (f.Length == 0) return new { name, deltaPct = double.NegativeInfinity, first = double.NaN };

                        double first = f[0];
                        double denom = Math.Max(Math.Abs(first), 1e-12);
                        double deltaPct = 100.0 * (f.Max() - f.Min()) / denom;
                        return new { name, deltaPct, first };
                    })
                    .Where(v => !double.IsNaN(v.first))
                    .OrderByDescending(v => v.deltaPct)
                    .Select(v => v.name)
                    .ToArray();

                int topN = (int)_nudTopN.Value;
                if (topN > ranked.Length) topN = ranked.Length;
                string[] pick = ranked.Take(topN).ToArray();

                double[] x = Col(_ne);

                foreach (string col in pick)
                {
                    double[] raw = Col(col);
                    int idx = Array.FindIndex(raw, v => !double.IsNaN(v));
                    if (idx < 0) continue;

                    double baseVal = Math.Abs(raw[idx]) < 1e-12 ? 1e-12 : raw[idx];
                    double[] y = raw.Select(v => double.IsNaN(v) ? double.NaN : v / baseVal).ToArray();

                    var s = new Series(col)
                    {
                        ChartType = SeriesChartType.Line,
                        ChartArea = "vars",
                        BorderWidth = 1,
                        IsVisibleInLegend = false
                    };

                    for (int i = 0; i < x.Length; i++) s.Points.AddXY(x[i], y[i]);

                    AddEndLabel(s);
                    _chart.Series.Add(s);
                }

                ZoomVarsYAxis();
            }

            private void ZoomVarsYAxis()
            {
                var ca = _chart.ChartAreas["vars"];
                double globalMin = double.PositiveInfinity;
                double globalMax = double.NegativeInfinity;

                foreach (var s in _chart.Series.Cast<Series>().Where(s => s.ChartArea == "vars"))
                {
                    foreach (var p in s.Points)
                    {
                        double y = p.YValues[0];
                        if (!double.IsNaN(y))
                        {
                            if (y < globalMin) globalMin = y;
                            if (y > globalMax) globalMax = y;
                        }
                    }
                }

                if (globalMax > globalMin && !double.IsInfinity(globalMin))
                {
                    double range = globalMax - globalMin;
                    double pad = range * 0.20;
                    if (pad < 0.02) pad = 0.02;

                    ca.AxisY.Minimum = globalMin - pad;
                    ca.AxisY.Maximum = globalMax + pad;
                    ca.AxisY.IntervalAutoMode = IntervalAutoMode.VariableCount;
                }
            }

            private static void AddEndLabel(Series s)
            {
                for (int i = s.Points.Count - 1; i >= 0; i--)
                {
                    var p = s.Points[i];
                    if (!double.IsNaN(p.YValues[0]))
                    {
                        p.Label = s.Name;
                        p.LabelForeColor = Color.DimGray;
                        s.SmartLabelStyle.Enabled = true;
                        s.SmartLabelStyle.AllowOutsidePlotArea = LabelOutsidePlotAreaStyle.Yes;
                        return;
                    }
                }
            }

            private void Chart_MouseMove(object sender, MouseEventArgs e)
            {
                HitTestResult hit = _chart.HitTest(e.X, e.Y, ChartElementType.DataPoint);
                string newHover = null;
                if (hit != null && hit.Series != null && hit.Series.ChartArea == "vars") newHover = hit.Series.Name;
                if (newHover != _hoverSeriesName) ApplyHoverHighlight(newHover);
            }

            private void ApplyHoverHighlight(string seriesName)
            {
                _hoverSeriesName = seriesName;

                var varSeries = _chart.Series.Cast<Series>().Where(s => s.ChartArea == "vars").ToList();
                if (varSeries.Count == 0) return;

                foreach (var s in varSeries)
                {
                    bool isHover = (seriesName != null && s.Name == seriesName);
                    s.BorderWidth = isHover ? 3 : 1;

                    Color c = s.Color;
                    s.Color = isHover ? Color.FromArgb(255, c.R, c.G, c.B) : Color.FromArgb(120, c.R, c.G, c.B);

                    s.ShadowOffset = isHover ? 1 : 0;
                    s.ShadowColor = isHover ? Color.Gray : Color.Empty;

                    var last = GetLastValidPoint(s);
                    if (last != null)
                    {
                        last.LabelForeColor = isHover ? Color.Black : Color.DimGray;
                        last.Font = isHover ? new Font(SystemFonts.DefaultFont, FontStyle.Bold)
                                            : new Font(SystemFonts.DefaultFont, FontStyle.Regular);
                    }
                }

                if (seriesName != null)
                {
                    var hovered = varSeries.FirstOrDefault(s => s.Name == seriesName);
                    if (hovered != null)
                    {
                        foreach (var s in varSeries) _chart.Series.Remove(s);
                        foreach (var s in varSeries) if (!ReferenceEquals(s, hovered)) _chart.Series.Add(s);
                        _chart.Series.Add(hovered);
                    }
                }
            }

            private DataPoint GetLastValidPoint(Series s)
            {
                for (int i = s.Points.Count - 1; i >= 0; i--)
                {
                    var p = s.Points[i];
                    if (!double.IsNaN(p.YValues[0])) return p;
                }
                return null;
            }

            private class VarStat
            {
                public string Name;
                public double First;
                public double Last;
                public double Min;
                public double Max;
                public double Range;
                public double DeltaPct;
                public bool Valid;
                public int Rank;
            }

            private void UpdateRanking()
            {
                string[] candidates = GetVariableColumns();
                if (candidates.Length == 0) { _gridTop.DataSource = null; _gridWorst.DataSource = null; return; }

                var stats = candidates
                    .Select(name =>
                    {
                        double[] a = Col(name);
                        double[] f = a.Where(v => !double.IsNaN(v)).ToArray();
                        if (f.Length == 0) return new VarStat { Name = name, Valid = false };

                        double first = f[0];
                        double last = f[f.Length - 1];
                        double min = f.Min();
                        double max = f.Max();
                        double range = max - min;
                        double denom = Math.Max(Math.Abs(first), 1e-12);
                        double delta = 100.0 * range / denom;

                        return new VarStat
                        {
                            Name = name,
                            Valid = true,
                            First = first,
                            Last = last,
                            Min = min,
                            Max = max,
                            Range = range,
                            DeltaPct = delta
                        };
                    })
                    .Where(s => s.Valid)
                    .ToList();

                var ordered = stats.OrderByDescending(s => s.DeltaPct).ToList();
                for (int i = 0; i < ordered.Count; i++) ordered[i].Rank = i + 1;

                int topN = (int)_nudTopN.Value;
                if (topN > ordered.Count) topN = ordered.Count;

                var listTop = ordered.Take(topN).OrderBy(s => s.Rank).ToList();
                var listWorst = ordered.Skip(Math.Max(0, ordered.Count - topN)).OrderByDescending(s => s.Rank).ToList();

                _gridTop.DataSource = BuildRankTable(listTop);
                _gridWorst.DataSource = BuildRankTable(listWorst);

                ApplyGridFormat(_gridTop);
                ApplyGridFormat(_gridWorst);
            }

            private static DataTable BuildRankTable(List<VarStat> src)
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("Rank", typeof(int));
                dt.Columns.Add("Name", typeof(string));
                dt.Columns.Add("First", typeof(double));
                dt.Columns.Add("Last", typeof(double));
                dt.Columns.Add("Min", typeof(double));
                dt.Columns.Add("Max", typeof(double));
                dt.Columns.Add("Delta", typeof(double));

                foreach (var s in src) dt.Rows.Add(s.Rank, s.Name, s.First, s.Last, s.Min, s.Max, s.DeltaPct / 100.0);
                return dt;
            }

            private static void ApplyGridFormat(DataGridView g)
            {
                var nameCol = g.Columns["Name"];
                if (nameCol != null)
                {
                    nameCol.MinimumWidth = 100;
                    nameCol.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                }

                if (g.Columns["Delta"] != null) g.Columns["Delta"].DefaultCellStyle.Format = "P1";
                if (g.Columns["First"] != null) g.Columns["First"].DefaultCellStyle.Format = "G6";
                if (g.Columns["Last"] != null) g.Columns["Last"].DefaultCellStyle.Format = "G6";
                if (g.Columns["Min"] != null) g.Columns["Min"].DefaultCellStyle.Format = "G6";
                if (g.Columns["Max"] != null) g.Columns["Max"].DefaultCellStyle.Format = "G6";

                foreach (DataGridViewColumn c in g.Columns) c.SortMode = DataGridViewColumnSortMode.Automatic;
            }

            private double[] Col(string col)
            {
                int n = _table.Rows.Count;
                double[] a = new double[n];
                for (int i = 0; i < n; i++)
                {
                    object v = _table.Rows[i][col];
                    a[i] = (v is double d) ? d : double.NaN;
                }
                return a;
            }
        }

        // =====================================================================
        // Goal tab (char viewer)
        // =====================================================================
        private class CharTabControl : UserControl
        {
            private TableLayoutPanel root;

            private FlowLayoutPanel bar;
            private ComboBox cmbSolution;
            private Button btnPrevSol, btnNextSol;
            private Label lblTopN;
            private NumericUpDown numTopN;

            private TableLayoutPanel content;
            private Chart chart1;
            private TableLayoutPanel gridLayout;
            private DataGridView grid1, grid2, grid3;

            private string _currentFile = "";
            private string[] _lines = Array.Empty<string>();
            private int _headerRow = -1;
            private char _sep = '\t';

            private class ErrCol { public int Index; public string Header; }
            private List<ErrCol> _errCols = new List<ErrCol>();

            private class SolutionEntry { public int Number; public string Token; public int RowIndex; }
            private List<SolutionEntry> _solutions = new List<SolutionEntry>();

            private class Row
            {
                public int Rank { get; set; }
                public string Name { get; set; }
                public double Value { get; set; }
                public double Share { get; set; }
            }
            private List<Row> _allRowsCurrent = new List<Row>();

            public CharTabControl()
            {
                BuildUI();
            }

            private void BuildUI()
            {
                root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, Padding = new Padding(8) };
                root.RowStyles.Add(new RowStyle(SizeType.Absolute, 56f));
                root.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
                root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
                Controls.Add(root);

                // Top bar (Prev/Next adjacent, TopN label+spin adjacent)
                bar = new FlowLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    FlowDirection = FlowDirection.LeftToRight,
                    WrapContents = false,
                    Padding = new Padding(4),
                    Margin = new Padding(0)
                };
                root.Controls.Add(bar, 0, 0);

                cmbSolution = new ComboBox
                {
                    DropDownStyle = ComboBoxStyle.DropDownList,
                    Width = 220,
                    Margin = new Padding(0, 2, 8, 0)
                };

                btnPrevSol = new Button { Text = "◀ Prev", Width = 64, Margin = new Padding(0, 2, 2, 0), Enabled = false };
                btnPrevSol.Click += (s, e) =>
                {
                    int i = cmbSolution.SelectedIndex - 1;
                    if (i >= 0) cmbSolution.SelectedIndex = i;
                };

                btnNextSol = new Button { Text = "Next ▶", Width = 64, Margin = new Padding(0, 2, 12, 0), Enabled = false };
                btnNextSol.Click += (s, e) =>
                {
                    int i = cmbSolution.SelectedIndex + 1;
                    if (i < cmbSolution.Items.Count) cmbSolution.SelectedIndex = i;
                };

                lblTopN = new Label { Text = "TopN", AutoSize = true, Margin = new Padding(0, 6, 4, 0) };
                numTopN = new NumericUpDown { Minimum = 1, Maximum = 9999, Value = 100, Width = 72, Margin = new Padding(0, 2, 0, 0) };
                numTopN.ValueChanged += (s, e) => ApplyTopNOnly();

                cmbSolution.SelectedIndexChanged += (s, e) =>
                {
                    ApplySelectedSolution();
                    UpdateSolutionNavButtons();
                };

                bar.Controls.AddRange(new Control[] { cmbSolution, btnPrevSol, btnNextSol, lblTopN, numTopN });

                content = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, Padding = new Padding(0) };
                content.RowStyles.Add(new RowStyle(SizeType.Percent, 50f));
                content.RowStyles.Add(new RowStyle(SizeType.Percent, 50f));
                content.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
                root.Controls.Add(content, 0, 1);

                chart1 = new Chart { Dock = DockStyle.Fill, Padding = new Padding(4) };
                var area = new ChartArea("area");
                area.AxisX.Interval = 1;
                area.AxisX.LabelStyle.Angle = -90;
                area.AxisX.Title = "";
                area.AxisY.Title = "Error";
                chart1.ChartAreas.Add(area);
                var series = new Series("ERR")
                {
                    ChartType = SeriesChartType.Column,
                    XValueType = ChartValueType.String,
                    YValueType = ChartValueType.Double
                };
                chart1.Series.Add(series);
                chart1.Titles.Clear();
                content.Controls.Add(chart1, 0, 0);

                gridLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 1, Padding = new Padding(0) };
                gridLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33f));
                gridLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33f));
                gridLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.34f));
                gridLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
                content.Controls.Add(gridLayout, 0, 1);

                grid1 = MakeGrid(); grid2 = MakeGrid(); grid3 = MakeGrid();
                gridLayout.Controls.Add(grid1, 0, 0);
                gridLayout.Controls.Add(grid2, 1, 0);
                gridLayout.Controls.Add(grid3, 2, 0);

                UpdateSolutionNavButtons();
            }

            private void UpdateSolutionNavButtons()
            {
                bool ok = (cmbSolution != null && cmbSolution.Items.Count > 0 && cmbSolution.SelectedIndex >= 0);
                btnPrevSol.Enabled = ok && cmbSolution.SelectedIndex > 0;
                btnNextSol.Enabled = ok && cmbSolution.SelectedIndex < cmbSolution.Items.Count - 1;
            }

            private DataGridView MakeGrid()
            {
                return new DataGridView
                {
                    Dock = DockStyle.Fill,
                    ReadOnly = true,
                    AllowUserToAddRows = false,
                    AllowUserToDeleteRows = false,
                    SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                    AutoGenerateColumns = true,
                    RowHeadersVisible = false,
                    Font = new Font(SystemFonts.MessageBoxFont.FontFamily, 8.5f),
                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                };
            }

            public void LoadCharFile(string charPath)
            {
                btnPrevSol.Enabled = false;
                btnNextSol.Enabled = false;

                try
                {
                    _currentFile = charPath;
                    if (string.IsNullOrWhiteSpace(_currentFile) || !File.Exists(_currentFile))
                    {
                        _errCols.Clear();
                        _solutions.Clear();
                        _allRowsCurrent.Clear();
                        UpdateSolutionNavButtons();
                        return;
                    }

                    _lines = Form1.ReadAllLinesSharedRetry(_currentFile);
                    if (_lines.Length == 0)
                    {
                        _errCols.Clear();
                        _solutions.Clear();
                        _allRowsCurrent.Clear();
                        UpdateSolutionNavButtons();
                        return;
                    }

                    _headerRow = 0;

                    var headerLine = _lines[_headerRow] ?? "";
                    _sep = headerLine.Contains('\t') ? '\t' : (headerLine.Contains(',') ? ',' : '\t');

                    var headers = Split(headerLine, _sep);
                    var headersNorm = headers.Select(NormalizeCell).ToArray();

                    int lastPipeIdx = Array.FindLastIndex(headersNorm, h => h == "|");
                    if (lastPipeIdx < 0)
                    {
                        _errCols.Clear();
                        _solutions.Clear();
                        _allRowsCurrent.Clear();
                        UpdateSolutionNavButtons();

                        MessageBox.Show(this.FindForm(),
                            "Header separator '|' was not found in the first line.",
                            "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    _errCols = headers
                        .Select((h, i) => new { i, n = NormalizeCell(h) })
                        .Select(x => new ErrCol { Index = x.i, Header = x.n })
                        .Where(ec => ec.Index > lastPipeIdx && !string.IsNullOrEmpty(ExtractErrKey(ec.Header)))
                        .ToList();

                    if (_errCols.Count == 0)
                    {
                        _solutions.Clear();
                        _allRowsCurrent.Clear();
                        UpdateSolutionNavButtons();

                        MessageBox.Show(this.FindForm(),
                            "No ERR columns were found to the right of the last '|' in the header.",
                            "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    _solutions.Clear();
                    int firstDataRow = _headerRow + 1;
                    var rx = new Regex(@"^solution[a-z]*\+(\d+)\.txt$", RegexOptions.IgnoreCase);

                    for (int r = firstDataRow; r < _lines.Length; r++)
                    {
                        if (string.IsNullOrWhiteSpace(_lines[r])) continue;
                        var cells = Split(_lines[r], _sep);
                        var token = FindSolutionTokenFromRight(cells);
                        if (string.IsNullOrEmpty(token)) continue;

                        var clean = NormalizeCell(token);
                        var baseName = Path.GetFileName(clean);
                        var m = rx.Match(baseName);
                        if (!m.Success) continue;

                        if (int.TryParse(m.Groups[1].Value, out int num))
                        {
                            var idx = _solutions.FindIndex(s => s.Number == num);
                            if (idx >= 0) _solutions.RemoveAt(idx);
                            _solutions.Add(new SolutionEntry { Number = num, Token = token, RowIndex = r });
                        }
                    }

                    if (_solutions.Count == 0)
                    {
                        _allRowsCurrent.Clear();
                        cmbSolution.DataSource = null;
                        UpdateSolutionNavButtons();

                        MessageBox.Show(this.FindForm(),
                            "No solution+N-like tokens were found after the header.",
                            "Info",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                        return;
                    }

                    _solutions = _solutions.OrderBy(s => s.Number).ToList();
                    cmbSolution.DataSource = _solutions
                        .Select(s => new ComboItem { Display = $"{s.Token}", Number = s.Number, RowIndex = s.RowIndex })
                        .ToList();
                    cmbSolution.DisplayMember = nameof(ComboItem.Display);
                    cmbSolution.ValueMember = nameof(ComboItem.Number);
                    cmbSolution.SelectedIndex = _solutions.Count - 1;

                    ApplySelectedSolution();
                    UpdateSolutionNavButtons();
                }
                catch (Exception ex)
                {
                    UpdateSolutionNavButtons();
                    MessageBox.Show(this.FindForm(), ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            private void ApplySelectedSolution()
            {
                if (cmbSolution.SelectedItem == null || _errCols.Count == 0)
                {
                    _allRowsCurrent.Clear();
                    return;
                }

                var item = (ComboItem)cmbSolution.SelectedItem;
                int rowIndex = item.RowIndex;
                if (rowIndex <= _headerRow || rowIndex >= _lines.Length) { _allRowsCurrent.Clear(); return; }
                if (string.IsNullOrWhiteSpace(_lines[rowIndex])) { _allRowsCurrent.Clear(); return; }

                var cells = Split(_lines[rowIndex], _sep);

                var list = new List<Row>(_errCols.Count);
                foreach (var col in _errCols)
                {
                    double value = 0.0;
                    if (col.Index < cells.Length)
                    {
                        var raw = (cells[col.Index] ?? "").Trim();
                        if (double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out double d)) value = d;
                    }

                    string name = ExtractErrKey(col.Header);
                    if (string.IsNullOrEmpty(name)) continue;

                    list.Add(new Row { Name = name, Value = value });
                }

                _allRowsCurrent = list.OrderByDescending(r => r.Value).ToList();
                ApplyTopNOnly();
            }

            private void ApplyTopNOnly()
            {
                if (_allRowsCurrent == null || _allRowsCurrent.Count == 0)
                {
                    chart1.Series["ERR"].Points.Clear();
                    grid1.DataSource = null;
                    grid2.DataSource = null;
                    grid3.DataSource = null;
                    return;
                }

                int topN = (int)numTopN.Value; if (topN < 1) topN = 1;

                var slice = _allRowsCurrent.Take(topN).ToList();
                double totalAll = _allRowsCurrent.Sum(r => r.Value);
                for (int i = 0; i < slice.Count; i++)
                {
                    slice[i].Rank = i + 1;
                    slice[i].Share = (totalAll > 0) ? (slice[i].Value / totalAll) : 0.0;
                }

                var srs = chart1.Series["ERR"];
                srs.Points.Clear();
                srs.Points.DataBindXY(slice.Select(x => x.Name).ToArray(), slice.Select(x => x.Value).ToArray());

                int per = (int)Math.Ceiling(slice.Count / 3.0);
                var seg1 = slice.Take(per).ToList();
                var seg2 = slice.Skip(per).Take(per).ToList();
                var seg3 = slice.Skip(2 * per).ToList();

                BindGrid(grid1, seg1);
                BindGrid(grid2, seg2);
                BindGrid(grid3, seg3);
            }

            private void BindGrid(DataGridView g, List<Row> rows)
            {
                g.DataSource = rows
                    .Select(x => new Row { Rank = x.Rank, Name = x.Name, Value = x.Value, Share = x.Share })
                    .ToList();

                var c = g.Columns[nameof(Row.Share)];
                if (c != null) c.DefaultCellStyle.Format = "P1";
            }

            private static string[] Split(string text, char sep) => (text ?? "").Split(new[] { sep }, StringSplitOptions.None);
            private static string NormalizeCell(string s) => (s ?? "").Trim().Trim('"', '\'');

            private static string ExtractErrKey(string headerNorm)
            {
                if (string.IsNullOrEmpty(headerNorm)) return "";

                var m = Regex.Match(headerNorm, @"^\(([^)]+)\)(?:\*.*)?$", RegexOptions.IgnoreCase);
                string tok = m.Success ? m.Groups[1].Value : headerNorm;

                tok = (tok ?? "").Trim();
                if (tok.Length == 0) return "";

                if (tok.StartsWith("ERR_", StringComparison.OrdinalIgnoreCase))
                    return tok.Substring(4);

                if (m.Success)
                    return tok;

                return "";
            }

            private static string FindSolutionTokenFromRight(string[] tokens)
            {
                if (tokens == null || tokens.Length == 0) return null;

                var normalized = tokens.Select(t => (t ?? "").Trim().Trim('"', '\'')).ToArray();
                for (int i = normalized.Length - 1; i >= 0; i--)
                {
                    var tk = normalized[i];
                    if (tk.Length == 0) continue;
                    if (tk.IndexOf("solution", StringComparison.OrdinalIgnoreCase) >= 0 &&
                        tk.Contains("+") &&
                        tk.EndsWith(".txt", StringComparison.OrdinalIgnoreCase))
                        return tk;
                }
                return null;
            }

            private class ComboItem { public string Display { get; set; } public int Number { get; set; } public int RowIndex { get; set; } }
        }
    }
}
