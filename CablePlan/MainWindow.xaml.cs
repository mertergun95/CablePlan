using Microsoft.Win32;
using Ookii.Dialogs.Wpf;
using PdfiumViewer;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfiumDoc = PdfiumViewer.PdfDocument;
using SharpPdfDocument = PdfSharp.Pdf.PdfDocument;

using IOPath = System.IO.Path;
using IOFile = System.IO.File;
using IODir = System.IO.Directory;

namespace CablePlan
{
    public partial class MainWindow : Window
    {
        private static readonly string AppSettingsPath = IOPath.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "CablePlan",
            "appsettings.json");

        private string _workspaceRoot = "";
        private string _workspacePdfDir = "";
        private string _workspaceDataDir = "";
        private string _workspaceMetaPath = "";
        private string _assignmentFilePath = "";

        private bool _isUpdatingCableSelection = false;
        private bool _isUpdatingSperrSelection = false;

        private string _currentPdfPath = "";
        private string _currentPdfBaseName = "";
        private string _planJsonPath = "";

        private BitmapSource? _baseBmp;
        private int _rotationDeg = 0;
        private double _zoom = 1.0;

        private readonly Dictionary<string, int> _pdfRotation = new(StringComparer.OrdinalIgnoreCase);

        private readonly List<Cable> _selectedCables = new();
        private readonly HashSet<string> _activeSperrpauseFilters = new(StringComparer.OrdinalIgnoreCase);
        private readonly List<Point> _draftPointsBase = new();

        // cableId -> assigned sperrpause names
        private Dictionary<string, HashSet<string>> _cableToSperrpauseAssignments =
            new(StringComparer.OrdinalIgnoreCase);

        private bool _midPanning = false;
        private Point _midStart;
        private double _midHStart, _midVStart;

        private bool _pendingLeftPan = false;
        private bool _leftPanActive = false;
        private Point _leftStart;
        private double _leftHStart, _leftVStart;

        private static readonly Brush CableHighlightBrush =
            new SolidColorBrush(Color.FromRgb(34, 76, 152)); // koyu mavi

        private static readonly Brush SperrpauseHighlightBrush =
            new SolidColorBrush(Color.FromRgb(255, 210, 0)); // sarı

        private enum DrawKind
        {
            None,
            Cable,
            Sperrpause
        }

        public class PointD
        {
            public double X { get; set; }
            public double Y { get; set; }
            public PointD() { }
            public PointD(double x, double y) { X = x; Y = y; }
        }

        public class Cable
        {
            public string Id { get; set; } = "";
            public List<PointD> Points { get; set; } = new();
            public PointD LabelPoint { get; set; } = new(0, 0);
        }

        public class Sperrpause
        {
            public string Id { get; set; } = "";
            public List<PointD> Points { get; set; } = new();
            public PointD LabelPoint { get; set; } = new(0, 0);
        }

        public class PlanData
        {
            public List<Cable> Cables { get; set; } = new();
            public List<Sperrpause> Sperrpauses { get; set; } = new();
        }

        public class WorkspaceMeta
        {
            public Dictionary<string, HashSet<string>> CableIndex { get; set; } =
                new(StringComparer.OrdinalIgnoreCase);

            public Dictionary<string, HashSet<string>> SperrpauseIndex { get; set; } =
                new(StringComparer.OrdinalIgnoreCase);
        }

        public class AppSettings
        {
            public string LastWorkspacePath { get; set; } = "";
        }

        private PlanData _data = new();
        private WorkspaceMeta _ws = new();

        public MainWindow()
        {
            InitializeComponent();

            ZoomTf.ScaleX = ZoomTf.ScaleY = _zoom;

            Loaded += (_, __) =>
            {
                AttachCableContextMenuHandlers();
                AttachSperrpauseContextMenuHandlers();
            };

            BtnChooseWorkspace.Click += (_, __) => ChooseWorkspaceFolder();
            BtnOpenWorkspaceFolder.Click += (_, __) => OpenWorkspaceFolder();

            BtnImportPdf.Click += (_, __) => ImportPdf();
            BtnOpenSelectedPdf.Click += (_, __) => OpenSelectedPdfExternally();

            BtnRotateLeft.Click += (_, __) => Rotate(-90);
            BtnRotateRight.Click += (_, __) => Rotate(90);

            BtnFitWidth.Click += (_, __) => FitWidth();
            BtnFitPage.Click += (_, __) => FitPage();
            BtnExportPng.Click += (_, __) => ExportCurrentViewAsPng();
            BtnExportPdf.Click += (_, __) => ExportCurrentViewAsPdf();
            BtnPrintView.Click += (_, __) => PrintCurrentView();

            TglDrawCable.Checked += (_, __) => ActivateDrawMode(DrawKind.Cable);
            TglDrawCable.Unchecked += (_, __) =>
            {
                if (GetCurrentDrawKind() == DrawKind.None) UpdateModeText();
            };

            TglDrawSperrpause.Checked += (_, __) => ActivateDrawMode(DrawKind.Sperrpause);
            TglDrawSperrpause.Unchecked += (_, __) =>
            {
                if (GetCurrentDrawKind() == DrawKind.None) UpdateModeText();
            };

            TglSetLabel.Checked += (_, __) => UpdateModeText();
            TglSetLabel.Unchecked += (_, __) => UpdateModeText();

            ExpDrawPanel.Expanded += (_, __) => UpdateModeText();
            ExpDrawPanel.Collapsed += (_, __) => CollapseDrawTools();

            TglMultiSelect.Checked += (_, __) => ApplyMultiSelectMode();
            TglMultiSelect.Unchecked += (_, __) => ApplyMultiSelectMode();

            TglCableBasedSperrFilter.Checked += (_, __) => RefreshCurrentPdfSperrpauseList();
            TglCableBasedSperrFilter.Unchecked += (_, __) => RefreshCurrentPdfSperrpauseList();
            TglSperrBasedCableFilter.Checked += (_, __) => RefreshCurrentPdfCableList();
            TglSperrBasedCableFilter.Unchecked += (_, __) => RefreshCurrentPdfCableList();

            BtnClearCableSelection.Click += (_, __) => ClearCableSelectionAndRefresh();

            BtnFinishCable.Click += (_, __) => FinishCable();
            BtnDeleteCable.Click += (_, __) => DeleteSelectedOrTypedCable();

            BtnFinishSperrpause.Click += (_, __) => FinishSperrpause();
            BtnDeleteSperrpause.Click += (_, __) => DeleteSelectedOrTypedSperrpause();

            BtnCancel.Click += (_, __) => CancelDrawing();
            BtnClearSperrFilter.Click += (_, __) => ClearSperrpauseFilters();

            LstPdfs.SelectionChanged += (_, __) =>
            {
                if (LstPdfs.SelectedItem is string baseName)
                    LoadPdfByBaseName(baseName);
            };

            LstCables.SelectionChanged += (_, __) => SyncSelectedCablesFromList();
            LstSperrpauses.SelectionChanged += LstSperrpauses_SelectionChanged;

            TxtSearch.TextChanged += (_, __) => RefreshCurrentPdfCableList();
            TxtGlobalCableSearch.TextChanged += (_, __) => RefreshGlobalCableUI();
            TxtSperrSearch.TextChanged += (_, __) => RefreshCurrentPdfSperrpauseList();

            LstGlobalCables.MouseDoubleClick += (_, __) =>
            {
                if (LstGlobalCables.SelectedItem is not string entry) return;

                var parts = entry.Split("—", StringSplitOptions.TrimEntries);
                if (parts.Length != 2) return;

                var id = parts[0].Trim();
                var pdfBase = parts[1].Trim();

                LoadPdfByBaseName(pdfBase);
                TxtSearch.Text = id;

                var cable = _data.Cables.FirstOrDefault(x => x.Id.Equals(id, StringComparison.OrdinalIgnoreCase));
                if (cable != null)
                {
                    ClearSelectedCables();
                    _selectedCables.Add(cable);
                    LstCables.SelectedItem = cable.Id;
                    RedrawAll();
                    FocusOnCables(_selectedCables);
                    RefreshCurrentPdfSperrpauseList();
                }
            };

            PreviewKeyDown += (_, e) =>
            {
                if (e.Key == Key.Escape)
                {
                    CancelDrawing();
                    e.Handled = true;
                }
                else if (e.Key == Key.Enter)
                {
                    if (GetCurrentDrawKind() == DrawKind.Cable)
                    {
                        FinishCable();
                        e.Handled = true;
                    }
                    else if (GetCurrentDrawKind() == DrawKind.Sperrpause)
                    {
                        FinishSperrpause();
                        e.Handled = true;
                    }
                }
            };

            UpdateModeText();
            ApplyMultiSelectMode();
            UpdateSperrFilterInfo();
            TryRestoreLastWorkspace();
        }

        private DrawKind GetCurrentDrawKind()
        {
            if (TglDrawCable.IsChecked == true) return DrawKind.Cable;
            if (TglDrawSperrpause.IsChecked == true) return DrawKind.Sperrpause;
            return DrawKind.None;
        }

        private void LstSperrpauses_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isUpdatingSperrSelection)
                return;

            _activeSperrpauseFilters.Clear();
            foreach (var id in LstSperrpauses.SelectedItems.OfType<string>())
                _activeSperrpauseFilters.Add(id);

            RefreshCurrentPdfSperrpauseList();
            RefreshCurrentPdfCableList();
            RedrawAll();
        }

        private void AttachCableContextMenuHandlers()
        {
            if (LstCables.ItemContainerGenerator == null)
                return;

            LstCables.ItemContainerGenerator.StatusChanged += (_, __) =>
            {
                if (LstCables.ItemContainerGenerator.Status != System.Windows.Controls.Primitives.GeneratorStatus.ContainersGenerated)
                    return;

                for (int i = 0; i < LstCables.Items.Count; i++)
                {
                    if (LstCables.ItemContainerGenerator.ContainerFromIndex(i) is ListBoxItem item &&
                        item.ContextMenu != null)
                    {
                        foreach (var obj in item.ContextMenu.Items)
                        {
                            if (obj is MenuItem mi && Equals(mi.Header, "Sperrpause zuweisen"))
                            {
                                mi.Click -= AssignSperrpause_Click;
                                mi.Click += AssignSperrpause_Click;
                            }
                        }
                    }
                }
            };
        }

        private void AttachSperrpauseContextMenuHandlers()
        {
            if (LstSperrpauses.ItemContainerGenerator == null)
                return;

            LstSperrpauses.ItemContainerGenerator.StatusChanged += (_, __) =>
            {
                if (LstSperrpauses.ItemContainerGenerator.Status != System.Windows.Controls.Primitives.GeneratorStatus.ContainersGenerated)
                    return;

                for (int i = 0; i < LstSperrpauses.Items.Count; i++)
                {
                    if (LstSperrpauses.ItemContainerGenerator.ContainerFromIndex(i) is ListBoxItem item &&
                        item.ContextMenu != null)
                    {
                        foreach (var obj in item.ContextMenu.Items)
                        {
                            if (obj is MenuItem mi && Equals(mi.Header, "Kabel zuordnen"))
                            {
                                mi.Click -= AssignCablesToSperrpause_Click;
                                mi.Click += AssignCablesToSperrpause_Click;
                            }
                        }
                    }
                }
            };
        }

        private void CollapseDrawTools()
        {
            ActivateDrawMode(DrawKind.None);
            TglSetLabel.IsChecked = false;
            UpdateModeText();
        }

        private void ActivateDrawMode(DrawKind kind)
        {
            if (!ExpDrawPanel.IsExpanded && kind != DrawKind.None)
                ExpDrawPanel.IsExpanded = true;

            if (kind == DrawKind.Cable)
            {
                TglDrawCable.IsChecked = true;
                TglDrawSperrpause.IsChecked = false;
            }
            else if (kind == DrawKind.Sperrpause)
            {
                TglDrawCable.IsChecked = false;
                TglDrawSperrpause.IsChecked = true;
            }
            else
            {
                TglDrawCable.IsChecked = false;
                TglDrawSperrpause.IsChecked = false;
            }

            CancelDrawing();
            UpdateModeText();
        }

        private void ChooseWorkspaceFolder()
        {
            var dlg = new VistaFolderBrowserDialog
            {
                Description = "Workspace-Ordner auswählen",
                UseDescriptionForTitle = true
            };

            if (dlg.ShowDialog() == true)
                SetWorkspace(dlg.SelectedPath);
        }

        private void TryRestoreLastWorkspace()
        {
            try
            {
                if (!File.Exists(AppSettingsPath))
                    return;

                var json = File.ReadAllText(AppSettingsPath);
                var settings = JsonSerializer.Deserialize<AppSettings>(json);
                var path = settings?.LastWorkspacePath?.Trim() ?? "";

                if (path.Length == 0 || !IODir.Exists(path))
                    return;

                SetWorkspace(path);
            }
            catch
            {
            }
        }

        private void SaveAppSettings()
        {
            try
            {
                var dir = IOPath.GetDirectoryName(AppSettingsPath);
                if (!string.IsNullOrWhiteSpace(dir))
                    IODir.CreateDirectory(dir);

                var data = new AppSettings
                {
                    LastWorkspacePath = _workspaceRoot
                };

                var json = JsonSerializer.Serialize(data, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(AppSettingsPath, json);
            }
            catch
            {
            }
        }

        private void SetWorkspace(string root)
        {
            _workspaceRoot = root;
            _workspacePdfDir = IOPath.Combine(root, "pdf");
            _workspaceDataDir = IOPath.Combine(root, "data");
            _workspaceMetaPath = IOPath.Combine(root, "workspace.json");
            _assignmentFilePath = IOPath.Combine(root, "sperrpause_assignments.json");

            IODir.CreateDirectory(_workspacePdfDir);
            IODir.CreateDirectory(_workspaceDataDir);

            TxtWorkspacePath.Text = _workspaceRoot;
            SaveAppSettings();

            LoadWorkspaceMeta();
            LoadAssignments();

            RefreshPdfList();
            RefreshGlobalCableUI();
            RefreshCurrentPdfSperrpauseList();
        }
        private void LstCablesItem_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is ListBoxItem item)
            {
                item.IsSelected = true;
            }
        }

        private void LstSperrpausesItem_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is ListBoxItem item)
            {
                item.IsSelected = true;
            }
        }

        private void LstSperrpausesItem_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is not ListBoxItem item)
                return;

            item.IsSelected = !item.IsSelected;
            e.Handled = true;
        }
        private void OpenWorkspaceFolder()
        {
            if (string.IsNullOrWhiteSpace(_workspaceRoot) || !IODir.Exists(_workspaceRoot)) return;
            Process.Start(new ProcessStartInfo { FileName = _workspaceRoot, UseShellExecute = true });
        }

        private void LoadWorkspaceMeta()
        {
            _ws = new WorkspaceMeta();

            if (!IOFile.Exists(_workspaceMetaPath))
                return;

            var json = IOFile.ReadAllText(_workspaceMetaPath);

            try
            {
                var obj = JsonSerializer.Deserialize<WorkspaceMeta>(json);
                if (obj != null)
                {
                    _ws = obj;
                    _ws.CableIndex ??= new(StringComparer.OrdinalIgnoreCase);
                    _ws.SperrpauseIndex ??= new(StringComparer.OrdinalIgnoreCase);
                    return;
                }
            }
            catch
            {
            }

            try
            {
                var old = JsonSerializer.Deserialize<Dictionary<string, string>>(json);
                if (old != null)
                {
                    _ws = new WorkspaceMeta();
                    foreach (var kv in old)
                    {
                        var id = (kv.Key ?? "").Trim().ToUpperInvariant();
                        var pdf = (kv.Value ?? "").Trim();
                        if (id.Length == 0 || pdf.Length == 0) continue;

                        if (!_ws.CableIndex.TryGetValue(id, out var set))
                        {
                            set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                            _ws.CableIndex[id] = set;
                        }
                        set.Add(pdf);
                    }
                    SaveWorkspaceMeta();
                }
            }
            catch
            {
                _ws = new WorkspaceMeta();
            }
        }

        private void SaveWorkspaceMeta()
        {
            try
            {
                var json = JsonSerializer.Serialize(_ws, new JsonSerializerOptions { WriteIndented = true });
                IOFile.WriteAllText(_workspaceMetaPath, json);
            }
            catch
            {
            }
        }

        private void LoadAssignments()
        {
            _cableToSperrpauseAssignments = new(StringComparer.OrdinalIgnoreCase);

            try
            {
                if (string.IsNullOrWhiteSpace(_assignmentFilePath))
                    return;

                if (!IOFile.Exists(_assignmentFilePath))
                {
                    IOFile.WriteAllText(_assignmentFilePath, "{}");
                    return;
                }

                var json = IOFile.ReadAllText(_assignmentFilePath);
                var data = JsonSerializer.Deserialize<Dictionary<string, HashSet<string>>>(json);

                if (data != null)
                    _cableToSperrpauseAssignments = new Dictionary<string, HashSet<string>>(data, StringComparer.OrdinalIgnoreCase);
            }
            catch
            {
                _cableToSperrpauseAssignments = new(StringComparer.OrdinalIgnoreCase);
            }
        }

        private void SaveAssignments()
        {
            try
            {
                if (string.IsNullOrWhiteSpace(_assignmentFilePath))
                    return;

                var json = JsonSerializer.Serialize(_cableToSperrpauseAssignments, new JsonSerializerOptions
                {
                    WriteIndented = true
                });

                IOFile.WriteAllText(_assignmentFilePath, json);
            }
            catch
            {
            }
        }

        private void RefreshPdfList()
        {
            if (string.IsNullOrWhiteSpace(_workspacePdfDir) || !IODir.Exists(_workspacePdfDir))
            {
                LstPdfs.ItemsSource = new List<string>();
                return;
            }

            var pdfs = IODir.GetFiles(_workspacePdfDir, "*.pdf")
                .Select(p => IOPath.GetFileNameWithoutExtension(p))
                .OrderBy(s => s)
                .ToList();

            LstPdfs.ItemsSource = pdfs;
        }

        private void ImportPdf()
        {
            if (string.IsNullOrWhiteSpace(_workspaceRoot))
            {
                MessageBox.Show("Bitte zuerst einen Workspace wählen.");
                return;
            }

            var dlg = new OpenFileDialog
            {
                Filter = "PDF|*.pdf",
                Multiselect = false
            };

            if (dlg.ShowDialog() != true) return;

            var src = dlg.FileName;
            var baseName = IOPath.GetFileNameWithoutExtension(src);
            var dest = IOPath.Combine(_workspacePdfDir, baseName + ".pdf");

            IOFile.Copy(src, dest, overwrite: true);
            RefreshPdfList();
        }

        private void OpenSelectedPdfExternally()
        {
            if (LstPdfs.SelectedItem is not string baseName) return;

            var pdfPath = IOPath.Combine(_workspacePdfDir, baseName + ".pdf");
            if (!IOFile.Exists(pdfPath)) return;

            Process.Start(new ProcessStartInfo { FileName = pdfPath, UseShellExecute = true });
        }

        private void LoadPdfByBaseName(string baseName)
        {
            if (string.IsNullOrWhiteSpace(_workspaceRoot)) return;

            bool samePdf = baseName.Equals(_currentPdfBaseName, StringComparison.OrdinalIgnoreCase);

            _currentPdfBaseName = baseName;
            _currentPdfPath = IOPath.Combine(_workspacePdfDir, baseName + ".pdf");
            _planJsonPath = IOPath.Combine(_workspaceDataDir, baseName + ".json");

            LoadPlanData();

            if (!samePdf)
            {
                _rotationDeg = _pdfRotation.TryGetValue(_currentPdfBaseName, out var deg) ? deg : 0;
                TxtRotation.Text = _rotationDeg + "°";

                LoadPdfFirstPageToBitmap();
                ApplyRotationToImage();
                FitPage();
            }

            ClearSelectedCables();
            CancelDrawing();

            RefreshCurrentPdfCableList();
            RefreshCurrentPdfSperrpauseList();
            RefreshGlobalCableUI();
            RedrawAll();
        }

        private void LoadPlanData()
        {
            _data = new PlanData();
            try
            {
                if (IOFile.Exists(_planJsonPath))
                {
                    _data = JsonSerializer.Deserialize<PlanData>(IOFile.ReadAllText(_planJsonPath)) ?? new PlanData();
                    _data.Cables ??= new();
                    _data.Sperrpauses ??= new();
                }
            }
            catch
            {
                _data = new PlanData();
            }
        }

        private void SavePlanData()
        {
            try
            {
                IOFile.WriteAllText(
                    _planJsonPath,
                    JsonSerializer.Serialize(_data, new JsonSerializerOptions { WriteIndented = true }));
            }
            catch
            {
            }
        }

        private void LoadPdfFirstPageToBitmap()
        {
            _baseBmp = null;
            PlanImage.Source = null;
            Overlay.Children.Clear();

            if (!IOFile.Exists(_currentPdfPath)) return;

            try
            {
                using var doc = PdfiumDoc.Load(_currentPdfPath);

                var size = doc.PageSizes[0];
                int dpi = 160;
                int w = (int)Math.Round(size.Width / 72.0 * dpi);
                int h = (int)Math.Round(size.Height / 72.0 * dpi);

                using var img = doc.Render(0, w, h, dpi, dpi, PdfRenderFlags.Annotations);
                _baseBmp = ConvertGdiToBitmapSource((System.Drawing.Bitmap)img);
            }
            catch (Exception ex)
            {
                MessageBox.Show("PDF Render Fehler:\n" + ex.Message);
            }
        }

        private static BitmapSource ConvertGdiToBitmapSource(System.Drawing.Bitmap bmp)
        {
            var hBitmap = bmp.GetHbitmap();
            try
            {
                var bs = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                    hBitmap, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
                bs.Freeze();
                return bs;
            }
            finally
            {
                NativeDeleteObject(hBitmap);
            }
        }

        [System.Runtime.InteropServices.DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr hObject);

        private static void NativeDeleteObject(IntPtr hBitmap) => DeleteObject(hBitmap);

        private void Rotate(int delta)
        {
            if (_baseBmp == null) return;

            _rotationDeg = (_rotationDeg + delta) % 360;
            if (_rotationDeg < 0) _rotationDeg += 360;
            _rotationDeg = _rotationDeg switch
            {
                0 or 90 or 180 or 270 => _rotationDeg,
                _ => 0
            };

            _pdfRotation[_currentPdfBaseName] = _rotationDeg;
            TxtRotation.Text = _rotationDeg + "°";

            ApplyRotationToImage();
            RedrawAll();
        }

        private void ApplyRotationToImage()
        {
            if (_baseBmp == null) return;

            BitmapSource src = _baseBmp;
            if (_rotationDeg != 0)
            {
                var tb = new TransformedBitmap(_baseBmp, new RotateTransform(_rotationDeg));
                tb.Freeze();
                src = tb;
            }

            PlanImage.Source = src;
            Overlay.Width = src.PixelWidth;
            Overlay.Height = src.PixelHeight;
        }

        private Point ViewToBase(Point pView)
        {
            if (_baseBmp == null) return pView;

            double W = _baseBmp.PixelWidth;
            double H = _baseBmp.PixelHeight;

            return _rotationDeg switch
            {
                0 => pView,
                90 => new Point(pView.Y, H - pView.X),
                180 => new Point(W - pView.X, H - pView.Y),
                270 => new Point(W - pView.Y, pView.X),
                _ => pView
            };
        }

        private Point BaseToView(Point pBase)
        {
            if (_baseBmp == null) return pBase;

            double W = _baseBmp.PixelWidth;
            double H = _baseBmp.PixelHeight;

            return _rotationDeg switch
            {
                0 => pBase,
                90 => new Point(H - pBase.Y, pBase.X),
                180 => new Point(W - pBase.X, H - pBase.Y),
                270 => new Point(pBase.Y, W - pBase.X),
                _ => pBase
            };
        }

        private void PdfScroll_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (PlanImage.Source == null) return;

            var mouseOnScroll = e.GetPosition(PdfScroll);
            var contentPtBefore = e.GetPosition(Overlay);

            double factor = e.Delta > 0 ? 1.15 : 1 / 1.15;
            double newZoom = Math.Max(0.2, Math.Min(8.0, _zoom * factor));

            _zoom = newZoom;
            ZoomTf.ScaleX = ZoomTf.ScaleY = _zoom;

            PdfScroll.UpdateLayout();

            PdfScroll.ScrollToHorizontalOffset(contentPtBefore.X * _zoom - mouseOnScroll.X);
            PdfScroll.ScrollToVerticalOffset(contentPtBefore.Y * _zoom - mouseOnScroll.Y);

            e.Handled = true;
        }

        private void FitWidth()
        {
            if (PlanImage.Source is not BitmapSource bmp) return;
            if (PdfScroll.ViewportWidth <= 1) return;

            _zoom = PdfScroll.ViewportWidth / bmp.PixelWidth;
            _zoom = Math.Max(0.05, Math.Min(8.0, _zoom));
            ZoomTf.ScaleX = ZoomTf.ScaleY = _zoom;

            PdfScroll.ScrollToHorizontalOffset(0);
            PdfScroll.ScrollToVerticalOffset(0);
        }

        private void FitPage()
        {
            if (PlanImage.Source is not BitmapSource bmp) return;
            if (PdfScroll.ViewportWidth <= 1 || PdfScroll.ViewportHeight <= 1) return;

            var zx = PdfScroll.ViewportWidth / bmp.PixelWidth;
            var zy = PdfScroll.ViewportHeight / bmp.PixelHeight;
            _zoom = Math.Min(zx, zy);

            _zoom = Math.Max(0.05, Math.Min(8.0, _zoom));
            ZoomTf.ScaleX = ZoomTf.ScaleY = _zoom;

            PdfScroll.ScrollToHorizontalOffset(0);
            PdfScroll.ScrollToVerticalOffset(0);
        }

        private RenderTargetBitmap? CaptureCurrentViewportBitmap(double exportScale = 1.0, double targetDpi = 96.0)
        {
            if (PlanImage.Source == null || PdfScroll.ActualWidth < 2 || PdfScroll.ActualHeight < 2)
                return null;

            PdfScroll.UpdateLayout();

            exportScale = Math.Max(1.0, exportScale);
            targetDpi = Math.Max(96.0, targetDpi);

            int width = (int)Math.Max(1, Math.Round(PdfScroll.ActualWidth * exportScale));
            int height = (int)Math.Max(1, Math.Round(PdfScroll.ActualHeight * exportScale));
            double dpi = targetDpi;

            var rtb = new RenderTargetBitmap(width, height, dpi, dpi, PixelFormats.Pbgra32);

            var visual = new DrawingVisual();
            using (var dc = visual.RenderOpen())
            {
                dc.PushTransform(new ScaleTransform(exportScale, exportScale));
                var vb = new VisualBrush(PdfScroll)
                {
                    Stretch = Stretch.None,
                    AlignmentX = AlignmentX.Left,
                    AlignmentY = AlignmentY.Top
                };
                dc.DrawRectangle(vb, null, new Rect(new Point(0, 0), new Size(PdfScroll.ActualWidth, PdfScroll.ActualHeight)));
            }

            rtb.Render(visual);
            return rtb;
        }

        private void ExportCurrentViewAsPng()
        {
            var bmp = CaptureCurrentViewportBitmap(exportScale: 3.0, targetDpi: 300.0);
            if (bmp == null)
            {
                MessageBox.Show("Aktuell gibt es keine sichtbare Planansicht zum Exportieren.");
                return;
            }

            var dlg = new SaveFileDialog
            {
                Filter = "PNG Bild|*.png",
                FileName = $"{_currentPdfBaseName}_view.png"
            };

            if (dlg.ShowDialog() != true)
                return;

            var encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(bmp));

            using var fs = File.Create(dlg.FileName);
            encoder.Save(fs);
        }

        private void ExportCurrentViewAsPdf()
        {
            var bmp = CaptureCurrentViewportBitmap(exportScale: 3.0, targetDpi: 300.0);
            if (bmp == null)
            {
                MessageBox.Show("Aktuell gibt es keine sichtbare Planansicht zum Exportieren.");
                return;
            }

            var dlg = new SaveFileDialog
            {
                Filter = "PDF Datei|*.pdf",
                FileName = $"{_currentPdfBaseName}_view.pdf"
            };

            if (dlg.ShowDialog() != true)
                return;

            byte[] pngBytes;
            var encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(bmp));
            using (var ms = new MemoryStream())
            {
                encoder.Save(ms);
                pngBytes = ms.ToArray();
            }

            string tempPngPath = IOPath.Combine(IOPath.GetTempPath(), $"cableplan_{Guid.NewGuid():N}.png");

            try
            {
                IOFile.WriteAllBytes(tempPngPath, pngBytes);

                using var doc = new SharpPdfDocument();
                var page = doc.AddPage();
                page.Width = bmp.PixelWidth * 72.0 / 96.0;
                page.Height = bmp.PixelHeight * 72.0 / 96.0;

                using (var gfx = XGraphics.FromPdfPage(page))
                using (var img = XImage.FromFile(tempPngPath))
                {
                    gfx.DrawImage(img, 0, 0, page.Width, page.Height);
                }

                doc.Save(dlg.FileName);
            }
            finally
            {
                if (IOFile.Exists(tempPngPath))
                    IOFile.Delete(tempPngPath);
            }
        }

        private void PrintCurrentView()
        {
            var bmp = CaptureCurrentViewportBitmap();
            if (bmp == null)
            {
                MessageBox.Show("Aktuell gibt es keine sichtbare Planansicht zum Drucken.");
                return;
            }

            var dlg = new PrintDialog();
            if (dlg.ShowDialog() != true)
                return;

            var image = new Image
            {
                Source = bmp,
                Stretch = Stretch.Uniform,
                Width = dlg.PrintableAreaWidth,
                Height = dlg.PrintableAreaHeight
            };

            image.Measure(new Size(dlg.PrintableAreaWidth, dlg.PrintableAreaHeight));
            image.Arrange(new Rect(new Point(0, 0), image.DesiredSize));
            dlg.PrintVisual(image, "CablePlan Aktuelle Ansicht");
        }

        private void UpdateModeText()
        {
            if (!ExpDrawPanel.IsExpanded)
            {
                TxtMode.Text = "Modus: Ansicht (Zeichnen eingeklappt)";
                return;
            }

            string mode = GetCurrentDrawKind() switch
            {
                DrawKind.Cable => "Modus: Kabel zeichnen",
                DrawKind.Sperrpause => "Modus: Sperrpause zeichnen",
                _ => TglSetLabel.IsChecked == true ? "Modus: Beschriftung" : "Modus: Ansicht"
            };

            TxtMode.Text = mode;
        }

        private void ApplyMultiSelectMode()
        {
            bool multi = TglMultiSelect.IsChecked == true;
            LstCables.SelectionMode = multi ? SelectionMode.Multiple : SelectionMode.Single;

            if (!multi && _selectedCables.Count > 1)
            {
                var keep = _selectedCables.FirstOrDefault();
                _selectedCables.Clear();
                if (keep != null)
                    _selectedCables.Add(keep);
            }

            RefreshCurrentPdfCableList();
            RedrawAll();
            FocusOnCables(_selectedCables);
            RefreshCurrentPdfSperrpauseList();
        }

        private void SyncSelectedCablesFromList()
        {
            if (_isUpdatingCableSelection)
                return;

            _selectedCables.Clear();

            if (LstCables.SelectionMode == SelectionMode.Single)
            {
                if (LstCables.SelectedItem is string id)
                {
                    var cable = _data.Cables.FirstOrDefault(x => x.Id.Equals(id, StringComparison.OrdinalIgnoreCase));
                    if (cable != null)
                        _selectedCables.Add(cable);
                }
            }
            else
            {
                foreach (var item in LstCables.SelectedItems)
                {
                    if (item is string id)
                    {
                        var cable = _data.Cables.FirstOrDefault(x => x.Id.Equals(id, StringComparison.OrdinalIgnoreCase));
                        if (cable != null)
                            _selectedCables.Add(cable);
                    }
                }
            }

            RedrawAll();
            FocusOnCables(_selectedCables);
            RefreshCurrentPdfSperrpauseList();
        }

        private void ClearSelectedCables()
        {
            _selectedCables.Clear();

            _isUpdatingCableSelection = true;
            try
            {
                if (LstCables.SelectionMode != SelectionMode.Single)
                    LstCables.SelectedItems.Clear();

                LstCables.SelectedItem = null;
            }
            finally
            {
                _isUpdatingCableSelection = false;
            }
        }

        private void ClearCableSelectionAndRefresh()
        {
            ClearSelectedCables();
            RedrawAll();
            RefreshCurrentPdfCableList();
            RefreshCurrentPdfSperrpauseList();
        }

        private void ClearSperrpauseFilters()
        {
            _activeSperrpauseFilters.Clear();

            _isUpdatingSperrSelection = true;
            try
            {
                LstSperrpauses.SelectedItems.Clear();
            }
            finally
            {
                _isUpdatingSperrSelection = false;
            }

            RedrawAll();
            UpdateSperrFilterInfo();
            RefreshCurrentPdfCableList();
        }

        private void PdfScroll_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Middle)
            {
                _midPanning = true;
                _midStart = e.GetPosition(PdfScroll);
                _midHStart = PdfScroll.HorizontalOffset;
                _midVStart = PdfScroll.VerticalOffset;
                PdfScroll.CaptureMouse();
                e.Handled = true;
            }
        }

        private void PdfScroll_PreviewMouseUp(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Middle && _midPanning)
            {
                _midPanning = false;
                PdfScroll.ReleaseMouseCapture();
                e.Handled = true;
                return;
            }

            if (e.ChangedButton == MouseButton.Left && _pendingLeftPan)
            {
                PdfScroll.ReleaseMouseCapture();

                if (!_leftPanActive)
                {
                    var pView = e.GetPosition(Overlay);
                    var hit = FindNearestCableByLabel(pView, 14 / _zoom);
                    if (hit != null)
                    {
                        if (LstCables.SelectionMode == SelectionMode.Single)
                        {
                            ClearSelectedCables();
                            _selectedCables.Add(hit);
                            LstCables.SelectedItem = hit.Id;
                        }
                        else
                        {
                            bool exists = _selectedCables.Any(x => x.Id.Equals(hit.Id, StringComparison.OrdinalIgnoreCase));
                            if (exists)
                            {
                                _selectedCables.RemoveAll(x => x.Id.Equals(hit.Id, StringComparison.OrdinalIgnoreCase));
                                LstCables.SelectedItems.Remove(hit.Id);
                            }
                            else
                            {
                                _selectedCables.Add(hit);
                                LstCables.SelectedItems.Add(hit.Id);
                            }
                        }

                        RedrawAll();
                        FocusOnCables(_selectedCables);
                        RefreshCurrentPdfSperrpauseList();
                    }
                }

                _pendingLeftPan = false;
                _leftPanActive = false;
                e.Handled = true;
            }
        }

        private void PdfScroll_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (_midPanning)
            {
                var now = e.GetPosition(PdfScroll);
                var delta = now - _midStart;
                PdfScroll.ScrollToHorizontalOffset(_midHStart - delta.X);
                PdfScroll.ScrollToVerticalOffset(_midVStart - delta.Y);
                e.Handled = true;
                return;
            }

            if (_pendingLeftPan)
            {
                var now = e.GetPosition(PdfScroll);
                var delta = now - _leftStart;

                if (!_leftPanActive && (Math.Abs(delta.X) > 3 || Math.Abs(delta.Y) > 3))
                    _leftPanActive = true;

                if (_leftPanActive)
                {
                    PdfScroll.ScrollToHorizontalOffset(_leftHStart - delta.X);
                    PdfScroll.ScrollToVerticalOffset(_leftVStart - delta.Y);
                }

                e.Handled = true;
            }
        }

        private void PdfScroll_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (PlanImage.Source == null) return;

            var pView = e.GetPosition(Overlay);

            if (TglSetLabel.IsChecked == true)
            {
                if (_selectedCables.Count == 0) return;

                var target = _selectedCables.First();
                var b = ViewToBase(pView);
                target.LabelPoint = new PointD(b.X, b.Y);

                SavePlanData();
                RefreshGlobalIndexesForCurrentPdf();
                RedrawAll();

                e.Handled = true;
                return;
            }

            if (GetCurrentDrawKind() != DrawKind.None)
            {
                var pBase = ViewToBase(pView);
                _draftPointsBase.Add(pBase);
                RedrawAll();
                e.Handled = true;
                return;
            }

            _pendingLeftPan = true;
            _leftPanActive = false;
            _leftStart = e.GetPosition(PdfScroll);
            _leftHStart = PdfScroll.HorizontalOffset;
            _leftVStart = PdfScroll.VerticalOffset;
            PdfScroll.CaptureMouse();
            e.Handled = true;
        }

        private void CancelDrawing()
        {
            _draftPointsBase.Clear();
            RedrawAll();
        }

        private void RefreshGlobalIndexesForCurrentPdf()
        {
            if (string.IsNullOrWhiteSpace(_currentPdfBaseName)) return;

            foreach (var key in _ws.CableIndex.Keys.ToList())
            {
                _ws.CableIndex[key].Remove(_currentPdfBaseName);
                if (_ws.CableIndex[key].Count == 0)
                    _ws.CableIndex.Remove(key);
            }

            foreach (var key in _ws.SperrpauseIndex.Keys.ToList())
            {
                _ws.SperrpauseIndex[key].Remove(_currentPdfBaseName);
                if (_ws.SperrpauseIndex[key].Count == 0)
                    _ws.SperrpauseIndex.Remove(key);
            }

            foreach (var c in _data.Cables)
            {
                var idKey = (c.Id ?? "").Trim().ToUpperInvariant();
                if (idKey.Length == 0) continue;
                AddCableToGlobalIndex(idKey, _currentPdfBaseName);
            }

            foreach (var s in _data.Sperrpauses)
            {
                var idKey = (s.Id ?? "").Trim().ToUpperInvariant();
                if (idKey.Length == 0) continue;
                AddSperrpauseToGlobalIndex(idKey, _currentPdfBaseName);
            }

            SaveWorkspaceMeta();
            RefreshGlobalCableUI();
        }

        private void FinishCable()
        {
            if (_draftPointsBase.Count < 2) return;

            var idRaw = (TxtCableId.Text ?? "").Trim();
            if (string.IsNullOrWhiteSpace(idRaw))
            {
                MessageBox.Show("Bitte Kabel-ID eingeben.");
                return;
            }

            var idKey = idRaw.Trim().ToUpperInvariant();

            var basePts = _draftPointsBase.ToList();
            var baseLabel = basePts[^1];

            var existing = _data.Cables.FirstOrDefault(x => x.Id.Equals(idRaw, StringComparison.OrdinalIgnoreCase));
            if (existing != null) _data.Cables.Remove(existing);

            var cable = new Cable
            {
                Id = idRaw,
                Points = basePts.Select(p => new PointD(p.X, p.Y)).ToList(),
                LabelPoint = new PointD(baseLabel.X, baseLabel.Y)
            };

            _data.Cables.Add(cable);
            SavePlanData();

            AddCableToGlobalIndex(idKey, _currentPdfBaseName);
            SaveWorkspaceMeta();
            RefreshGlobalCableUI();

            _draftPointsBase.Clear();
            ClearSelectedCables();

            RefreshCurrentPdfCableList();
            RedrawAll();
            RefreshCurrentPdfSperrpauseList();
        }

        private void FinishSperrpause()
        {
            if (_draftPointsBase.Count < 2) return;

            var idRaw = (TxtSperrpauseId.Text ?? "").Trim();
            if (string.IsNullOrWhiteSpace(idRaw))
            {
                MessageBox.Show("Bitte Sperrpause-ID eingeben.");
                return;
            }

            var idKey = idRaw.Trim().ToUpperInvariant();

            var basePts = _draftPointsBase.ToList();
            var baseLabel = basePts[^1];

            var sperr = new Sperrpause
            {
                Id = idRaw,
                Points = basePts.Select(p => new PointD(p.X, p.Y)).ToList(),
                LabelPoint = new PointD(baseLabel.X, baseLabel.Y)
            };

            _data.Sperrpauses.Add(sperr);
            SavePlanData();

            AddSperrpauseToGlobalIndex(idKey, _currentPdfBaseName);
            SaveWorkspaceMeta();

            _activeSperrpauseFilters.Add(idRaw);
            _draftPointsBase.Clear();

            RefreshCurrentPdfSperrpauseList();
            RedrawAll();
        }

        private void DeleteSelectedOrTypedCable()
        {
            if (string.IsNullOrWhiteSpace(_planJsonPath)) return;

            var idsToDelete = new List<string>();

            var typed = (TxtCableId.Text ?? "").Trim();
            if (!string.IsNullOrWhiteSpace(typed))
            {
                idsToDelete.Add(typed);
            }
            else if (_selectedCables.Count > 0)
            {
                idsToDelete.AddRange(_selectedCables.Select(x => x.Id));
            }

            if (idsToDelete.Count == 0) return;

            foreach (var id in idsToDelete.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                var c = _data.Cables.FirstOrDefault(x => x.Id.Equals(id, StringComparison.OrdinalIgnoreCase));
                if (c == null) continue;

                _data.Cables.Remove(c);
                RemoveCableFromGlobalIndex(id.Trim().ToUpperInvariant(), _currentPdfBaseName);
            }

            SavePlanData();
            SaveWorkspaceMeta();

            ClearSelectedCables();
            RefreshCurrentPdfCableList();
            RefreshGlobalCableUI();
            RedrawAll();
            RefreshCurrentPdfSperrpauseList();
        }

        private void DeleteSelectedOrTypedSperrpause()
        {
            if (string.IsNullOrWhiteSpace(_planJsonPath)) return;

            var idsToDelete = new List<string>();

            var typed = (TxtSperrpauseId.Text ?? "").Trim();
            if (!string.IsNullOrWhiteSpace(typed))
            {
                idsToDelete.Add(typed);
            }
            else
            {
                idsToDelete.AddRange(LstSperrpauses.SelectedItems.OfType<string>());
            }

            if (idsToDelete.Count == 0) return;

            foreach (var id in idsToDelete.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                var toRemove = _data.Sperrpauses.Where(x => x.Id.Equals(id, StringComparison.OrdinalIgnoreCase)).ToList();
                foreach (var s in toRemove)
                    _data.Sperrpauses.Remove(s);

                RemoveSperrpauseFromGlobalIndex(id.Trim().ToUpperInvariant(), _currentPdfBaseName);
                _activeSperrpauseFilters.Remove(id);
            }

            SavePlanData();
            SaveWorkspaceMeta();

            RefreshCurrentPdfSperrpauseList();
            RedrawAll();
        }

        private void RedrawAll()
        {
            Overlay.Children.Clear();

            if (_draftPointsBase.Count > 0)
            {
                var draftBrush = GetCurrentDrawKind() == DrawKind.Sperrpause ? SperrpauseHighlightBrush : Brushes.Orange;

                var draft = new Polyline
                {
                    Stroke = draftBrush,
                    StrokeThickness = 3,
                    StrokeLineJoin = PenLineJoin.Round
                };

                foreach (var pb in _draftPointsBase)
                    draft.Points.Add(BaseToView(pb));

                Overlay.Children.Add(draft);
            }

            foreach (var sperr in _data.Sperrpauses.Where(x => _activeSperrpauseFilters.Contains(x.Id)))
            {
                var under = new Polyline
                {
                    Stroke = Brushes.Black,
                    StrokeThickness = 9,
                    StrokeLineJoin = PenLineJoin.Round,
                    Opacity = 0.35
                };

                var hi = new Polyline
                {
                    Stroke = SperrpauseHighlightBrush,
                    StrokeThickness = 6,
                    StrokeLineJoin = PenLineJoin.Round
                };

                foreach (var pt in sperr.Points)
                {
                    var v = BaseToView(new Point(pt.X, pt.Y));
                    under.Points.Add(v);
                    hi.Points.Add(v);
                }

                Overlay.Children.Add(under);
                Overlay.Children.Add(hi);

                var lpV = BaseToView(new Point(sperr.LabelPoint.X, sperr.LabelPoint.Y));

                var marker = new Ellipse
                {
                    Width = 12,
                    Height = 12,
                    Fill = SperrpauseHighlightBrush,
                    Stroke = Brushes.Black,
                    StrokeThickness = 1.2
                };
                Canvas.SetLeft(marker, lpV.X - marker.Width / 2);
                Canvas.SetTop(marker, lpV.Y - marker.Height / 2);
                Overlay.Children.Add(marker);

                var tb = new TextBlock
                {
                    Text = sperr.Id,
                    Foreground = Brushes.Black,
                    Background = SperrpauseHighlightBrush,
                    FontSize = 14,
                    FontWeight = FontWeights.Bold
                };
                Canvas.SetLeft(tb, lpV.X + 10);
                Canvas.SetTop(tb, lpV.Y - 10);
                Overlay.Children.Add(tb);
            }

            foreach (var cable in _selectedCables)
            {
                var under = new Polyline
                {
                    Stroke = Brushes.Black,
                    StrokeThickness = 10,
                    StrokeLineJoin = PenLineJoin.Round,
                    Opacity = 0.35
                };

                var hi = new Polyline
                {
                    Stroke = CableHighlightBrush,
                    StrokeThickness = 7,
                    StrokeLineJoin = PenLineJoin.Round
                };

                foreach (var pt in cable.Points)
                {
                    var v = BaseToView(new Point(pt.X, pt.Y));
                    under.Points.Add(v);
                    hi.Points.Add(v);
                }

                Overlay.Children.Add(under);
                Overlay.Children.Add(hi);

                var lpV = BaseToView(new Point(cable.LabelPoint.X, cable.LabelPoint.Y));

                var marker = new Ellipse
                {
                    Width = 14,
                    Height = 14,
                    Fill = CableHighlightBrush,
                    Stroke = Brushes.White,
                    StrokeThickness = 1.5
                };
                Canvas.SetLeft(marker, lpV.X - marker.Width / 2);
                Canvas.SetTop(marker, lpV.Y - marker.Height / 2);
                Overlay.Children.Add(marker);

                var tb = new TextBlock
                {
                    Text = cable.Id,
                    Foreground = CableHighlightBrush,
                    Background = Brushes.White,
                    FontSize = 16,
                    FontWeight = FontWeights.Bold
                };
                Canvas.SetLeft(tb, lpV.X + 12);
                Canvas.SetTop(tb, lpV.Y - 12);
                Overlay.Children.Add(tb);
            }
        }

        private Cable? FindNearestCableByLabel(Point pView, double maxDist)
        {
            Cable? best = null;
            double bestD = double.MaxValue;

            foreach (var c in _data.Cables)
            {
                var lpV = BaseToView(new Point(c.LabelPoint.X, c.LabelPoint.Y));
                var dx = lpV.X - pView.X;
                var dy = lpV.Y - pView.Y;
                var d = Math.Sqrt(dx * dx + dy * dy);

                if (d < bestD)
                {
                    bestD = d;
                    best = c;
                }
            }

            return best != null && bestD <= maxDist ? best : null;
        }

        private void FocusOnCables(IEnumerable<Cable> cables)
        {
            var list = cables?.ToList() ?? new List<Cable>();
            if (list.Count == 0 || PlanImage.Source is not BitmapSource) return;

            var allViewPoints = new List<Point>();

            foreach (var cable in list)
            {
                foreach (var pt in cable.Points)
                    allViewPoints.Add(BaseToView(new Point(pt.X, pt.Y)));

                allViewPoints.Add(BaseToView(new Point(cable.LabelPoint.X, cable.LabelPoint.Y)));
            }

            if (allViewPoints.Count == 0) return;

            double minX = allViewPoints.Min(p => p.X);
            double minY = allViewPoints.Min(p => p.Y);
            double maxX = allViewPoints.Max(p => p.X);
            double maxY = allViewPoints.Max(p => p.Y);

            double width = Math.Max(1, maxX - minX);
            double height = Math.Max(1, maxY - minY);

            const double margin = 320;

            if (PdfScroll.ViewportWidth > 20 && PdfScroll.ViewportHeight > 20)
            {
                double fitZoomX = PdfScroll.ViewportWidth / (width + margin);
                double fitZoomY = PdfScroll.ViewportHeight / (height + margin);
                double newZoom = Math.Min(fitZoomX, fitZoomY);
                newZoom = Math.Max(0.2, Math.Min(2.0, newZoom));

                _zoom = newZoom;
                ZoomTf.ScaleX = ZoomTf.ScaleY = _zoom;
                PdfScroll.UpdateLayout();
            }

            double centerX = (minX + maxX) / 2.0;
            double centerY = (minY + maxY) / 2.0;

            PdfScroll.ScrollToHorizontalOffset(Math.Max(0, centerX * _zoom - PdfScroll.ViewportWidth / 2));
            PdfScroll.ScrollToVerticalOffset(Math.Max(0, centerY * _zoom - PdfScroll.ViewportHeight / 2));
        }

        private void RefreshCurrentPdfCableList()
        {
            var q = (TxtSearch.Text ?? "").Trim();

            IEnumerable<string> visibleList = _data.Cables
                .Select(c => c.Id)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(s => s);

            bool sperrBasedCableFilterActive = TglSperrBasedCableFilter.IsChecked == true;
            if (sperrBasedCableFilterActive && _activeSperrpauseFilters.Count > 0)
            {
                var allowedCableIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var kv in _cableToSperrpauseAssignments)
                {
                    if (kv.Value.Any(sp => _activeSperrpauseFilters.Contains(sp)))
                        allowedCableIds.Add(kv.Key);
                }

                visibleList = visibleList
                    .Where(id => allowedCableIds.Contains(id))
                    .OrderBy(s => s);
            }

            if (!string.IsNullOrWhiteSpace(q))
                visibleList = visibleList
                    .Where(id => id.Contains(q, StringComparison.OrdinalIgnoreCase))
                    .OrderBy(s => s);

            var finalList = visibleList.ToList();
            var selectedIds = _selectedCables
                .Select(x => x.Id)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            _isUpdatingCableSelection = true;
            try
            {
                LstCables.ItemsSource = null;
                LstCables.ItemsSource = finalList;

                if (LstCables.SelectionMode == SelectionMode.Single)
                {
                    var firstSelected = finalList.FirstOrDefault(id => selectedIds.Contains(id));
                    LstCables.SelectedItem = firstSelected;
                }
                else
                {
                    LstCables.SelectedItems.Clear();

                    foreach (var id in finalList)
                    {
                        if (selectedIds.Contains(id))
                            LstCables.SelectedItems.Add(id);
                    }
                }
            }
            finally
            {
                _isUpdatingCableSelection = false;
            }
        }

        private void RefreshCurrentPdfSperrpauseList()
        {
            var q = (TxtSperrSearch.Text ?? "").Trim();

            IEnumerable<string> visibleList = GetVisibleSperrpauseIdsForCurrentPlan()
                .OrderBy(s => s);

            if (!string.IsNullOrWhiteSpace(q))
                visibleList = visibleList
                    .Where(id => id.Contains(q, StringComparison.OrdinalIgnoreCase))
                    .OrderBy(s => s);

            var finalList = visibleList.ToList();

            _isUpdatingSperrSelection = true;
            try
            {
                LstSperrpauses.ItemsSource = null;
                LstSperrpauses.ItemsSource = finalList;
                LstSperrpauses.SelectedItems.Clear();

                foreach (var id in finalList)
                {
                    if (_activeSperrpauseFilters.Contains(id))
                        LstSperrpauses.SelectedItems.Add(id);
                }
            }
            finally
            {
                _isUpdatingSperrSelection = false;
            }

            UpdateSperrFilterInfo();
        }

        private List<string> GetVisibleSperrpauseIdsForCurrentPlan()
        {
            var allCurrentIds = _data.Sperrpauses
                .Select(s => s.Id)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            bool cableFilterActive = TglCableBasedSperrFilter.IsChecked == true;
            if (!cableFilterActive)
                return allCurrentIds;

            if (_selectedCables.Count == 0)
                return allCurrentIds;

            var allowed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var cable in _selectedCables)
            {
                if (_cableToSperrpauseAssignments.TryGetValue(cable.Id, out var assigned))
                {
                    foreach (var sp in assigned)
                        allowed.Add(sp);
                }
            }

            if (allowed.Count == 0)
                return new List<string>();

            return allCurrentIds
                .Where(id => allowed.Contains(id))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private void UpdateSperrFilterInfo()
        {
            bool cableFilterActive = TglCableBasedSperrFilter.IsChecked == true;
            bool sperrBasedCableFilterActive = TglSperrBasedCableFilter.IsChecked == true;
            var visibleCount = GetVisibleSperrpauseIdsForCurrentPlan().Count;

            if (!cableFilterActive)
            {
                TxtSperrFilterInfo.Text = sperrBasedCableFilterActive
                    ? "Sperr→Kabel-Filter aktiv, Kabel→Sperr-Filter aus"
                    : "Alle Sperrpausen sichtbar (Kabel-Filter aus)";
                return;
            }

            if (_selectedCables.Count == 0)
            {
                TxtSperrFilterInfo.Text = "Alle Sperrpausen sichtbar";
                return;
            }

            TxtSperrFilterInfo.Text = $"{visibleCount} passende Sperrpausen";
        }

        private void AddCableToGlobalIndex(string idUpper, string pdfBase)
        {
            if (!_ws.CableIndex.TryGetValue(idUpper, out var set))
            {
                set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                _ws.CableIndex[idUpper] = set;
            }
            set.Add(pdfBase);
        }

        private void RemoveCableFromGlobalIndex(string idUpper, string pdfBase)
        {
            if (!_ws.CableIndex.TryGetValue(idUpper, out var set)) return;

            set.Remove(pdfBase);
            if (set.Count == 0)
                _ws.CableIndex.Remove(idUpper);
        }

        private void AddSperrpauseToGlobalIndex(string idUpper, string pdfBase)
        {
            if (!_ws.SperrpauseIndex.TryGetValue(idUpper, out var set))
            {
                set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                _ws.SperrpauseIndex[idUpper] = set;
            }
            set.Add(pdfBase);
        }

        private void RemoveSperrpauseFromGlobalIndex(string idUpper, string pdfBase)
        {
            if (!_ws.SperrpauseIndex.TryGetValue(idUpper, out var set)) return;

            set.Remove(pdfBase);
            if (set.Count == 0)
                _ws.SperrpauseIndex.Remove(idUpper);
        }

        private void RefreshGlobalCableUI()
        {
            if (string.IsNullOrWhiteSpace(_workspaceRoot))
            {
                LstGlobalCables.ItemsSource = new List<string>();
                return;
            }

            var q = (TxtGlobalCableSearch.Text ?? "").Trim();
            var rows = new List<string>();

            foreach (var kv in _ws.CableIndex.OrderBy(k => k.Key))
            {
                var id = kv.Key;
                foreach (var pdf in kv.Value.OrderBy(x => x))
                {
                    if (!string.IsNullOrWhiteSpace(q) &&
                        !id.Contains(q, StringComparison.OrdinalIgnoreCase) &&
                        !pdf.Contains(q, StringComparison.OrdinalIgnoreCase))
                        continue;

                    rows.Add($"{id}  —  {pdf}");
                }
            }

            LstGlobalCables.ItemsSource = rows;
        }

        private HashSet<string> GetAssignedSperrpausesForCable(string cableId)
        {
            if (_cableToSperrpauseAssignments.TryGetValue(cableId, out var set))
                return new HashSet<string>(set, StringComparer.OrdinalIgnoreCase);

            return new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        }

        private List<string> GetCurrentPlanDistinctSperrpauses()
        {
            return _data.Sperrpauses
                .Select(s => s.Id)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(s => s)
                .ToList();
        }

        private void AssignSperrpause_Click(object sender, RoutedEventArgs e)
        {
            var targetCableIds = GetTargetCableIdsForAssignment();

            if (targetCableIds.Count == 0)
            {
                MessageBox.Show("Bitte zuerst ein Kabel auswählen.");
                return;
            }

            var currentPlanSperrpauses = GetCurrentPlanDistinctSperrpauses();

            // Multi selection varsa ortak atanmışları baz alıyoruz:
            // sadece tüm seçili kablolarda ortak olanlar işaretli gelsin
            HashSet<string> alreadyAssigned;
            if (targetCableIds.Count == 1)
            {
                alreadyAssigned = GetAssignedSperrpausesForCable(targetCableIds[0]);
            }
            else
            {
                alreadyAssigned = GetAssignedSperrpausesForCable(targetCableIds[0]);

                foreach (var cableId in targetCableIds.Skip(1))
                {
                    alreadyAssigned.IntersectWith(GetAssignedSperrpausesForCable(cableId));
                }
            }

            var wnd = new SperrpauseAssignmentWindow(
                targetCableIds,
                currentPlanSperrpauses,
                alreadyAssigned)
            {
                Owner = this
            };

            if (wnd.ShowDialog() == true)
            {
                foreach (var cableId in targetCableIds)
                {
                    _cableToSperrpauseAssignments[cableId] =
                        new HashSet<string>(wnd.SelectedSperrpauses, StringComparer.OrdinalIgnoreCase);
                }

                SaveAssignments();
                RefreshCurrentPdfSperrpauseList();
                RefreshCurrentPdfCableList();
            }
        }

        private void AssignCablesToSperrpause_Click(object sender, RoutedEventArgs e)
        {
            if (LstSperrpauses.SelectedItem is not string sperrpauseId || string.IsNullOrWhiteSpace(sperrpauseId))
            {
                MessageBox.Show("Bitte zuerst eine Sperrpause auswählen.");
                return;
            }

            var currentPlanCableIds = _data.Cables
                .Select(x => x.Id)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x)
                .ToList();

            var assignedCableIds = _cableToSperrpauseAssignments
                .Where(kv => kv.Value.Contains(sperrpauseId))
                .Select(kv => kv.Key)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            var wnd = new CableAssignmentWindow(sperrpauseId, currentPlanCableIds, assignedCableIds)
            {
                Owner = this
            };

            if (wnd.ShowDialog() != true)
                return;

            var selectedIds = wnd.SelectedCableIds
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            foreach (var cableId in currentPlanCableIds)
            {
                if (!_cableToSperrpauseAssignments.TryGetValue(cableId, out var currentSet))
                {
                    currentSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    _cableToSperrpauseAssignments[cableId] = currentSet;
                }

                if (selectedIds.Contains(cableId))
                    currentSet.Add(sperrpauseId);
                else
                    currentSet.Remove(sperrpauseId);
            }

            foreach (var cableId in _cableToSperrpauseAssignments.Keys.ToList())
            {
                if (_cableToSperrpauseAssignments[cableId].Count == 0)
                    _cableToSperrpauseAssignments.Remove(cableId);
            }

            SaveAssignments();
            RefreshCurrentPdfSperrpauseList();
            RefreshCurrentPdfCableList();
            RedrawAll();
        }
        private List<string> GetTargetCableIdsForAssignment()
        {
            var result = new List<string>();

            // 1) Multi selection varsa hepsini al
            if (_selectedCables.Count > 0)
            {
                result.AddRange(_selectedCables
                    .Select(x => x.Id)
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Distinct(StringComparer.OrdinalIgnoreCase));
            }

            // 2) Eğer listeden gelmediyse, sağ tıklanan item seçili olabilir
            if (result.Count == 0 && LstCables.SelectedItem is string selectedId && !string.IsNullOrWhiteSpace(selectedId))
            {
                result.Add(selectedId);
            }

            return result
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x)
                .ToList();
        }
    }
}
