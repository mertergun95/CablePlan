using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace CablePlan
{
    public partial class SperrpauseAssignmentWindow : Window
    {
        private readonly List<string> _allSperrpauses;
        private readonly HashSet<string> _selected;
        private readonly Dictionary<string, CheckBox> _checkBoxes =
            new Dictionary<string, CheckBox>(StringComparer.OrdinalIgnoreCase);

        public List<string> SelectedSperrpauses { get; private set; } = new();

        public SperrpauseAssignmentWindow(
            IEnumerable<string> cableIds,
            IEnumerable<string> availableSperrpauses,
            IEnumerable<string> alreadyAssigned)
        {
            InitializeComponent();

            _allSperrpauses = availableSperrpauses
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x)
                .ToList();

            _selected = new HashSet<string>(
                alreadyAssigned ?? Enumerable.Empty<string>(),
                StringComparer.OrdinalIgnoreCase);

            var cableList = cableIds?
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x)
                .ToList() ?? new List<string>();

            if (cableList.Count == 0)
            {
                TxtCableHeader.Text = "Kabel: -";
            }
            else if (cableList.Count == 1)
            {
                TxtCableHeader.Text = $"Kabel: {cableList[0]}";
            }
            else
            {
                TxtCableHeader.Text = $"Kabel ({cableList.Count}): {string.Join(", ", cableList)}";
            }

            TxtSearch.TextChanged += (_, __) => RefreshList();
            BtnSelectAll.Click += (_, __) => SelectAllVisible();
            BtnClearAll.Click += (_, __) => ClearAllVisible();
            BtnSave.Click += (_, __) => SaveAndClose();
            BtnCancel.Click += (_, __) => Close();

            RefreshList();
        }

        private void RefreshList()
        {
            ItemsPanel.Children.Clear();
            _checkBoxes.Clear();

            var q = (TxtSearch.Text ?? "").Trim();

            IEnumerable<string> list = _allSperrpauses;

            if (!string.IsNullOrWhiteSpace(q))
                list = list.Where(x => x.Contains(q, StringComparison.OrdinalIgnoreCase));

            var finalList = list.ToList();

            if (finalList.Count == 0)
            {
                ItemsPanel.Children.Add(new TextBlock
                {
                    Text = "Keine Sperrpause gefunden.",
                    Foreground = System.Windows.Media.Brushes.Gray,
                    Margin = new Thickness(4)
                });
                return;
            }

            foreach (var id in finalList)
            {
                var cb = new CheckBox
                {
                    Content = id,
                    IsChecked = _selected.Contains(id),
                    Margin = new Thickness(4),
                    FontSize = 14
                };

                cb.Checked += (_, __) => _selected.Add(id);
                cb.Unchecked += (_, __) => _selected.Remove(id);

                _checkBoxes[id] = cb;
                ItemsPanel.Children.Add(cb);
            }
        }

        private void SelectAllVisible()
        {
            foreach (var kv in _checkBoxes)
            {
                kv.Value.IsChecked = true;
                _selected.Add(kv.Key);
            }
        }

        private void ClearAllVisible()
        {
            foreach (var kv in _checkBoxes)
            {
                kv.Value.IsChecked = false;
                _selected.Remove(kv.Key);
            }
        }

        private void SaveAndClose()
        {
            SelectedSperrpauses = _selected
                .OrderBy(x => x)
                .ToList();

            DialogResult = true;
            Close();
        }
    }
}