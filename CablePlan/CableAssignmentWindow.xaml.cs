using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace CablePlan
{
    public partial class CableAssignmentWindow : Window
    {
        private readonly List<string> _allCableIds;
        private readonly HashSet<string> _selected;
        private readonly Dictionary<string, CheckBox> _checkBoxes =
            new(StringComparer.OrdinalIgnoreCase);

        public List<string> SelectedCableIds { get; private set; } = new();

        public CableAssignmentWindow(
            string sperrpauseId,
            IEnumerable<string> availableCableIds,
            IEnumerable<string> alreadyAssignedCableIds)
        {
            InitializeComponent();

            TxtSperrpauseHeader.Text = $"Sperrpause: {sperrpauseId}";

            _allCableIds = availableCableIds
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x)
                .ToList();

            _selected = new HashSet<string>(
                alreadyAssignedCableIds ?? Enumerable.Empty<string>(),
                StringComparer.OrdinalIgnoreCase);

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

            IEnumerable<string> list = _allCableIds;
            if (!string.IsNullOrWhiteSpace(q))
                list = list.Where(x => x.Contains(q, StringComparison.OrdinalIgnoreCase));

            var finalList = list.ToList();

            if (finalList.Count == 0)
            {
                ItemsPanel.Children.Add(new TextBlock
                {
                    Text = "Keine Kabel gefunden.",
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
            SelectedCableIds = _selected
                .OrderBy(x => x)
                .ToList();

            DialogResult = true;
            Close();
        }
    }
}
