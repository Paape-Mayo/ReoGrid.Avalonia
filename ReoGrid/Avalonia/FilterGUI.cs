/*****************************************************************************
 * 
 * ReoGrid - .NET Spreadsheet Control
 * 
 * https://reogrid.net/
 *
 * THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
 * KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR
 * PURPOSE.
 *
 * Author: Jingwood <jingwood at unvell.com>
 *
 * Copyright (c) 2012-2023 Jingwood <jingwood at unvell.com>
 * Copyright (c) 2012-2023 unvell inc. All rights reserved.
 * 
 ****************************************************************************/

#if AVALONIA

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Controls.Primitives.PopupPositioning;
using Avalonia.Controls.Templates;
using Avalonia.Data;
using Avalonia.Layout;
using unvell.ReoGrid.Data;

namespace unvell.ReoGrid.AvaloniaPlatform
{
    /// <summary>
    /// Drop-down panel of a column filter: the two sorts, the column's values to show or hide,
    /// and OK and Cancel.
    /// </summary>
    public sealed class ColumnFilterContextMenu
    {
        private const double PanelWidth = 240;

        private readonly AutoColumnFilter.AutoColumnFilterBody headerBody;
        private bool syncing;

        internal ColumnFilterContextMenu(AutoColumnFilter.AutoColumnFilterBody headerBody)
        {
            this.headerBody = headerBody;

            this.SortAZButton = CreateMenuButton(LanguageResource.Filter_SortAtoZ);
            this.SortAZButton.Click += (_, _) => this.Sort(SortOrder.Ascending);
            this.SortZAButton = CreateMenuButton(LanguageResource.Filter_SortZtoA);
            this.SortZAButton.Click += (_, _) => this.Sort(SortOrder.Descending);

            this.SelectAllCheckBox = new CheckBox { Content = LanguageResource.Filter_SelectAll };
            this.SelectAllCheckBox.IsCheckedChanged += (_, _) => this.OnSelectAllChanged();

            var list = new ListBox
            {
                Height = 200,
                ItemsSource = this.Items,
                ItemTemplate = new FuncDataTemplate<FilterItem>((_, _) =>
                {
                    var check = new CheckBox();
                    check.Bind(ContentControl.ContentProperty, new ReflectionBinding(nameof(FilterItem.Text)));
                    check.Bind(ToggleButton.IsCheckedProperty, new ReflectionBinding(nameof(FilterItem.IsChecked)) { Mode = BindingMode.TwoWay });
                    return check;
                }, supportsRecycling: true),
            };

            this.OkButton = new Button { Content = LanguageResource.Button_OK, MinWidth = 72 };
            this.OkButton.Click += (_, _) => this.Apply();
            this.CancelButton = new Button { Content = LanguageResource.Button_Cancel, MinWidth = 72 };

            this.Flyout = new Flyout
            {
                Placement = PlacementMode.AnchorAndGravity,
                PlacementAnchor = PopupAnchor.TopLeft,
                PlacementGravity = PopupGravity.BottomRight,
                Content = new StackPanel
                {
                    Width = PanelWidth,
                    Spacing = 4,
                    Children =
                    {
                        this.SortAZButton,
                        this.SortZAButton,
                        new Separator(),
                        this.SelectAllCheckBox,
                        list,
                        new StackPanel
                        {
                            Orientation = Orientation.Horizontal,
                            HorizontalAlignment = HorizontalAlignment.Right,
                            Spacing = 6,
                            Children = { this.OkButton, this.CancelButton },
                        },
                    },
                },
            };

            this.CancelButton.Click += (_, _) => this.Flyout.Hide();
        }

        /// <summary>
        /// Flyout that shows the panel under the column's filter button.
        /// </summary>
        public Flyout Flyout { get; }

        /// <summary>
        /// Button that sorts the filter range by this column, A to Z.
        /// </summary>
        public Button SortAZButton { get; }

        /// <summary>
        /// Button that sorts the filter range by this column, Z to A.
        /// </summary>
        public Button SortZAButton { get; }

        /// <summary>
        /// Check box that checks or clears every value; it shows neither state while only some are checked.
        /// </summary>
        public CheckBox SelectAllCheckBox { get; }

        /// <summary>
        /// Button that applies the checked values as the filter.
        /// </summary>
        public Button OkButton { get; }

        /// <summary>
        /// Button that closes the panel without changing the filter.
        /// </summary>
        public Button CancelButton { get; }

        /// <summary>
        /// The column's distinct values, with blank cells listed last as one item.
        /// </summary>
        public ObservableCollection<FilterItem> Items { get; } = new ObservableCollection<FilterItem>();

        internal static void ShowFilterPanel(AutoColumnFilter.AutoColumnFilterBody headerBody, unvell.ReoGrid.Graphics.Point point)
        {
            var worksheet = headerBody.ColumnHeader?.Worksheet;
            if (worksheet?.ControlAdapter?.ControlInstance is not Control control) return;

            var headerRect = Views.ColumnHeaderView.GetColHeaderBounds(worksheet, headerBody.ColumnHeader!.Index, point);
            if (headerRect.Width == 0 || headerRect.Height == 0) return;

            var buttonRect = headerBody.GetColumnFilterButtonRect(headerRect.Size);

            var panel = headerBody.ContextMenu ??= new ColumnFilterContextMenu(headerBody);
            panel.Reload();
            panel.Flyout.HorizontalOffset = Math.Max(0, headerRect.Right - PanelWidth);
            panel.Flyout.VerticalOffset = headerRect.Top + buttonRect.Bottom + 1;
            panel.Flyout.ShowAt(control);
        }

        private static Button CreateMenuButton(string text) => new Button
        {
            Content = text,
            HorizontalAlignment = HorizontalAlignment.Stretch,
            HorizontalContentAlignment = HorizontalAlignment.Left,
        };

        /// <summary>
        /// Rebuild the list from the column's current values, checked as the applied filter has
        /// them, so changes cancelled last time do not come back.
        /// </summary>
        private void Reload()
        {
            var allShown = this.headerBody.IsSelectAll != false;

            foreach (var item in this.Items) item.PropertyChanged -= this.OnItemChanged;
            this.Items.Clear();

            foreach (var text in this.headerBody.GetDistinctItems() ?? new List<string>())
            {
                var isBlank = text == LanguageResource.Filter_Blanks;
                var isChecked = allShown
                    || (isBlank ? this.headerBody.ContainsBlank : this.headerBody.selectedTextItems.Contains(text));
                var item = new FilterItem(text, isBlank, isChecked);
                item.PropertyChanged += this.OnItemChanged;
                this.Items.Add(item);
            }

            this.headerBody.DataDirty = false;
            this.SyncSelectAll();
        }

        private void OnItemChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (!this.syncing) this.SyncSelectAll();
        }

        private void SyncSelectAll()
        {
            var checkedCount = this.Items.Count(item => item.IsChecked);

            this.syncing = true;
            this.SelectAllCheckBox.IsChecked = checkedCount == this.Items.Count ? true : checkedCount == 0 ? false : null;
            this.syncing = false;
        }

        private void OnSelectAllChanged()
        {
            if (this.syncing || this.SelectAllCheckBox.IsChecked is not bool check) return;

            this.syncing = true;
            foreach (var item in this.Items) item.IsChecked = check;
            this.syncing = false;
        }

        private void Apply()
        {
            if (this.Items.All(item => item.IsChecked))
            {
                this.headerBody.IsSelectAll = true;
            }
            else
            {
                this.headerBody.IsSelectAll = false;
                this.headerBody.ContainsBlank = this.Items.Any(item => item.IsBlank && item.IsChecked);
                this.headerBody.selectedTextItems.Clear();
                this.headerBody.selectedTextItems.AddRange(
                    this.Items.Where(item => !item.IsBlank && item.IsChecked).Select(item => item.Text));
            }

            this.Flyout.Hide();
            this.headerBody.autoFilter.Apply();
        }

        private void Sort(SortOrder order)
        {
            this.Flyout.Hide();

            var worksheet = this.headerBody.ColumnHeader?.Worksheet;
            if (worksheet == null) return;

            try
            {
                worksheet.SortColumn(this.headerBody.ColumnHeader!.Index, this.headerBody.autoFilter.ApplyRange, order);
            }
            catch (Exception ex)
            {
                worksheet.NotifyExceptionHappen(ex);
            }
        }
    }

    /// <summary>
    /// A value listed in a column filter's panel.
    /// </summary>
    public sealed class FilterItem : INotifyPropertyChanged
    {
        private bool isChecked;

        internal FilterItem(string text, bool isBlank, bool isChecked)
        {
            this.Text = text;
            this.IsBlank = isBlank;
            this.isChecked = isChecked;
        }

        /// <summary>
        /// Text of the value as its cells display it.
        /// </summary>
        public string Text { get; }

        /// <summary>
        /// True for the item that stands for blank cells.
        /// </summary>
        public bool IsBlank { get; }

        /// <summary>
        /// Whether rows holding this value stay shown when the filter applies.
        /// </summary>
        public bool IsChecked
        {
            get => this.isChecked;
            set
            {
                if (this.isChecked == value) return;

                this.isChecked = value;
                this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(this.IsChecked)));
            }
        }

        /// <inheritdoc/>
        public event PropertyChangedEventHandler? PropertyChanged;
    }
}

#endif
