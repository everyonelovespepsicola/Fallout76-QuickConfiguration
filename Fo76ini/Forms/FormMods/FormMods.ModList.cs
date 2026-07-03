using System.Collections;
using Fo76ini.Interface;
using Fo76ini.Mods;
using Fo76ini.API;
using Fo76ini.Utilities;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace Fo76ini
{
    partial class FormMods
    {
        /*
         * Rendering of the list
         */
        #region Rendering

        public enum ModListStyle
        {
            Standard = 0,
            Alternative = 1
        }

        public class ModListRow
        {
            public ManagedMod mod;

            public bool Enabled { get; set; } = false;
            public bool IsUpdateAvailable { get; set; } = false;
            public int? LoadOrder { get; set; } = null;
            public DateTime? DateCreated { get; set; } = null;

            /*
             * Standard style columns
             */
            public string ModTitle { get; set; } = "";
            public string ModDescription { get; set; } = "";
            public string InstallStatus { get; set; } = "";
            public string InstallInfo { get; set; } = "";

            /*
             * Alternative style columns
             */
            public string ModVersion { get; set; } = "";
            public string InstallMethod { get; set; } = "";
            public string InstallInto { get; set; } = "";
            public string ArchiveName { get; set; } = "";
            public string ArchivePreset { get; set; } = "";
            public string IsFrozen { get; set; } = "";

            /*
             * Colors
             */
            public Color InstallStatusColor = Color.Black;
            public Color InstallMethodColor = Color.Black;
            public Color ArchivePresetColor = Color.Black;
            public Color AltFrozenColor = Color.Black;

            public ModListRow(ManagedMod mod)
            {
                this.mod = mod;
                NMMod remoteMod = mod.RemoteInfo;

                this.Enabled = mod.Enabled;

                if (Directory.Exists(mod.ManagedFolderPath))
                    this.DateCreated = Directory.GetCreationTime(mod.ManagedFolderPath);

                bool showRemoteModNames = Configuration.Mods.ShowRemoteModNames;


                /*
                 * Mod name & info
                 */

                this.ModVersion = mod.Version;

                if (remoteMod != null)
                {
                    this.ModTitle = showRemoteModNames ? remoteMod.Title : mod.Title;

                    // Version
                    if (mod.Version != "")
                    {
                        this.ModDescription += $"Version {mod.Version} by {mod.RemoteInfo.Author}";

                        if (ModHelpers.CompareVersion(mod.Version, remoteMod.LatestVersion) < 0)
                        {
                            // Update available:
                            this.ModDescription = $"{Localization.GetString("updateAvailable")}: {remoteMod.LatestVersion}";
                            this.ModVersion = $"{mod.Version} ({remoteMod.LatestVersion})";
                            this.IsUpdateAvailable = true;
                        }
                    }
                    else
                    {
                        // Author
                        this.ModDescription = "by " + mod.RemoteInfo.Author;
                    }
                }
                else
                {
                    this.ModTitle = mod.Title;

                    if (mod.Version != "")
                        this.ModDescription = $"Version {mod.Version} ";
                }


                /*
                 * Installation status:
                 */

                if (mod.IsDeploymentNecessary())
                {
                    this.InstallStatusColor = Theming.GetColor("Mod.ListPendingColor", Color.Blue);
                    if (mod.Enabled && !mod.Deployed)
                    {
                        this.InstallStatus = Localization.GetString("modTablePendingInstallation");
                    }
                    else if (!mod.Enabled && mod.Deployed)
                    {
                        this.InstallStatus = Localization.GetString("modTablePendingRemoval");
                    }
                    else
                    {
                        this.InstallStatus = Localization.GetString("modTablePendingChanges");
                    }
                    if (!mod.Frozen && mod.Freeze)
                    {
                        this.InstallStatus += $" ({Localization.GetString("modTableFreeze")})";
                    }
                }
                else
                {
                    this.InstallStatus = mod.Enabled ? Localization.GetString("enabled") : Localization.GetString("disabled");
                    this.InstallStatusColor = mod.Enabled ? Theming.GetColor("Mod.ListEnabledColor", Color.Green) : Theming.GetColor("Mod.ListDisabledColor", Color.Red);
                    if (mod.Freeze && mod.Frozen)
                    {
                        this.InstallStatus += $" ({Localization.GetString("modTableFrozen")})";
                    }
                }


                /*
                 * Installation information:
                 */

                // Which preset?
                string installPreset = "?";
                if (mod.Method == ManagedMod.DeploymentMethod.SeparateBA2)
                {
                    bool isCompressed = mod.Compression == Archive2.Compression.Default;
                    switch (mod.Format)
                    {
                        case Archive2.Format.General:
                            if (isCompressed)
                            {
                                installPreset = Localization.GetString("modsTablePresetGeneral"); // General
                                this.ArchivePresetColor = Theming.GetColor("Mod.PresetGeneralColor", Color.OrangeRed);
                            }
                            else
                            {
                                installPreset = Localization.GetString("modsTablePresetSoundFX"); // Sound FX
                                this.ArchivePresetColor = Theming.GetColor("Mod.PresetSoundColor", Color.RoyalBlue);
                            }
                            break;
                        case Archive2.Format.DDS:
                            installPreset = Localization.GetString("modsTablePresetTextures");    // Textures
                            this.ArchivePresetColor = Theming.GetColor("Mod.PresetTexturesColor", Color.DarkGreen);
                            break;
                        case null: // null means auto-detect
                            installPreset = Localization.GetString("auto");                       // Auto-detect
                            this.ArchivePresetColor = Theming.GetColor("Mod.PresetAutoColor", Color.DimGray);
                            break;
                        default:
                            installPreset = Localization.GetString("unknown");                    // Please select
                            this.ArchivePresetColor = Color.Red;
                            break;
                    }
                    if (mod.Compression == null) // null means auto-detect
                    {
                        installPreset = Localization.GetString("auto");                           // Auto-detect
                        this.ArchivePresetColor = Theming.GetColor("Mod.PresetAutoColor", Color.DimGray);
                    }

                    this.ArchivePreset = installPreset;
                }

                // How is the mod (going to be) installed?
                switch (mod.Method)
                {
                    case ManagedMod.DeploymentMethod.BundledBA2:
                        this.InstallInfo = String.Format(Localization.GetString("modTableInstallInfoBundledBA2"), "\"Data\\Bundled*.ba2\"");
                        this.InstallMethod = Localization.GetString("modsTableTypeBundled");
                        this.InstallInto = "\"Data\"";
                        this.ArchiveName = "Bundled*.ba2";
                        this.InstallMethodColor = Theming.GetColor("Mod.InstallBundledColor", Color.OrangeRed);
                        break;
                    case ManagedMod.DeploymentMethod.SeparateBA2:
                        this.InstallInfo = String.Format(Localization.GetString("modTableInstallInfoSeparateBA2"), $"\"Data\\{mod.ArchiveName}\"", $"\"{installPreset}\"");
                        this.InstallMethod = Localization.GetString("modsTableTypeSeparate");
                        this.InstallInto = "\"Data\"";
                        this.ArchiveName = mod.ArchiveName;
                        if (mod.Frozen)
                        {
                            this.IsFrozen = Localization.GetString("yes");
                            this.InstallMethod = Localization.GetString("modsTableTypeSeparateFrozen");
                            this.AltFrozenColor = Theming.GetColor("Mod.FrozenColor", Color.DarkCyan);
                        }
                        else if (mod.Freeze)
                        {
                            this.IsFrozen = Localization.GetString("modTableFrozenPending");
                            this.AltFrozenColor = Theming.GetColor("Mod.FreezePendingColor", Color.Blue);
                        }
                        else
                        {
                            this.IsFrozen = Localization.GetString("no");
                        }
                        this.InstallMethodColor = Theming.GetColor("Mod.InstallSeparateColor", Color.OrangeRed);
                        break;
                    case ManagedMod.DeploymentMethod.LooseFiles:
                        this.InstallInfo = String.Format(Localization.GetString("modTableInstallInfoLooseFiles"), $"\"{mod.RootFolder}\"");
                        this.InstallMethod = Localization.GetString("modsTableTypeLoose");
                        this.InstallInto = $"\"{mod.RootFolder}\"";
                        this.InstallMethodColor = Theming.GetColor("Mod.InstallLooseColor", Color.MediumVioletRed);
                        break;
                }
            }
        }

        public void InitializeDataGridView()
        {
            this.dataGridViewMods.AutoGenerateColumns = false;
            this.dataGridViewMods.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            this.dataGridViewMods.AllowUserToAddRows = false;
            this.dataGridViewMods.AllowUserToDeleteRows = false;
            this.dataGridViewMods.ReadOnly = true;
            this.dataGridViewMods.RowHeadersVisible = false;
            this.dataGridViewMods.Font = new Font(this.Font.FontFamily, 9.5f, FontStyle.Regular);
            this.dataGridViewMods.RowTemplate.Height = 26;

            this.dataGridViewMods.Columns.Clear();

            this.dataGridViewMods.Columns.Add(new DataGridViewCheckBoxColumn
            {
                DataPropertyName = "Enabled",
                HeaderText = "✔",
                Width = 30,
                SortMode = DataGridViewColumnSortMode.Programmatic
            });

            this.dataGridViewMods.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "LoadOrder",
                HeaderText = "Order",
                Width = 45,
                SortMode = DataGridViewColumnSortMode.Programmatic
            });

            this.dataGridViewMods.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "ModTitle",
                HeaderText = "Mod Name",
                Width = 400,
                SortMode = DataGridViewColumnSortMode.Programmatic
            });

            this.dataGridViewMods.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "InstallStatus",
                HeaderText = "Status",
                Width = 110,
                SortMode = DataGridViewColumnSortMode.Programmatic
            });

            this.dataGridViewMods.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "InstallInfo",
                HeaderText = "Installation",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                SortMode = DataGridViewColumnSortMode.Programmatic
            });

            this.dataGridViewMods.CellFormatting += dataGridViewMods_CellFormatting;
            this.dataGridViewMods.ColumnHeaderMouseClick += dataGridViewMods_ColumnHeaderMouseClick;
            this.dataGridViewMods.CellClick += dataGridViewMods_CellClick;
            this.dataGridViewMods.CellDoubleClick += dataGridViewMods_CellClick;

            this.dataGridViewMods.AllowDrop = true;
            this.dataGridViewMods.MouseMove += dataGridViewMods_MouseMove;
            this.dataGridViewMods.MouseDown += dataGridViewMods_MouseDown;
            this.dataGridViewMods.DragOver += dataGridViewMods_DragOver;
            this.dataGridViewMods.DragDrop += dataGridViewMods_DragDrop;

            ApplyGridTheme();
        }

        /// <summary>
        /// Applies the current theme's grid background colors to the mod list DataGridView.
        /// Must be called after Theming.ApplyTheme() so that the theme vars are loaded.
        /// Uses dedicated Mod.Grid.* vars to avoid colliding with text-color keys.
        /// </summary>
        public void ApplyGridTheme()
        {
            this.dataGridViewMods.BackgroundColor = Theming.GetColor("Mod.Grid.BackgroundColor", Color.White);
            this.dataGridViewMods.GridColor = Theming.GetColor("Mod.Grid.GridLineColor", Color.LightGray);

            this.dataGridViewMods.DefaultCellStyle.BackColor = Theming.GetColor("Mod.Grid.CellBackColor", Color.White);
            this.dataGridViewMods.DefaultCellStyle.ForeColor = Theming.GetColor("Mod.Grid.CellForeColor", Color.Black);
            this.dataGridViewMods.DefaultCellStyle.SelectionBackColor = Theming.GetColor("Mod.Grid.SelectionBackColor", Color.LightBlue);
            this.dataGridViewMods.DefaultCellStyle.SelectionForeColor = Theming.GetColor("Mod.Grid.SelectionForeColor", Color.Black);

            this.dataGridViewMods.AlternatingRowsDefaultCellStyle.BackColor = Theming.GetColor("Mod.Grid.AlternatingRowBackColor", Color.FromArgb(240, 240, 240));

            this.dataGridViewMods.ColumnHeadersDefaultCellStyle.BackColor = Theming.GetColor("Mod.Grid.HeaderBackColor", Color.Gainsboro);
            this.dataGridViewMods.ColumnHeadersDefaultCellStyle.ForeColor = Theming.GetColor("Mod.Grid.HeaderForeColor", Color.Black);
            this.dataGridViewMods.EnableHeadersVisualStyles = false;

            // Remove then re-add the checkbox cell painter so this method is safe to call multiple times.
            this.dataGridViewMods.CellPainting -= DataGridViewMods_PaintCheckBoxCell;
            this.dataGridViewMods.CellPainting += DataGridViewMods_PaintCheckBoxCell;
        }

        /// <summary>
        /// Paints the "Enabled" checkbox column cells with a themed green glyph
        /// instead of the default OS-rendered checkbox.
        /// </summary>
        private void DataGridViewMods_PaintCheckBoxCell(object sender, DataGridViewCellPaintingEventArgs e)
        {
            // Only intercept the checkbox column (column 0, DataPropertyName == "Enabled").
            if (e.ColumnIndex < 0 || e.RowIndex < 0) return;
            if (this.dataGridViewMods.Columns[e.ColumnIndex].DataPropertyName != "Enabled") return;

            e.Handled = true; // We take over all painting for this cell.

            // --- Background ---
            bool isSelected = (e.State & DataGridViewElementStates.Selected) != 0;
            bool isAltRow   = (e.RowIndex % 2 == 1);

            Color cellBack = isSelected
                ? Theming.GetColor("Mod.Grid.SelectionBackColor", Color.LightBlue)
                : (isAltRow
                    ? Theming.GetColor("Mod.Grid.AlternatingRowBackColor", Color.WhiteSmoke)
                    : Theming.GetColor("Mod.Grid.CellBackColor",  Color.White));

            using (var brush = new SolidBrush(cellBack))
                e.Graphics.FillRectangle(brush, e.CellBounds);

            // --- Grid line (bottom border) ---
            using (var pen = new Pen(Theming.GetColor("Mod.Grid.GridLineColor", Color.LightGray)))
                e.Graphics.DrawLine(pen,
                    e.CellBounds.Left,  e.CellBounds.Bottom - 1,
                    e.CellBounds.Right, e.CellBounds.Bottom - 1);

            // --- Custom checkbox glyph ---
            bool isChecked = (e.Value is bool b && b) ||
                             (e.Value is CheckState cs && cs == CheckState.Checked);

            Color accent = Theming.GetColor("Mod.ListEnabledColor", Color.Green);
            Color back   = cellBack;

            const int GlyphSize = 13;
            int glyphX = e.CellBounds.Left + (e.CellBounds.Width  - GlyphSize) / 2;
            int glyphY = e.CellBounds.Top  + (e.CellBounds.Height - GlyphSize) / 2;
            var glyphRect = new Rectangle(glyphX, glyphY, GlyphSize, GlyphSize);

            // Erase any residual OS rendering.
            using (var brush = new SolidBrush(back))
                e.Graphics.FillRectangle(brush, glyphRect);

            // Border box.
            using (var pen = new Pen(accent, 1f))
                e.Graphics.DrawRectangle(pen, glyphRect.X, glyphRect.Y,
                                         glyphRect.Width - 1, glyphRect.Height - 1);

            if (isChecked)
            {
                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                using (var pen = new Pen(accent, 2f))
                    e.Graphics.DrawLines(pen, new[]
                    {
                        new Point(glyphRect.X + 2,              glyphRect.Y + GlyphSize / 2 - 1),
                        new Point(glyphRect.X + GlyphSize / 2 - 1, glyphRect.Bottom - 3),
                        new Point(glyphRect.Right - 2,          glyphRect.Y + 2)
                    });
                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.Default;
            }
        }



        public void UpdateDataGridView()
        {
            if (this.Mods == null)
                return;

            // Remember selected rows:
            List<ManagedMod> selectedMods = GetSelectedMods();

            // Remember scroll position:
            int scrollIndex = this.dataGridViewMods.FirstDisplayedScrollingRowIndex;

            string searchText = this.toolStripTextBoxSearch.Text;
            if (searchText == "Search...") searchText = "";
            bool isSearching = !string.IsNullOrWhiteSpace(searchText);

            List<ModListRow> list = new List<ModListRow>();
            int loadOrder = 1;
            foreach (ManagedMod mod in this.Mods)
            {
                ModListRow row = new ModListRow(mod);
                if (mod.Enabled)
                    row.LoadOrder = loadOrder++;

                if (!isSearching ||
                    row.ModTitle.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    row.ModDescription.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    row.InstallInfo.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    list.Add(row);
                }
            }

            // Apply sorting
            if (!string.IsNullOrEmpty(currentSortColumn) && currentSortOrder != SortOrder.None)
            {
                list.Sort((x, y) =>
                {
                    int result = 0;
                    switch (currentSortColumn)
                    {
                        case "Enabled":
                            result = x.Enabled.CompareTo(y.Enabled);
                            break;
                        case "LoadOrder":
                            if (x.LoadOrder == null && y.LoadOrder == null) result = 0;
                            else if (x.LoadOrder == null) result = 1;
                            else if (y.LoadOrder == null) result = -1;
                            else result = x.LoadOrder.Value.CompareTo(y.LoadOrder.Value);
                            break;
                        case "ModTitle":
                            result = string.Compare(x.ModTitle, y.ModTitle, StringComparison.OrdinalIgnoreCase);
                            break;
                        case "InstallStatus":
                            result = string.Compare(x.InstallStatus, y.InstallStatus, StringComparison.OrdinalIgnoreCase);
                            break;
                        case "InstallInfo":
                            result = string.Compare(x.InstallInfo, y.InstallInfo, StringComparison.OrdinalIgnoreCase);
                            break;
                    }

                    // Fallback to ModTitle if same
                    if (result == 0 && currentSortColumn != "ModTitle")
                    {
                        result = string.Compare(x.ModTitle, y.ModTitle, StringComparison.OrdinalIgnoreCase);
                    }

                    return currentSortOrder == SortOrder.Ascending ? result : -result;
                });
            }

            this.dataGridViewMods.DataSource = list;

            // Set the sort glyph
            foreach (DataGridViewColumn col in this.dataGridViewMods.Columns)
            {
                if (col.DataPropertyName == currentSortColumn)
                    col.HeaderCell.SortGlyphDirection = currentSortOrder;
                else
                    col.HeaderCell.SortGlyphDirection = SortOrder.None;
            }

            SetSelectedMods(selectedMods);

            // Restore scroll position:
            if (scrollIndex >= 0 && scrollIndex < this.dataGridViewMods.Rows.Count)
            {
                this.dataGridViewMods.FirstDisplayedScrollingRowIndex = scrollIndex;
            }
        }

        private void dataGridViewMods_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= this.dataGridViewMods.Rows.Count || e.ColumnIndex < 0)
                return;

            ModListRow row = this.dataGridViewMods.Rows[e.RowIndex].DataBoundItem as ModListRow;
            if (row == null) return;

            string colName = this.dataGridViewMods.Columns[e.ColumnIndex].DataPropertyName;

            if (colName == "InstallStatus")
                e.CellStyle.ForeColor = row.InstallStatusColor;
            else if (colName == "InstallInfo")
                e.CellStyle.ForeColor = row.InstallMethodColor;
            else if (colName == "ModTitle" && row.IsUpdateAvailable)
                e.CellStyle.ForeColor = Theming.GetColor("Mod.ListUpdateAvailableColor", Color.Fuchsia);
        }

        private string currentSortColumn = "LoadOrder";
        private SortOrder currentSortOrder = SortOrder.Ascending;

        private void dataGridViewMods_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            string columnName = this.dataGridViewMods.Columns[e.ColumnIndex].DataPropertyName;

            if (currentSortColumn == columnName)
            {
                currentSortOrder = currentSortOrder == SortOrder.Ascending ? SortOrder.Descending : SortOrder.Ascending;
            }
            else
            {
                currentSortColumn = columnName;
                currentSortOrder = SortOrder.Ascending;
            }

            UpdateDataGridView();
        }
        #endregion


        /*
         * Manage the list
         * e.g. Getter/Setter
         */
        #region Managing
        private List<ManagedMod> GetSelectedMods()
        {
            List<ManagedMod> selectedMods = new List<ManagedMod>();
            if (this.dataGridViewMods.SelectedRows != null)
            {
                foreach (DataGridViewRow row in this.dataGridViewMods.SelectedRows)
                {
                    ModListRow modRow = row.DataBoundItem as ModListRow;
                    if (modRow != null)
                        selectedMods.Add(modRow.mod);
                }
            }
            return selectedMods;
        }

        private void SetSelectedMods(List<ManagedMod> selectedMods)
        {
            this.dataGridViewMods.ClearSelection();

            if (selectedMods == null || selectedMods.Count == 0)
                return;

            foreach (DataGridViewRow row in this.dataGridViewMods.Rows)
            {
                ModListRow modRow = row.DataBoundItem as ModListRow;
                if (modRow != null && selectedMods.Contains(modRow.mod))
                {
                    row.Selected = true;
                }
            }
        }

        private List<int> GetSelectedIndices()
        {
            List<int> selectedIndices = new List<int>();
            if (this.dataGridViewMods.SelectedRows != null)
            {
                foreach (DataGridViewRow row in this.dataGridViewMods.SelectedRows)
                {
                    ModListRow modRow = row.DataBoundItem as ModListRow;
                    if (modRow != null)
                        selectedIndices.Add(this.Mods.IndexOf(modRow.mod));
                }
            }
            return selectedIndices;
        }

        private void SetSelectedIndex(int selectedIndex)
        {
            foreach (DataGridViewRow row in this.dataGridViewMods.Rows)
                row.Selected = row.Index == selectedIndex;
        }

        private void SelectAll()
        {
            this.dataGridViewMods.SelectAll();
        }

        private void DeselectAll()
        {
            this.dataGridViewMods.ClearSelection();
        }

        private ModListRow GetSelectedRow()
        {
            if (this.dataGridViewMods.SelectedRows.Count == 1)
                return (ModListRow)this.dataGridViewMods.SelectedRows[0].DataBoundItem;
            return null;
        }

        private int GetSelectedRowsCount()
        {
            return this.dataGridViewMods.SelectedRows.Count;
        }

        private bool AreMultipleRowsSelected()
        {
            return GetSelectedRowsCount() > 1;
        }

        private bool IsOnlyOneRowSelected()
        {
            return GetSelectedRowsCount() == 1;
        }

        #endregion


        /*
         * When the list changes...
         */
        #region Event handler

        // Enable/Disable mod on checked changed:
        private void dataGridViewMods_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (isUpdating)
                return;

            if (e.RowIndex < 0 || e.ColumnIndex < 0)
                return;

            if (this.dataGridViewMods.Columns[e.ColumnIndex].DataPropertyName == "Enabled")
            {
                ModListRow row = (ModListRow)this.dataGridViewMods.Rows[e.RowIndex].DataBoundItem;
                row.mod.Enabled = !row.mod.Enabled;

                UpdateUI();
                Mods.Save();
            }
        }

        // Mod(s) selected
        bool suppressSelectionChangedEventOnce = false;
        private void dataGridViewMods_SelectionChanged(object sender, EventArgs e)
        {
            if (isUpdating)
                return;

            /*
             * isUpdating workaround is no longer working...
             * We cannot detect whether the user or the program has changed the selection...
             * So the only thing we can do is to "suppress" this event, if the list gets updated.
             * Otherwise, the program will behave in weird ways.
             */
            if (suppressSelectionChangedEventOnce)
            {
                suppressSelectionChangedEventOnce = false;
                return;
            }

            List<ManagedMod> mods = new List<ManagedMod>();
            if (this.dataGridViewMods.SelectedRows != null)
            {
                foreach (DataGridViewRow row in this.dataGridViewMods.SelectedRows)
                {
                    ModListRow modRow = row.DataBoundItem as ModListRow;
                    if (modRow != null)
                        mods.Add(modRow.mod);
                }
            }
            this.EditMods(mods);
        }

        #endregion

        /*
         * Drag & drop to reorder rows and install files
         */
        #region Drag & drop

        private Rectangle dragBoxFromMouseDown;
        private int rowIndexFromMouseDown = -1;
        private int rowIndexOfItemUnderMouseToDrop = -1;

        private void dataGridViewMods_MouseMove(object sender, MouseEventArgs e)
        {
            if ((e.Button & MouseButtons.Left) == MouseButtons.Left)
            {
                if (dragBoxFromMouseDown != Rectangle.Empty && !dragBoxFromMouseDown.Contains(e.X, e.Y))
                {
                    DragDropEffects dropEffect = this.dataGridViewMods.DoDragDrop(this.dataGridViewMods.Rows[rowIndexFromMouseDown], DragDropEffects.Move);
                }
            }
        }

        private void dataGridViewMods_MouseDown(object sender, MouseEventArgs e)
        {
            rowIndexFromMouseDown = this.dataGridViewMods.HitTest(e.X, e.Y).RowIndex;
            if (rowIndexFromMouseDown != -1)
            {
                Size dragSize = SystemInformation.DragSize;
                dragBoxFromMouseDown = new Rectangle(new Point(e.X - (dragSize.Width / 2), e.Y - (dragSize.Height / 2)), dragSize);
            }
            else
                dragBoxFromMouseDown = Rectangle.Empty;
        }

        private void dataGridViewMods_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(DataGridViewRow)))
                e.Effect = DragDropEffects.Move;
            else if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
            else
                e.Effect = DragDropEffects.None;
        }

        private void dataGridViewMods_DragDrop(object sender, DragEventArgs e)
        {
            Point clientPoint = this.dataGridViewMods.PointToClient(new Point(e.X, e.Y));
            rowIndexOfItemUnderMouseToDrop = this.dataGridViewMods.HitTest(clientPoint.X, clientPoint.Y).RowIndex;

            // 1. Reordering rows within the grid
            if (e.Data.GetDataPresent(typeof(DataGridViewRow)))
            {
                if (rowIndexOfItemUnderMouseToDrop == -1)
                    rowIndexOfItemUnderMouseToDrop = this.dataGridViewMods.Rows.Count - 1;

                if (e.Effect == DragDropEffects.Move && rowIndexFromMouseDown != -1 && rowIndexOfItemUnderMouseToDrop != -1 && rowIndexFromMouseDown != rowIndexOfItemUnderMouseToDrop)
                {
                    // Force the list to be sorted by LoadOrder so the visual order matches the underlying intent
                    if (currentSortColumn != "LoadOrder" || currentSortOrder != SortOrder.Ascending)
                    {
                        currentSortColumn = "LoadOrder";
                        currentSortOrder = SortOrder.Ascending;
                    }

                    DataGridViewRow rowToMove = e.Data.GetData(typeof(DataGridViewRow)) as DataGridViewRow;
                    ModListRow draggedModRow = rowToMove.DataBoundItem as ModListRow;

                    if (draggedModRow != null)
                    {
                        List<ManagedMod> rebuiltList = new List<ManagedMod>();
                        foreach (DataGridViewRow row in this.dataGridViewMods.Rows)
                            if (row.Index != rowIndexFromMouseDown) // Add everything except the dragged row
                                rebuiltList.Add(((ModListRow)row.DataBoundItem).mod);

                        if (rowIndexOfItemUnderMouseToDrop >= rebuiltList.Count)
                            rebuiltList.Add(draggedModRow.mod); // Drop at the bottom
                        else
                            rebuiltList.Insert(rowIndexOfItemUnderMouseToDrop, draggedModRow.mod); // Insert at cursor

                        if (rebuiltList.Count == this.Mods.Mods.Count)
                        {
                            this.Mods.Clear();
                            this.Mods.Mods.AddRange(rebuiltList);
                            UpdateModList();
                            Mods.Save();
                            SetSelectedMods(new List<ManagedMod> { draggedModRow.mod });
                        }
                    }
                }
            }
            // 2. Dropping mod files from your PC into the grid
            else if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files != null && files.Length > 0)
                    InstallBulkThreaded(files, rowIndexOfItemUnderMouseToDrop == -1 ? this.Mods.Count : rowIndexOfItemUnderMouseToDrop);
            }
        }
        #endregion

        /*
         * Some actions
         */
        #region Actions

        private void OpenSelectedModsFolder()
        {
            // Open the folder if one mod is selected:
            if (IsOnlyOneRowSelected())
            {
                string path = GetSelectedRow().mod.ManagedFolderPath;
                if (Directory.Exists(path))
                    Utils.OpenExplorer(path);
                else
                    MsgBox.Get("modDirNotExist").FormatText(path).Show(MessageBoxIcon.Error);
            }
            // Otherwise open the parent folder (either if none or multiple rows are selected):
            else
            {
                string path = Path.Combine(this.game.GamePath, "Mods");
                if (Directory.Exists(path))
                    Utils.OpenExplorer(path);
            }
        }

        private void MoveSelectedModsUp()
        {
            var selectedRows = new List<DataGridViewRow>();
            foreach (DataGridViewRow row in this.dataGridViewMods.SelectedRows)
                selectedRows.Add(row);
            selectedRows.Sort((r1, r2) => r1.Index.CompareTo(r2.Index));

            foreach (DataGridViewRow row in selectedRows)
            {
                Mods.MoveModUp(Mods.IndexOf(((ModListRow)row.DataBoundItem).mod));
            }
        }

        private void MoveSelectedModsDown()
        {
            var selectedRows = new List<DataGridViewRow>();
            foreach (DataGridViewRow row in this.dataGridViewMods.SelectedRows)
                selectedRows.Add(row);
            selectedRows.Sort((r1, r2) => r2.Index.CompareTo(r1.Index)); // Reverse order for moving down

            foreach (DataGridViewRow row in selectedRows)
            {
                Mods.MoveModDown(Mods.IndexOf(((ModListRow)row.DataBoundItem).mod));
            }
        }

        private void ToggleCheckboxes()
        {
            // Behavior:
            //  - If at least one mod is unchecked, check all boxes.
            //  - If ALL mods are checked, uncheck all boxes.

            bool state = true;
            foreach (ManagedMod mod in Mods)
                if (!mod.Enabled)
                    state = false;

            if (!state)
            {
                DialogResult res = MessageBox.Show(
                    "Are you sure you want to enable all mods?",
                    "Enable All Mods",
                    MessageBoxButtons.OKCancel,
                    MessageBoxIcon.Question);
                if (res != DialogResult.OK)
                    return;
            }

            foreach (ManagedMod mod in Mods)
                mod.Enabled = !state;
        }

        private void DeleteSelectedMods()
        {
            if (IsOnlyOneRowSelected())
            {
                ManagedMod mod = GetSelectedRow().mod;
                DialogResult res = MsgBox.Get("deleteQuestion").FormatTitle(mod.Title).FormatText(mod.Title).Show(MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (res == DialogResult.Yes)
                    DeleteModThreaded(Mods.IndexOf(mod));
            }
            else if (AreMultipleRowsSelected())
            {
                string count = GetSelectedRowsCount().ToString();
                DialogResult res = MsgBox.Get("deleteMultipleQuestion").FormatTitle(count).FormatText(count).Show(MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (res == DialogResult.Yes)
                    DeleteModsBulkThreaded(GetSelectedIndices());
            }
        }

        private void FreezeSelectedMods()
        {
            // foreach (ModListRow row in this.objectListViewMods.SelectedObjects)
            //     row.mod.Freeze = true;

            foreach (DataGridViewRow row in this.dataGridViewMods.SelectedRows)
            {
                ((ModListRow)row.DataBoundItem).mod.Freeze = true;
            }
        }

        private void UnfreezeSelectedMods()
        {
            // foreach (ModListRow row in this.objectListViewMods.SelectedObjects)
            // {
            //     ModActions.Unfreeze(row.mod);
            //     row.mod.Freeze = false;
            // }

            foreach (DataGridViewRow row in this.dataGridViewMods.SelectedRows)
            {
                ModListRow modRow = (ModListRow)row.DataBoundItem;
                ModActions.Unfreeze(modRow.mod);
                modRow.mod.Freeze = false;
            }
        }

        #endregion
    }
}
