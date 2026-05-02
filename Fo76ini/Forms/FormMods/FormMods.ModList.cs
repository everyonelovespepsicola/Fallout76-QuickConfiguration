﻿using System.Collections;
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
                Width = 300,
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

            if (Theming.CurrentTheme == ThemeType.Dark)
            {
                this.dataGridViewMods.BackgroundColor = Color.FromArgb(34, 34, 34);
                this.dataGridViewMods.GridColor = Color.FromArgb(50, 50, 50);

                this.dataGridViewMods.DefaultCellStyle.BackColor = Color.FromArgb(34, 34, 34);
                this.dataGridViewMods.DefaultCellStyle.ForeColor = Color.White;
                this.dataGridViewMods.DefaultCellStyle.SelectionBackColor = Color.FromArgb(65, 65, 65);
                this.dataGridViewMods.DefaultCellStyle.SelectionForeColor = Color.White;

                this.dataGridViewMods.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(40, 40, 40);
                this.dataGridViewMods.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                this.dataGridViewMods.EnableHeadersVisualStyles = false;
            }
        }

        public void UpdateDataGridView()
        {
            // Remember selected rows:
            List<int> selectedIndices = GetSelectedIndices();

            List<ModListRow> list = new List<ModListRow>();
            int loadOrder = 1;
            foreach (ManagedMod mod in this.Mods)
            {
                ModListRow row = new ModListRow(mod);
                if (mod.Enabled)
                    row.LoadOrder = loadOrder++;
                list.Add(row);
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

            SetSelectedIndices(selectedIndices);
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
        private List<int> GetSelectedIndices()
        {
            List<int> selectedIndices = new List<int>();
            if (this.dataGridViewMods.SelectedRows != null)
            {
                foreach (DataGridViewRow row in this.dataGridViewMods.SelectedRows)
                    selectedIndices.Add(row.Index);
            }
            return selectedIndices;
        }

        private void SetSelectedIndices(List<int> selectedIndices)
        {
            foreach (DataGridViewRow row in this.dataGridViewMods.Rows)
                row.Selected = selectedIndices.Contains(row.Index);
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

        /*
        // Enable/Disable mod on checked changed:
        private void objectListViewMods_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            if (isUpdating)
                return;

            // The e.Item.Index is the index of the item in the sorted list.
            // This is not the index of the mod in the 'Mods' list.
            // We have to get the model object and find its index.
            OLVListItem item = (OLVListItem)e.Item;
            ModListRow row = (ModListRow)item.RowObject;
            int modIndex = this.Mods.IndexOf(row.mod);

            if (e.Item.Checked)
                Mods.EnableMod(modIndex);
            else
                Mods.DisableMod(modIndex);

            UpdateUI();
        }
        */

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
            /*
            List<int> selectedIndices = new List<int>();
            foreach (ModListRow row in this.objectListViewMods.SelectedObjects)
                selectedIndices.Add(Mods.MoveModUp(Mods.IndexOf(row.mod)));
            SetSelectedIndices(selectedIndices);
            */

            List<int> selectedIndices = new List<int>();
            foreach (DataGridViewRow row in this.dataGridViewMods.SelectedRows)
            {
                selectedIndices.Add(Mods.MoveModUp(Mods.IndexOf(((ModListRow)row.DataBoundItem).mod)));
            }
            SetSelectedIndices(selectedIndices);
        }

        private void MoveSelectedModsDown()
        {
            /*
            List<int> selectedIndices = new List<int>();
            foreach (ModListRow row in this.objectListViewMods.SelectedObjects)
                selectedIndices.Add(Mods.MoveModDown(Mods.IndexOf(row.mod)));
            SetSelectedIndices(selectedIndices);
            */

            List<int> selectedIndices = new List<int>();
            foreach (DataGridViewRow row in this.dataGridViewMods.SelectedRows)
            {
                selectedIndices.Add(Mods.MoveModDown(Mods.IndexOf(((ModListRow)row.DataBoundItem).mod)));
            }
            SetSelectedIndices(selectedIndices);
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
