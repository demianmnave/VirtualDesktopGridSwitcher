﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VirtualDesktopGridSwitcher.Settings {
    public partial class SettingsDialog : Form {

        private SettingValues settings;

        public SettingsDialog(SettingValues settings) {
            
            this.settings = settings;

            InitializeComponent();

            LoadValues();
        }

        protected override void OnVisibleChanged(EventArgs e) {
            base.OnVisibleChanged(e);

            if (this.Visible) {
                LoadValues();
            }
        }

        private void LoadValues() {
            textBoxRows.Text = settings.Rows.ToString();
            textBoxColumns.Text = settings.Columns.ToString();

            checkBoxWrapAround.Checked = settings.WrapAround;

            checkBoxCtrlModifier.Checked = settings.CtrlModifier;
            checkBoxWinModifier.Checked = settings.WinModifier;
            checkBoxAltModifier.Checked = settings.AltModifier;
            checkBoxShiftModifier.Checked = settings.ShiftModifier;

            checkBoxFKeys.Checked = settings.FKeysForNumbers;
        }

        private bool SaveValues() {
            try {
                var rows = int.Parse(textBoxRows.Text);
                var cols = int.Parse(textBoxColumns.Text);

                if (rows * cols > 20) {
                    var result = 
                        MessageBox.Show(this,
                                        (rows * cols) + " desktops is not recommended for windows performance. Continue?", 
                                        "Warning",
                                        MessageBoxButtons.YesNo,
                                        MessageBoxIcon.Warning);
                    if (result != DialogResult.Yes) {
                        return false;
                    }
                }

                if (rows * cols < settings.Rows * settings.Columns) {
                    MessageBox.Show(this, "Unrequired desktops will not be removed");
                } else if (rows * cols > settings.Rows * settings.Columns) {
                    MessageBox.Show(this, "More desktops will be created to fill the grid if necessary");
                }

                settings.Rows = rows;
                settings.Columns = cols;
            } catch {
                MessageBox.Show(this, "Values for Rows and Columns must be numbers only");
                return false;
            }

            settings.WrapAround = checkBoxWrapAround.Checked;

            settings.CtrlModifier = checkBoxCtrlModifier.Checked;
            settings.WinModifier = checkBoxWinModifier.Checked;
            settings.AltModifier = checkBoxAltModifier.Checked;
            settings.ShiftModifier = checkBoxShiftModifier.Checked;

            settings.FKeysForNumbers = checkBoxFKeys.Checked;

            if (!settings.ApplySettings()) {
                MessageBox.Show(this, "Failed to apply settings", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (!settings.Save()) {
                MessageBox.Show(this, "Failed to save settings to file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            return true;
        }

        private void buttonOK_Click(object sender, EventArgs e) {
            SaveValues();
        }
    }
}
