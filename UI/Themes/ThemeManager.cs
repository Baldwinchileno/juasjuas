using System;
using System.Drawing;
using System.Windows.Forms;

namespace AdminSERMAC.Core.Theme
{
    public static class ThemeManager
    {
        private static bool isDarkMode = false;
        public static event EventHandler ThemeChanged;

        public static bool IsDarkMode
        {
            get => isDarkMode;
            set
            {
                if (isDarkMode != value)
                {
                    isDarkMode = value;
                    ThemeChanged?.Invoke(null, EventArgs.Empty);
                    SaveThemePreference();
                }
            }
        }

        private static Color GetBackgroundColor() => IsDarkMode ? Color.FromArgb(33, 37, 41) : Color.White;
        private static Color GetTextColor() => IsDarkMode ? Color.FromArgb(248, 249, 250) : Color.Black;
        private static Color GetPrimaryColor() => IsDarkMode ? Color.FromArgb(0, 123, 255) : Color.FromArgb(0, 122, 204);
        private static Color GetSecondaryColor() => IsDarkMode ? Color.FromArgb(108, 117, 125) : Color.FromArgb(108, 117, 125);
        private static Color GetDangerColor() => IsDarkMode ? Color.FromArgb(220, 53, 69) : Color.FromArgb(220, 53, 69);
        private static Color GetGridHeaderColor() => IsDarkMode ? Color.FromArgb(52, 58, 64) : Color.FromArgb(240, 240, 240);
        private static Color GetGridAlternateColor() => IsDarkMode ? Color.FromArgb(44, 48, 52) : Color.FromArgb(249, 249, 249);

        public static void ApplyTheme(Form form)
        {
            ApplyThemeToControl(form);

            foreach (Control control in form.Controls)
            {
                ApplyThemeToControl(control);
            }
        }

        private static void ApplyThemeToControl(Control control)
        {
            control.BackColor = GetBackgroundColor();
            control.ForeColor = GetTextColor();

            switch (control)
            {
                case Button button:
                    ApplyButtonTheme(button);
                    break;
                case DataGridView grid:
                    ApplyGridTheme(grid);
                    break;
                case TextBox textBox:
                    ApplyTextBoxTheme(textBox);
                    break;
                case ComboBox comboBox:
                    ApplyComboBoxTheme(comboBox);
                    break;
            }

            foreach (Control child in control.Controls)
            {
                ApplyThemeToControl(child);
            }
        }

        private static void ApplyButtonTheme(Button button)
        {
            if (button.Tag?.ToString() == "primary")
            {
                button.BackColor = GetPrimaryColor();
                button.ForeColor = Color.White;
            }
            else if (button.Tag?.ToString() == "danger")
            {
                button.BackColor = GetDangerColor();
                button.ForeColor = Color.White;
            }

            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = 0;
        }

        private static void ApplyGridTheme(DataGridView grid)
        {
            grid.BackgroundColor = GetBackgroundColor();
            grid.GridColor = GetSecondaryColor();
            grid.DefaultCellStyle.BackColor = GetBackgroundColor();
            grid.DefaultCellStyle.ForeColor = GetTextColor();
            grid.ColumnHeadersDefaultCellStyle.BackColor = GetGridHeaderColor();
            grid.ColumnHeadersDefaultCellStyle.ForeColor = GetTextColor();
            grid.AlternatingRowsDefaultCellStyle.BackColor = GetGridAlternateColor();
            grid.EnableHeadersVisualStyles = false;
        }

        private static void ApplyTextBoxTheme(TextBox textBox)
        {
            textBox.BorderStyle = BorderStyle.FixedSingle;
            textBox.BackColor = GetBackgroundColor();
            textBox.ForeColor = GetTextColor();
        }

        private static void ApplyComboBoxTheme(ComboBox comboBox)
        {
            comboBox.BackColor = GetBackgroundColor();
            comboBox.ForeColor = GetTextColor();
        }

        private static void SaveThemePreference()
        {
            try
            {
                Properties.Settings.Default.IsDarkMode = IsDarkMode;
                Properties.Settings.Default.Save();
            }
            catch
            {
                // Manejar el caso donde las Settings no están disponibles
            }
        }

        public static void LoadThemePreference()
        {
            try
            {
                IsDarkMode = Properties.Settings.Default.IsDarkMode;
            }
            catch
            {
                IsDarkMode = false; // Valor por defecto
            }
        }
    }
}