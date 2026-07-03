using FastColoredTextBoxNS;
using Fo76ini.Controls;
using Fo76ini.Properties;
using Fo76ini.Utilities;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Fo76ini.Interface
{
    public enum ThemeType
    {
        Dark,
        Light,
        System
    }

    public class Theming
    {
        private static ThemeType _currentTheme;
        public static ThemeType CurrentTheme { get { return _currentTheme; } }

        private static Dictionary<string, string> _vars = new Dictionary<string, string>();

        public static String ThemesPath = Path.Combine(Shared.AppConfigFolder, "themes");

        public static Theme DarkTheme
        {
            get
            {
                String themePath = Path.Combine(ThemesPath, "dark.yml");
                if (File.Exists(themePath))
                    return Theme.ReadThemeFromFile(themePath);
                else
                    return Theme.ReadThemeFromString(Localization.GetTextResource("dark.yml"));
            }
        }

        public static Theme LightTheme
        {
            get
            {
                String themePath = Path.Combine(ThemesPath, "light.yml");
                if (File.Exists(themePath))
                    return Theme.ReadThemeFromFile(themePath);
                else
                    return Theme.ReadThemeFromString(Localization.GetTextResource("light.yml"));
            }
        }

        public static Theme SystemTheme
        {
            get { return DetectSystemTheme() == ThemeType.Dark ? DarkTheme : LightTheme; }
        }

        public static ThemeType DetectSystemTheme ()
        {
            try
            {
                RegistryKey registry =
                    Registry.CurrentUser.OpenSubKey(
                        @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");

                if (registry == null)
                    return ThemeType.Light;
                else if ((int)registry.GetValue("AppsUseLightTheme") == 1) // SystemUsesLightTheme
                    return ThemeType.Light;
                else
                    return ThemeType.Dark;
            }
            catch
            {
                return ThemeType.Light;
            }
        }

        #region Theme variables

        public static string Get(string key)
        {
            return _vars[key];
        }

        public static string Get(string key, string defaultValue)
        {
            if (_vars.ContainsKey(key))
                return _vars[key];
            return defaultValue;
        }

        public static Color GetColor(string key)
        {
            return Utils.ParseColor(Get(key));
        }

        public static Color GetColor(string key, string defaultValue)
        {
            return Utils.ParseColor(Get(key, defaultValue));
        }

        public static Color GetColor(string key, Color defaultValue)
        {
            if (_vars.ContainsKey(key))
                return Utils.ParseColor(_vars[key]);
            return defaultValue;
        }

        #endregion

        public static void ApplyTheme(ThemeType theme, Control control)
        {
            try
            {
                switch (theme)
                {
                    case ThemeType.Dark:
                        Theming.ApplyTheme(Theming.DarkTheme, control);
                        break;
                    case ThemeType.Light:
                        Theming.ApplyTheme(Theming.LightTheme, control);
                        break;
                    case ThemeType.System:
                        Theming.ApplyTheme(Theming.SystemTheme, control);
                        break;
                }
                Configuration.Appearance.AppTheme = theme;
            }
            catch (Exception e)
            {
                MsgBox.Show($"Error: Couldn't apply '{theme}' theme.", e.ToString(), MessageBoxIcon.Error);
            }
        }

        protected static void ApplyTheme(Theme theme, Control.ControlCollection controls)
        {
            foreach (Control control in controls)
                ApplyTheme(theme, control);
        }

        protected static void ApplyTheme(Theme theme, ToolStripItemCollection items)
        {
            foreach (ToolStripItem item in items)
                ApplyTheme(theme, item);
        }

        protected static void ApplyTheme(Theme theme, Control control)
        {
            _currentTheme = theme.Type;
            _vars = theme.Vars;

            String controlType = control.GetType().Name;
            String controlName = control.Name;
            String tag = "Default";
            if (control.Tag != null)
                tag = control.Tag.ToString();

            // Apply "Default" styles first:
            foreach (VisualStyle style in theme.GetDefaultStylesForControl(control))
                ApplyStyle(style, control);

            // then apply specialized styles because they have precedence:
            foreach (VisualStyle style in theme.GetSpecializedStylesForControl(control))
                ApplyStyle(style, control);

            // Hook themed glyph painting for CheckBox and RadioButton.
            // ForeColor is already set to green by the YAML above, so we use it
            // as the accent color when drawing the custom glyph.
            if (control is CheckBox || control is RadioButton)
                EnsureGlyphPainted(control);

            /*
             * Special controls and components:
             */

            // Apply theme to the contextmenu, if available:
            if (control.ContextMenuStrip != null)
                ApplyTheme(theme, control.ContextMenuStrip);

            if (control is MenuStrip)
                ApplyTheme(theme, ((MenuStrip)control).Items);

            if (control is ToolStrip)
                ApplyTheme(theme, ((ToolStrip)control).Items);

            // Apply theme to the tool tip, if available:
            if (control is IExposeComponents)
                ApplyTheme(theme, ((IExposeComponents)control).ToolTip);


            // Apply theme to children:
            ApplyTheme(theme, control.Controls);
        }

        protected static void ApplyTheme(Theme theme, Component comp)
        {
            String controlType = comp.GetType().Name;

            foreach (VisualStyle style in theme.Styles)
            {
                string controlRegex = Utils.WildCardToRegular(style.ControlType);
                if (Regex.IsMatch(controlType, controlRegex))
                    ApplyStyle(style, comp);
            }

            if (comp is ToolStripMenuItem)
                ApplyTheme(theme, ((ToolStripMenuItem)comp).DropDownItems);
        }

        private static void ApplyStyle(VisualStyle style, Control control)
        {
            if (style == null || control == null)
                return;

            foreach (KeyValuePair<String, object> rule in style.Rules)
            {
                PropertyInfo property = control.GetType().GetProperty(rule.Key);
                object parent = control;

                // Support Button.FlatAppearance
                // e.g. Button.FlatAppearance.BorderColor
                // =>
                // Button:
                //     BorderColor: "#333"
                if (property == null && control is Button)
                {
                    parent = ((Button)control).FlatAppearance;
                    property = ((Button)control).FlatAppearance.GetType().GetProperty(rule.Key);
                }

                SetProperty(property, parent, rule.Value.ToString());
            }
        }

        private static void ApplyStyle(VisualStyle style, Component comp)
        {
            if (style == null || comp == null)
                return;

            foreach (KeyValuePair<String, object> rule in style.Rules)
            {
                PropertyInfo property = comp.GetType().GetProperty(rule.Key);
                object parent = comp;

                SetProperty(property, comp, rule.Value.ToString());
            }
        }

        #region Themed glyph painting (CheckBox / RadioButton)

        // Tracks controls that already have a Paint handler attached so we never
        // add a second one even when ApplyTheme is called multiple times.
        private static readonly HashSet<Control> _glyphThemedControls = new HashSet<Control>();

        private static void EnsureGlyphPainted(Control control)
        {
            if (_glyphThemedControls.Contains(control))
            {
                // Handler already wired – just trigger a repaint so the new
                // ForeColor (which changed with the theme) is reflected.
                control.Invalidate();
                return;
            }

            _glyphThemedControls.Add(control);
            // Clean up when the control is disposed so we don't leak.
            control.Disposed += (s, e) => _glyphThemedControls.Remove((Control)s);

            if (control is CheckBox cb)
                cb.Paint += OnCheckBoxPaint;
            else if (control is RadioButton rb)
                rb.Paint += OnRadioButtonPaint;

            control.Invalidate();
        }

        private static void OnCheckBoxPaint(object sender, PaintEventArgs e)
        {
            CheckBox cb = (CheckBox)sender;
            Color accent = cb.ForeColor;
            Color back   = GetEffectiveBackColor(cb);
            Graphics g   = e.Graphics;

            // 13×13 logical units — the standard WinForms checkbox glyph size.
            int size    = cb.LogicalToDeviceUnits(13);
            int glyphY  = (cb.Height - size) / 2;
            var glyphRect = new Rectangle(1, glyphY, size, size);

            // Completely erase the OS-rendered glyph.
            using (var brush = new SolidBrush(back))
                g.FillRectangle(brush, glyphRect.X - 1, glyphRect.Y - 1,
                                glyphRect.Width + 2, glyphRect.Height + 2);

            // Border box.
            using (var pen = new Pen(accent, 1f))
                g.DrawRectangle(pen, glyphRect.X, glyphRect.Y,
                                glyphRect.Width - 1, glyphRect.Height - 1);

            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            if (cb.CheckState == CheckState.Checked)
            {
                // V-shaped checkmark.
                using (var pen = new Pen(accent, 2f))
                    g.DrawLines(pen, new[]
                    {
                        new Point(glyphRect.X + 2,                    glyphRect.Y + size / 2 - 1),
                        new Point(glyphRect.X + size / 2 - 1,         glyphRect.Bottom - 3),
                        new Point(glyphRect.Right - 2,                glyphRect.Y + 2)
                    });
            }
            else if (cb.CheckState == CheckState.Indeterminate)
            {
                // Dash for indeterminate state.
                int pad = 3;
                using (var brush = new SolidBrush(accent))
                    g.FillRectangle(brush,
                        glyphRect.X + pad,
                        glyphRect.Y + (glyphRect.Height - 2) / 2,
                        glyphRect.Width - pad * 2, 2);
            }
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.Default;
        }

        private static void OnRadioButtonPaint(object sender, PaintEventArgs e)
        {
            RadioButton rb = (RadioButton)sender;
            Color accent = rb.ForeColor;
            Color back   = GetEffectiveBackColor(rb);
            Graphics g   = e.Graphics;

            int size    = rb.LogicalToDeviceUnits(13);
            int glyphY  = (rb.Height - size) / 2;
            var glyphRect = new Rectangle(1, glyphY, size, size);

            // Erase OS glyph.
            using (var brush = new SolidBrush(back))
                g.FillRectangle(brush, glyphRect.X - 1, glyphRect.Y - 1,
                                glyphRect.Width + 2, glyphRect.Height + 2);

            // Outer circle.
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            using (var pen = new Pen(accent, 1f))
                g.DrawEllipse(pen, glyphRect.X, glyphRect.Y,
                              glyphRect.Width - 1, glyphRect.Height - 1);

            // Filled bullet when selected.
            if (rb.Checked)
            {
                int pad = 3;
                var bullet = new Rectangle(
                    glyphRect.X + pad, glyphRect.Y + pad,
                    glyphRect.Width  - pad * 2 - 1,
                    glyphRect.Height - pad * 2 - 1);
                using (var brush = new SolidBrush(accent))
                    g.FillEllipse(brush, bullet);
            }
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.Default;
        }

        /// <summary>Walks up the parent chain to find the first non-transparent BackColor.</summary>
        private static Color GetEffectiveBackColor(Control control)
        {
            Control c = control;
            while (c != null && (c.BackColor == Color.Transparent || c.BackColor.A == 0))
                c = c.Parent;
            return c?.BackColor ?? SystemColors.Control;
        }

        #endregion

        private static void SetProperty(PropertyInfo property, object parent, string value)
        {
            if (property != null)
            {
                // e.g. convert "var(BackColor)" => "#333"
                if (value.StartsWith("var"))
                {
                    int left = value.IndexOf('(');
                    int right = value.IndexOf(')');

                    if (left >= 0 && right >= 0)
                    {
                        string varName = value.Substring(left + 1, right - left - 1);
                        value = Get(varName, value);
                    }
                }

                // Convert the string value to the type of the property, then set the property's value.
                if (property.PropertyType == typeof(string))
                    property.SetValue(parent, value, null);
                else if (property.PropertyType == typeof(int))
                    property.SetValue(parent, Convert.ToInt32(value), null);
                else if (property.PropertyType == typeof(bool))
                    property.SetValue(parent, Convert.ToBoolean(value), null);
                else if (property.PropertyType == typeof(Color))
                    property.SetValue(parent, Utils.ParseColor(value), null);
                else if (property.PropertyType == typeof(Image))
                {
                    Image img = (Image)Resources.ResourceManager.GetObject(value);
                    property.SetValue(parent, img, null);
                }
                else if (property.PropertyType == typeof(FlatStyle))
                {
                    if (Enum.TryParse(value, out FlatStyle e))
                        property.SetValue(parent, e, null);
                }
                else if (property.PropertyType == typeof(BorderStyle))
                {
                    if (Enum.TryParse(value, out BorderStyle e))
                        property.SetValue(parent, e, null);
                }
            }
        }
    }
}
