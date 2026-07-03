using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace ExcelCommonTools.UI
{
    /// <summary>
    /// Shared styling utilities for all dialog windows.
    /// Provides a consistent modern look across the add-in.
    /// </summary>
    internal static class DialogStyleHelper
    {
        // ── Color palette ──────────────────────────────────────────────
        public static readonly Color AccentColor    = Color.FromArgb(66, 133, 244);   // Google-blue
        public static readonly Color AccentDark     = Color.FromArgb(53, 106, 203);   // hover/press
        public static readonly Color HeaderBg       = Color.FromArgb(240, 244, 252);  // very light blue
        public static readonly Color HeaderBorder   = Color.FromArgb(200, 215, 240);  // subtle line
        public static readonly Color FormBg         = Color.FromArgb(255, 255, 255);  // white
        public static readonly Color GroupBorder    = Color.FromArgb(218, 225, 238);  // light blue-gray
        public static readonly Color TextPrimary    = Color.FromArgb(32, 33, 36);     // near-black
        public static readonly Color TextSecondary  = Color.FromArgb(95, 99, 104);    // gray
        public static readonly Color BtnCancelBg    = Color.FromArgb(241, 243, 244);  // light gray
        public static readonly Color BtnCancelHover = Color.FromArgb(226, 229, 233);
        public static readonly Color InputBorder    = Color.FromArgb(200, 210, 225);
        public static readonly Color InputFocus     = Color.FromArgb(66, 133, 244);

        // ── Font ───────────────────────────────────────────────────────
        private static Font _baseFont;
        private static Font _headerFont;
        private static Font _headerDescFont;
        private static Font _buttonFont;
        private static Font _labelFont;

        public static Font BaseFont
        {
            get
            {
                if (_baseFont == null)
                    _baseFont = CreateSafeFont("Microsoft YaHei UI", 9f);
                return _baseFont;
            }
        }

        public static Font HeaderFont
        {
            get
            {
                if (_headerFont == null)
                    _headerFont = CreateSafeFont("Microsoft YaHei UI", 12f, FontStyle.Bold);
                return _headerFont;
            }
        }

        public static Font HeaderDescFont
        {
            get
            {
                if (_headerDescFont == null)
                    _headerDescFont = CreateSafeFont("Microsoft YaHei UI", 9f, FontStyle.Regular);
                return _headerDescFont;
            }
        }

        public static Font ButtonFont
        {
            get
            {
                if (_buttonFont == null)
                    _buttonFont = CreateSafeFont("Microsoft YaHei UI", 9f);
                return _buttonFont;
            }
        }

        public static Font LabelFont
        {
            get
            {
                if (_labelFont == null)
                    _labelFont = CreateSafeFont("Microsoft YaHei UI", 9f);
                return _labelFont;
            }
        }

        private static Font CreateSafeFont(string familyName, float size, FontStyle style = FontStyle.Regular)
        {
            try
            {
                var f = new Font(familyName, size, style);
                // If the system doesn't have the font, it falls back — that's fine
                return f;
            }
            catch
            {
                return new Font(SystemFonts.DefaultFont.FontFamily, size, style);
            }
        }

        // ── Form-level styling ─────────────────────────────────────────

        /// <summary>
        /// Apply base styling to a dialog form (background, font).
        /// Call this in the constructor after setting basic form properties.
        /// </summary>
        public static void ApplyFormStyle(Form form)
        {
            form.BackColor = FormBg;
            form.Font = BaseFont;
            form.ForeColor = TextPrimary;
        }

        /// <summary>
        /// Add a styled header panel at the top of the form.
        /// Returns the Y position below the header for placing controls.
        /// </summary>
        public static int AddHeader(Form form, string title, string description = null)
        {
            int headerHeight = description != null ? 60 : 46;

            var headerPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = headerHeight,
                BackColor = HeaderBg
            };
            headerPanel.Paint += HeaderPanel_Paint;

            var lblTitle = new Label
            {
                Text = title,
                Font = HeaderFont,
                ForeColor = AccentDark,
                Location = new Point(18, description != null ? 8 : 12),
                AutoSize = true,
                BackColor = Color.Transparent
            };
            headerPanel.Controls.Add(lblTitle);

            if (description != null)
            {
                var lblDesc = new Label
                {
                    Text = description,
                    Font = HeaderDescFont,
                    ForeColor = TextSecondary,
                    Location = new Point(18, 30),
                    AutoSize = true,
                    BackColor = Color.Transparent
                };
                headerPanel.Controls.Add(lblDesc);
            }

            form.Controls.Add(headerPanel);
            // Bring header to front so it paints over other docked controls
            headerPanel.BringToFront();

            return headerHeight;
        }

        private static void HeaderPanel_Paint(object sender, PaintEventArgs e)
        {
            var panel = (Panel)sender;
            // Draw bottom border line
            using (var pen = new Pen(HeaderBorder, 1))
            {
                e.Graphics.DrawLine(pen, 0, panel.Height - 1, panel.Width, panel.Height - 1);
            }
        }

        // ── Button styling ─────────────────────────────────────────────

        /// <summary>
        /// Style a button as the primary (accent) action button.
        /// </summary>
        public static void StylePrimaryButton(Button btn)
        {
            btn.FlatStyle = FlatStyle.Flat;
            btn.BackColor = AccentColor;
            btn.ForeColor = Color.White;
            btn.Font = ButtonFont;
            btn.FlatAppearance.BorderSize = 0;
            btn.FlatAppearance.MouseOverBackColor = AccentDark;
            btn.FlatAppearance.MouseDownBackColor = Color.FromArgb(40, 90, 170);
            btn.Cursor = Cursors.Hand;
        }

        /// <summary>
        /// Style a button as a secondary (cancel) button.
        /// </summary>
        public static void StyleSecondaryButton(Button btn)
        {
            btn.FlatStyle = FlatStyle.Flat;
            btn.BackColor = BtnCancelBg;
            btn.ForeColor = TextPrimary;
            btn.Font = ButtonFont;
            btn.FlatAppearance.BorderSize = 1;
            btn.FlatAppearance.BorderColor = Color.FromArgb(200, 204, 210);
            btn.FlatAppearance.MouseOverBackColor = BtnCancelHover;
            btn.FlatAppearance.MouseDownBackColor = Color.FromArgb(210, 214, 220);
            btn.Cursor = Cursors.Hand;
        }

        /// <summary>
        /// Style a small action button (e.g. "Browse...").
        /// </summary>
        public static void StyleSmallButton(Button btn)
        {
            btn.FlatStyle = FlatStyle.Flat;
            btn.BackColor = BtnCancelBg;
            btn.ForeColor = TextPrimary;
            btn.Font = CreateSafeFont("Microsoft YaHei UI", 8.5f);
            btn.FlatAppearance.BorderSize = 1;
            btn.FlatAppearance.BorderColor = Color.FromArgb(200, 204, 210);
            btn.FlatAppearance.MouseOverBackColor = BtnCancelHover;
            btn.Cursor = Cursors.Hand;
        }

        // ── Input control styling ──────────────────────────────────────

        public static void StyleTextBox(TextBox tb)
        {
            tb.BorderStyle = BorderStyle.FixedSingle;
            tb.Font = BaseFont;
        }

        public static void StyleNumericUpDown(NumericUpDown nud)
        {
            nud.BorderStyle = BorderStyle.FixedSingle;
            nud.Font = BaseFont;
        }

        public static void StyleCheckBox(CheckBox chk)
        {
            chk.Font = BaseFont;
            chk.ForeColor = TextPrimary;
        }

        public static void StyleRadioButton(RadioButton rb)
        {
            rb.Font = BaseFont;
            rb.ForeColor = TextPrimary;
        }

        public static void StyleLabel(Label lbl)
        {
            lbl.Font = LabelFont;
            lbl.ForeColor = TextPrimary;
        }

        // ── GroupBox styling ───────────────────────────────────────────

        /// <summary>
        /// Style a GroupBox with a subtle flat border and accent-colored text.
        /// </summary>
        public static void StyleGroupBox(GroupBox grp)
        {
            grp.Font = CreateSafeFont("Microsoft YaHei UI", 9f, FontStyle.Bold);
            grp.ForeColor = AccentDark;
            grp.Paint += GroupBox_Paint;
        }

        private static void GroupBox_Paint(object sender, PaintEventArgs e)
        {
            var grp = (GroupBox)sender;
            // Draw rounded rectangle border
            var rect = new Rectangle(0, 8, grp.Width - 1, grp.Height - 9);
            using (var pen = new Pen(GroupBorder, 1))
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                e.Graphics.DrawRectangle(pen, rect);
            }

            // Draw group text
            if (!string.IsNullOrEmpty(grp.Text))
            {
                var textSize = e.Graphics.MeasureString(grp.Text, grp.Font);
                e.Graphics.DrawString(grp.Text, grp.Font, new SolidBrush(grp.ForeColor), 8, 0);
            }
        }

        // ── ListBox / CheckedListBox styling ───────────────────────────

        public static void StyleListBox(ListBox lb)
        {
            lb.BorderStyle = BorderStyle.FixedSingle;
            lb.Font = BaseFont;
            lb.ItemHeight = 22;
        }

        public static void StyleCheckedListBox(CheckedListBox clb)
        {
            clb.BorderStyle = BorderStyle.FixedSingle;
            clb.Font = BaseFont;
            clb.CheckOnClick = true;
        }

        // ── Utility: create styled OK/Cancel button pair ───────────────

        /// <summary>
        /// Create and add styled OK and Cancel buttons to a form.
        /// Returns (btnOK, btnCancel).
        /// </summary>
        public static (Button btnOK, Button btnCancel) AddActionButtons(Form form, int y)
        {
            var panel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = form.Height - y - 30, // approximate; caller adjusts
                BackColor = Color.Transparent
            };

            var btnOK = new Button
            {
                Text = "确定",
                DialogResult = DialogResult.OK,
                Width = 88,
                Height = 32
            };
            StylePrimaryButton(btnOK);

            var btnCancel = new Button
            {
                Text = "取消",
                DialogResult = DialogResult.Cancel,
                Width = 88,
                Height = 32
            };
            StyleSecondaryButton(btnCancel);

            return (btnOK, btnCancel);
        }
    }
}
