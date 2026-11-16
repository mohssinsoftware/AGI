using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

namespace PKNK_CNPM.Forms
{
    public class FormAffichagePratique : Form
    {
        private readonly ComboBox _densitySelector;
        private readonly ComboBox _textSizeSelector;
        private readonly Label _statusLabel;

        public FormAffichagePratique()
        {
            Text = "Affichage plus pratique";
            BackColor = Color.FromArgb(245, 248, 250);
            Font = new Font("Segoe UI", 9F);
            MinimumSize = new Size(700, 460);

            Panel header = new Panel
            {
                Dock = DockStyle.Top,
                Height = 96,
                Padding = new Padding(18, 8, 18, 8),
                BackColor = Color.White
            };
            Label title = new Label
            {
                Text = "Affichage plus pratique",
                Dock = DockStyle.Top,
                Height = 34,
                Font = new Font("Segoe UI", 16F, FontStyle.Bold),
                ForeColor = Color.FromArgb(15, 70, 75)
            };
            Label description = new Label
            {
                Text = "Adaptez la lisibilité des tableaux sans modifier la taille ni la disposition des formulaires.",
                Dock = DockStyle.Fill,
                ForeColor = Color.FromArgb(71, 85, 105)
            };
            header.Controls.Add(description);
            header.Controls.Add(title);

            TableLayoutPanel options = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                ColumnCount = 2,
                Padding = new Padding(24),
                BackColor = Color.White
            };
            options.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 210));
            options.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            Label densityLabel = CreateOptionLabel("Espacement des lignes");
            _densitySelector = CreateSelector("Standard", "Compact", "Confortable");
            Label textSizeLabel = CreateOptionLabel("Texte des tableaux");
            _textSizeSelector = CreateSelector("Normal", "Grand");
            options.Controls.Add(densityLabel, 0, 0);
            options.Controls.Add(_densitySelector, 1, 0);
            options.Controls.Add(textSizeLabel, 0, 1);
            options.Controls.Add(_textSizeSelector, 1, 1);

            Label hint = new Label
            {
                Text = "Les préférences sont conservées pour les prochaines ouvertures. Elles s'appliquent aux listes et tableaux de l'application.",
                Dock = DockStyle.Top,
                Height = 54,
                Padding = new Padding(24, 12, 24, 0),
                ForeColor = Color.FromArgb(71, 85, 105)
            };

            Panel footer = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 66,
                Padding = new Padding(18, 10, 18, 8),
                BackColor = Color.White
            };
            Button applyButton = new Button
            {
                Text = "Appliquer",
                Dock = DockStyle.Right,
                Width = 130,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(15, 118, 110),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Cursor = Cursors.Hand,
                UseVisualStyleBackColor = false
            };
            applyButton.FlatAppearance.BorderSize = 0;
            applyButton.Click += (sender, e) => SaveAndApply();
            _statusLabel = new Label
            {
                Dock = DockStyle.Fill,
                ForeColor = Color.FromArgb(71, 85, 105),
                TextAlign = ContentAlignment.MiddleLeft
            };
            footer.Controls.Add(applyButton);
            footer.Controls.Add(_statusLabel);

            Panel content = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(18),
                AutoScroll = true,
                BackColor = Color.FromArgb(245, 248, 250)
            };
            content.Controls.Add(hint);
            content.Controls.Add(options);

            Controls.Add(content);
            Controls.Add(footer);
            Controls.Add(header);
            LoadPreferences();
        }

        private static Label CreateOptionLabel(string text)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                ForeColor = Color.FromArgb(39, 52, 67),
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft
            };
        }

        private static ComboBox CreateSelector(params string[] values)
        {
            ComboBox selector = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10F),
                Margin = new Padding(4, 8, 4, 8)
            };
            selector.Items.AddRange(values);
            return selector;
        }

        private void LoadPreferences()
        {
            string error;
            if (!DisplayPreferences.Load(out error))
            {
                MessageBox.Show("Impossible de lire les préférences d'affichage : " + error,
                    "Préférences", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            _densitySelector.SelectedItem = DisplayPreferences.DensityName;
            _textSizeSelector.SelectedItem = DisplayPreferences.LargeText ? "Grand" : "Normal";
            if (_densitySelector.SelectedIndex < 0)
            {
                _densitySelector.SelectedItem = "Standard";
            }
            if (_textSizeSelector.SelectedIndex < 0)
            {
                _textSizeSelector.SelectedItem = "Normal";
            }
        }

        private void SaveAndApply()
        {
            DisplayPreferences.DensityName = Convert.ToString(_densitySelector.SelectedItem, CultureInfo.CurrentCulture);
            DisplayPreferences.LargeText = Convert.ToString(_textSizeSelector.SelectedItem, CultureInfo.CurrentCulture) == "Grand";
            string error;
            if (!DisplayPreferences.Save(out error))
            {
                _statusLabel.Text = "Préférences non enregistrées : " + error;
                MessageBox.Show("Les préférences n'ont pas pu être enregistrées : " + error,
                    "Erreur d'enregistrement", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            foreach (Form form in Application.OpenForms)
            {
                DisplayPreferences.ApplyTo(form);
            }
            _statusLabel.Text = "Affichage appliqué et enregistré pour les prochaines ouvertures.";
        }
    }

    internal static class ResponsiveScreenLayout
    {
        internal static void ConfigureFilterHeader(Panel header, FlowLayoutPanel filters, int fixedContentHeight)
        {
            filters.WrapContents = true;
            EventHandler updateHeight = (sender, e) =>
            {
                int availableWidth = Math.Max(1, header.ClientSize.Width - header.Padding.Horizontal);
                Size preferredSize = filters.GetPreferredSize(new Size(availableWidth, 0));
                int desiredHeight = header.Padding.Vertical + fixedContentHeight + Math.Max(36, preferredSize.Height);
                if (header.Height != desiredHeight)
                {
                    header.Height = desiredHeight;
                }
            };

            header.SizeChanged += updateHeight;
            updateHeight(header, EventArgs.Empty);
        }

        internal static void ConfigureActionFooter(Panel footer, FlowLayoutPanel actions, int preferredActionsWidth, int minimumHeight)
        {
            actions.WrapContents = true;
            EventHandler updateLayout = (sender, e) =>
            {
                int availableWidth = Math.Max(1, footer.ClientSize.Width - footer.Padding.Horizontal);
                int actionsWidth = Math.Min(preferredActionsWidth, Math.Max(1, availableWidth - 160));
                if (actions.Width != actionsWidth)
                {
                    actions.Width = actionsWidth;
                }

                Size preferredSize = actions.GetPreferredSize(new Size(actionsWidth, 0));
                int desiredHeight = Math.Max(minimumHeight, footer.Padding.Vertical + preferredSize.Height);
                if (footer.Height != desiredHeight)
                {
                    footer.Height = desiredHeight;
                }
            };

            footer.SizeChanged += updateLayout;
            updateLayout(footer, EventArgs.Empty);
        }
    }

    internal static class DisplayPreferences
    {
        private static readonly string SettingsPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "AGI", "affichage.txt");
        private static bool _loaded;

        internal static string DensityName { get; set; } = "Standard";
        internal static bool LargeText { get; set; }

        internal static bool Load(out string error)
        {
            error = string.Empty;
            if (_loaded)
            {
                return true;
            }

            try
            {
                if (File.Exists(SettingsPath))
                {
                    string[] lines = File.ReadAllLines(SettingsPath);
                    if (lines.Length > 0 && (lines[0] == "Compact" || lines[0] == "Standard" || lines[0] == "Confortable"))
                    {
                        DensityName = lines[0];
                    }
                    if (lines.Length > 1)
                    {
                        LargeText = string.Equals(lines[1], "Grand", StringComparison.OrdinalIgnoreCase);
                    }
                }
                _loaded = true;
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        internal static bool Save(out string error)
        {
            error = string.Empty;
            try
            {
                string directory = Path.GetDirectoryName(SettingsPath);
                if (string.IsNullOrEmpty(directory))
                {
                    throw new IOException("Le dossier de préférences est invalide.");
                }
                Directory.CreateDirectory(directory);
                File.WriteAllLines(SettingsPath, new[] { DensityName, LargeText ? "Grand" : "Normal" });
                _loaded = true;
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        internal static void ApplyTo(Control root)
        {
            DataGridView grid = root as DataGridView;
            if (grid != null)
            {
                ApplyToGrid(grid);
            }
            foreach (Control child in root.Controls)
            {
                ApplyTo(child);
            }
        }

        private static void ApplyToGrid(DataGridView grid)
        {
            int rowHeight;
            int headerHeight;
            switch (DensityName)
            {
                case "Compact":
                    rowHeight = 26;
                    headerHeight = 34;
                    break;
                case "Confortable":
                    rowHeight = 40;
                    headerHeight = 44;
                    break;
                default:
                    rowHeight = 32;
                    headerHeight = 38;
                    break;
            }

            float fontSize = LargeText ? 10F : 9F;
            grid.RowTemplate.Height = rowHeight;
            grid.ColumnHeadersHeight = headerHeight;
            grid.DefaultCellStyle.Font = new Font("Segoe UI", fontSize, FontStyle.Regular);
            grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", fontSize, FontStyle.Bold);
            foreach (DataGridViewRow row in grid.Rows)
            {
                if (!row.IsNewRow)
                {
                    row.Height = rowHeight;
                }
            }
        }
    }
}
