using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace PKNK_CNPM.Forms
{
    public class FormSuiviImpayes : Form
    {
        private readonly DataGridView _grid;
        private readonly TextBox _searchBox;
        private readonly ComboBox _managementFilter;
        private readonly Label _countLabel;
        private readonly Label _updatedLabel;
        private DataTable _arrears;

        public FormSuiviImpayes()
        {
            Text = "Suivi des impayés";
            BackColor = Color.FromArgb(245, 248, 250);
            Font = new Font("Segoe UI", 9F);
            MinimumSize = new Size(760, 500);

            Panel header = new Panel
            {
                Dock = DockStyle.Top,
                Height = 126,
                Padding = new Padding(20, 12, 20, 8),
                BackColor = Color.White
            };

            Label title = new Label
            {
                Text = "Suivi des impayés",
                Dock = DockStyle.Top,
                Height = 34,
                Font = new Font("Segoe UI", 16F, FontStyle.Bold),
                ForeColor = Color.FromArgb(15, 70, 75)
            };

            Label description = new Label
            {
                Text = "Les valeurs sont reprises du champ « impy » de la fiche locataire. Cet écran ne modifie aucun encaissement.",
                Dock = DockStyle.Top,
                Height = 28,
                ForeColor = Color.FromArgb(71, 85, 105)
            };

            FlowLayoutPanel filters = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(0, 7, 0, 0),
                BackColor = Color.White
            };

            _searchBox = new TextBox
            {
                Width = 330,
                Height = 30,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.FromArgb(39, 52, 67)
            };
            _searchBox.TextChanged += (sender, e) => ApplyFilters();

            _managementFilter = new ComboBox
            {
                Width = 220,
                DropDownStyle = ComboBoxStyle.DropDownList,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9F)
            };
            _managementFilter.Items.Add("Toutes les gérances");
            _managementFilter.SelectedIndex = 0;
            _managementFilter.SelectedIndexChanged += ManagementFilterChanged;

            Button refreshButton = CreateButton("Actualiser", Color.FromArgb(15, 118, 110));
            refreshButton.Click += (sender, e) => LoadData();
            Button exportButton = CreateButton("Exporter CSV", Color.FromArgb(51, 65, 85));
            exportButton.Click += (sender, e) => ExportCsv();

            filters.Controls.Add(_searchBox);
            filters.Controls.Add(_managementFilter);
            filters.Controls.Add(refreshButton);
            filters.Controls.Add(exportButton);
            header.Controls.Add(filters);
            header.Controls.Add(description);
            header.Controls.Add(title);
            ResponsiveScreenLayout.ConfigureFilterHeader(header, filters, title.Height + description.Height);

            Panel footer = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 36,
                Padding = new Padding(14, 8, 14, 4),
                BackColor = Color.White
            };
            _updatedLabel = new Label
            {
                Dock = DockStyle.Right,
                AutoSize = true,
                ForeColor = Color.FromArgb(100, 116, 139),
                TextAlign = ContentAlignment.MiddleRight
            };
            _countLabel = new Label
            {
                Dock = DockStyle.Left,
                AutoSize = true,
                ForeColor = Color.FromArgb(39, 52, 67),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };
            footer.Controls.Add(_updatedLabel);
            footer.Controls.Add(_countLabel);

            _grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                AutoGenerateColumns = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                RowHeadersVisible = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                GridColor = Color.FromArgb(226, 232, 240),
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 38,
                RowTemplate = { Height = 32 },
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(15, 118, 110),
                    ForeColor = Color.White,
                    Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                    SelectionBackColor = Color.FromArgb(15, 118, 110)
                },
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.White,
                    ForeColor = Color.FromArgb(39, 52, 67),
                    SelectionBackColor = Color.FromArgb(204, 235, 231),
                    SelectionForeColor = Color.FromArgb(39, 52, 67)
                },
                AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(245, 248, 250)
                }
            };

            Controls.Add(_grid);
            Controls.Add(footer);
            Controls.Add(header);

            LoadData();
        }

        private static Button CreateButton(string text, Color color)
        {
            Button button = new Button
            {
                Text = text,
                Width = 112,
                Height = 30,
                FlatStyle = FlatStyle.Flat,
                BackColor = color,
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Cursor = Cursors.Hand,
                UseVisualStyleBackColor = false
            };
            button.FlatAppearance.BorderSize = 0;
            return button;
        }

        private void LoadData()
        {
            DataTable tenants = new DataTable();
            string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + AGI_DB.PathDataBase.Trim() + ";";
            string query = "SELECT [code loc] AS [Code locataire], [CODE GER] AS [Code gérance], " +
                "[nom loc] AS [Nom locataire], [adresse] AS [Adresse], [total] AS [Total enregistré], " +
                "[impy] AS [Impayé enregistré (impy)] FROM [loc] " +
                "WHERE [code loc] IS NOT NULL ORDER BY [nom loc]";

            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(query, connection))
                {
                    adapter.Fill(tenants);
                }

                _arrears = CreateArrearsTable(tenants);
                PopulateManagementFilter(_arrears);
                ApplyFilters();
                _updatedLabel.Text = "Actualisé le " + DateTime.Now.ToString("dd/MM/yyyy à HH:mm", CultureInfo.CurrentCulture);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Impossible de charger le suivi des impayés : " + ex.Message,
                    "Erreur de chargement", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static DataTable CreateArrearsTable(DataTable tenants)
        {
            DataTable arrears = tenants.Clone();
            foreach (DataRow tenant in tenants.Rows)
            {
                object value = tenant["Impayé enregistré (impy)"];
                if (HasOutstandingValue(value))
                {
                    arrears.ImportRow(tenant);
                }
            }
            return arrears;
        }

        private static bool HasOutstandingValue(object value)
        {
            if (value == null || value == DBNull.Value)
            {
                return false;
            }

            string text = Convert.ToString(value, CultureInfo.CurrentCulture).Trim();
            if (text.Length == 0)
            {
                return false;
            }

            decimal amount;
            if (decimal.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out amount) ||
                decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out amount))
            {
                return amount != 0;
            }

            foreach (char character in text)
            {
                if (char.IsLetter(character) || (char.IsDigit(character) && character != '0'))
                {
                    return true;
                }
            }

            return false;
        }

        private void PopulateManagementFilter(DataTable arrears)
        {
            object selected = _managementFilter.SelectedItem;
            _managementFilter.SelectedIndexChanged -= ManagementFilterChanged;
            _managementFilter.Items.Clear();
            _managementFilter.Items.Add("Toutes les gérances");
            foreach (DataRow row in arrears.Rows)
            {
                string management = row["Code gérance"] == DBNull.Value
                    ? string.Empty
                    : Convert.ToString(row["Code gérance"], CultureInfo.CurrentCulture).Trim();
                if (management.Length > 0 && !_managementFilter.Items.Contains(management))
                {
                    _managementFilter.Items.Add(management);
                }
            }
            _managementFilter.SelectedIndex = 0;
            _managementFilter.SelectedIndexChanged += ManagementFilterChanged;
        }

        private void ManagementFilterChanged(object sender, EventArgs e)
        {
            ApplyFilters();
        }

        private void ApplyFilters()
        {
            if (_arrears == null)
            {
                _grid.DataSource = null;
                _countLabel.Text = "Chargement des dossiers...";
                return;
            }

            string search = _searchBox.Text.Trim();
            string selectedManagement = _managementFilter.SelectedItem as string;
            DataTable filtered = _arrears.Clone();

            foreach (DataRow row in _arrears.Rows)
            {
                string management = row["Code gérance"] == DBNull.Value
                    ? string.Empty
                    : Convert.ToString(row["Code gérance"], CultureInfo.CurrentCulture).Trim();
                if (!string.IsNullOrEmpty(selectedManagement) && selectedManagement != "Toutes les gérances" &&
                    !string.Equals(management, selectedManagement, StringComparison.CurrentCultureIgnoreCase))
                {
                    continue;
                }

                bool matchesSearch = search.Length == 0;
                if (!matchesSearch)
                {
                    foreach (object value in row.ItemArray)
                    {
                        if (value != DBNull.Value && Convert.ToString(value, CultureInfo.CurrentCulture)
                            .IndexOf(search, StringComparison.CurrentCultureIgnoreCase) >= 0)
                        {
                            matchesSearch = true;
                            break;
                        }
                    }
                }

                if (matchesSearch)
                {
                    filtered.ImportRow(row);
                }
            }

            _grid.DataSource = filtered;
            _countLabel.Text = filtered.Rows.Count.ToString("N0", CultureInfo.CurrentCulture) +
                (filtered.Rows.Count == 1 ? " dossier avec impayé" : " dossiers avec impayé");
            if (_grid.Columns.Contains("Impayé enregistré (impy)"))
            {
                _grid.Columns["Impayé enregistré (impy)"].DefaultCellStyle.BackColor = Color.FromArgb(255, 247, 237);
                _grid.Columns["Impayé enregistré (impy)"].DefaultCellStyle.ForeColor = Color.FromArgb(154, 52, 18);
            }
        }

        private void ExportCsv()
        {
            if (_grid.DataSource == null || _grid.Rows.Count == 0)
            {
                MessageBox.Show("Aucun dossier à exporter.", "Export CSV", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (SaveFileDialog dialog = new SaveFileDialog())
            {
                dialog.Filter = "Fichier CSV (*.csv)|*.csv";
                dialog.DefaultExt = "csv";
                dialog.AddExtension = true;
                dialog.FileName = "Suivi_Impayes_" + DateTime.Today.ToString("yyyyMMdd", CultureInfo.InvariantCulture) + ".csv";
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                try
                {
                    using (StreamWriter writer = new StreamWriter(dialog.FileName, false, new UTF8Encoding(true)))
                    {
                        string[] headers = new string[_grid.Columns.Count];
                        for (int columnIndex = 0; columnIndex < _grid.Columns.Count; columnIndex++)
                        {
                            headers[columnIndex] = EscapeCsv(_grid.Columns[columnIndex].HeaderText);
                        }
                        writer.WriteLine(string.Join(";", headers));

                        foreach (DataGridViewRow row in _grid.Rows)
                        {
                            string[] values = new string[_grid.Columns.Count];
                            for (int columnIndex = 0; columnIndex < _grid.Columns.Count; columnIndex++)
                            {
                                object value = row.Cells[columnIndex].Value;
                                values[columnIndex] = EscapeCsv(value == null || value == DBNull.Value
                                    ? string.Empty
                                    : Convert.ToString(value, CultureInfo.CurrentCulture));
                            }
                            writer.WriteLine(string.Join(";", values));
                        }
                    }

                    MessageBox.Show("L'export a été créé : " + dialog.FileName,
                        "Export terminé", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("L'export CSV a échoué : " + ex.Message,
                        "Erreur d'export", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private static string EscapeCsv(string value)
        {
            return "\"" + (value ?? string.Empty).Replace("\"", "\"\"") + "\"";
        }
    }
}
