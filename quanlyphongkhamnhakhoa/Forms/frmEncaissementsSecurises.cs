using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace PKNK_CNPM.Forms
{
    public class FormEncaissementsSecurises : Form
    {
        private readonly TextBox _searchBox;
        private readonly ComboBox _statusFilter;
        private readonly ComboBox _managementFilter;
        private readonly DataGridView _resultsGrid;
        private readonly Label _summaryLabel;
        private readonly Label _statusLabel;
        private DataTable _allResults;

        public FormEncaissementsSecurises()
        {
            Text = "Encaissements plus sécurisés";
            BackColor = Color.FromArgb(245, 248, 250);
            Font = new Font("Segoe UI", 9F);
            MinimumSize = new Size(850, 540);

            Panel header = new Panel
            {
                Dock = DockStyle.Top,
                Height = 145,
                Padding = new Padding(18, 8, 18, 8),
                BackColor = Color.White
            };
            Label title = new Label
            {
                Text = "Encaissements plus sécurisés",
                Dock = DockStyle.Top,
                Height = 34,
                Font = new Font("Segoe UI", 16F, FontStyle.Bold),
                ForeColor = Color.FromArgb(15, 70, 75)
            };
            Label description = new Label
            {
                Text = "Contrôles en lecture seule : quittances répétées, champs obligatoires et montants à vérifier.",
                Dock = DockStyle.Top,
                Height = 27,
                ForeColor = Color.FromArgb(71, 85, 105)
            };
            FlowLayoutPanel filters = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(0, 6, 0, 0),
                BackColor = Color.White
            };

            _searchBox = new TextBox
            {
                Width = 260,
                Height = 29,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.FromArgb(39, 52, 67)
            };
            _searchBox.TextChanged += (sender, e) => ApplyFilters();

            _statusFilter = new ComboBox
            {
                Width = 145,
                DropDownStyle = ComboBoxStyle.DropDownList,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9F)
            };
            _statusFilter.Items.AddRange(new object[] { "Tous les contrôles", "À vérifier", "Contrôlé" });
            _statusFilter.SelectedIndex = 0;
            _statusFilter.SelectedIndexChanged += (sender, e) => ApplyFilters();

            _managementFilter = new ComboBox
            {
                Width = 170,
                DropDownStyle = ComboBoxStyle.DropDownList,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9F)
            };
            _managementFilter.Items.Add("Toutes les gérances");
            _managementFilter.SelectedIndex = 0;
            _managementFilter.SelectedIndexChanged += ManagementFilterChanged;

            Button refreshButton = new Button
            {
                Text = "Actualiser",
                Width = 105,
                Height = 29,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(15, 118, 110),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Cursor = Cursors.Hand,
                UseVisualStyleBackColor = false
            };
            refreshButton.FlatAppearance.BorderSize = 0;
            refreshButton.Click += (sender, e) => LoadData();

            filters.Controls.Add(_searchBox);
            filters.Controls.Add(_statusFilter);
            filters.Controls.Add(_managementFilter);
            filters.Controls.Add(refreshButton);
            header.Controls.Add(filters);
            header.Controls.Add(description);
            header.Controls.Add(title);
            ResponsiveScreenLayout.ConfigureFilterHeader(header, filters, title.Height + description.Height);

            _resultsGrid = new DataGridView
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
            _resultsGrid.DataBindingComplete += ResultsGridDataBindingComplete;

            Panel footer = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 58,
                Padding = new Padding(12, 4, 12, 4),
                BackColor = Color.White
            };
            _summaryLabel = new Label
            {
                Dock = DockStyle.Top,
                Height = 25,
                ForeColor = Color.FromArgb(39, 52, 67),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft
            };
            _statusLabel = new Label
            {
                Dock = DockStyle.Fill,
                ForeColor = Color.FromArgb(71, 85, 105),
                TextAlign = ContentAlignment.MiddleLeft
            };
            footer.Controls.Add(_statusLabel);
            footer.Controls.Add(_summaryLabel);

            Controls.Add(_resultsGrid);
            Controls.Add(footer);
            Controls.Add(header);
            LoadData();
        }

        private void LoadData()
        {
            DataTable source = new DataTable();
            string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + AGI_DB.PathDataBase.Trim() + ";";
            string query = "SELECT [N], [quittance], [code_ger], [code_loc], [Nom_Locataire], " +
                "[date_encs], [total], [def_loy], [intitulé] FROM [Encaissement] ORDER BY [date_encs] DESC";

            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(query, connection))
                {
                    adapter.Fill(source);
                }

                _allResults = BuildControlResults(source);
                PopulateManagementFilter(_allResults);
                ApplyFilters();
            }
            catch (Exception ex)
            {
                _allResults = null;
                _resultsGrid.DataSource = null;
                _summaryLabel.Text = string.Empty;
                _statusLabel.Text = "Échec du chargement des contrôles.";
                MessageBox.Show("Impossible de charger les contrôles d'encaissement : " + ex.Message,
                    "Erreur de chargement", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static DataTable BuildControlResults(DataTable source)
        {
            Dictionary<string, int> receiptCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (DataRow row in source.Rows)
            {
                string receipt = GetValue(row, "quittance");
                string management = GetValue(row, "code_ger");
                if (receipt.Length == 0 || management.Length == 0)
                {
                    continue;
                }

                string key = management.Trim() + "|" + receipt.Trim();
                int count;
                receiptCounts.TryGetValue(key, out count);
                receiptCounts[key] = count + 1;
            }

            DataTable results = new DataTable();
            results.Columns.Add("N°", typeof(string));
            results.Columns.Add("Quittance", typeof(string));
            results.Columns.Add("Code gérance", typeof(string));
            results.Columns.Add("Code locataire", typeof(string));
            results.Columns.Add("Nom locataire", typeof(string));
            results.Columns.Add("Date", typeof(string));
            results.Columns.Add("Montant", typeof(string));
            results.Columns.Add("Déficit", typeof(string));
            results.Columns.Add("Période", typeof(string));
            results.Columns.Add("Contrôle", typeof(string));
            results.Columns.Add("Anomalies à vérifier", typeof(string));

            foreach (DataRow row in source.Rows)
            {
                string receipt = GetValue(row, "quittance");
                string management = GetValue(row, "code_ger");
                string tenantCode = GetValue(row, "code_loc");
                List<string> issues = new List<string>();
                if (receipt.Length == 0)
                {
                    issues.Add("Quittance absente");
                }
                else if (management.Length > 0)
                {
                    int duplicateCount;
                    if (receiptCounts.TryGetValue(management.Trim() + "|" + receipt.Trim(), out duplicateCount) && duplicateCount > 1)
                    {
                        issues.Add("Quittance répétée pour cette gérance");
                    }
                }

                if (management.Length == 0)
                {
                    issues.Add("Code gérance absent");
                }
                if (tenantCode.Length == 0)
                {
                    issues.Add("Code locataire absent");
                }
                if (row["date_encs"] == DBNull.Value || GetValue(row, "date_encs").Length == 0)
                {
                    issues.Add("Date absente");
                }
                if (row["total"] == DBNull.Value || !TryGetAmount(row["total"], out decimal amount) || amount <= 0)
                {
                    issues.Add("Montant nul, négatif ou non numérique");
                }

                results.Rows.Add(
                    GetValue(row, "N"), receipt, management, tenantCode, GetValue(row, "Nom_Locataire"),
                    FormatDate(row["date_encs"]), GetValue(row, "total"), GetValue(row, "def_loy"),
                    GetValue(row, "intitulé"), issues.Count == 0 ? "Contrôlé" : "À vérifier",
                    string.Join(" ; ", issues.ToArray()));
            }
            return results;
        }

        private void PopulateManagementFilter(DataTable results)
        {
            string previousSelection = _managementFilter.SelectedItem as string;
            _managementFilter.SelectedIndexChanged -= ManagementFilterChanged;
            _managementFilter.Items.Clear();
            _managementFilter.Items.Add("Toutes les gérances");
            foreach (DataRow row in results.Rows)
            {
                string code = Convert.ToString(row["Code gérance"], CultureInfo.CurrentCulture).Trim();
                if (code.Length > 0 && !_managementFilter.Items.Contains(code))
                {
                    _managementFilter.Items.Add(code);
                }
            }

            if (!string.IsNullOrEmpty(previousSelection) && _managementFilter.Items.Contains(previousSelection))
            {
                _managementFilter.SelectedItem = previousSelection;
            }
            else
            {
                _managementFilter.SelectedIndex = 0;
            }
            _managementFilter.SelectedIndexChanged += ManagementFilterChanged;
        }

        private void ManagementFilterChanged(object sender, EventArgs e)
        {
            ApplyFilters();
        }

        private void ApplyFilters()
        {
            if (_allResults == null)
            {
                _resultsGrid.DataSource = null;
                return;
            }

            string search = _searchBox.Text.Trim();
            string selectedStatus = _statusFilter.SelectedItem as string;
            string selectedManagement = _managementFilter.SelectedItem as string;
            DataTable filtered = _allResults.Clone();

            foreach (DataRow row in _allResults.Rows)
            {
                if (selectedStatus != "Tous les contrôles" &&
                    !string.Equals(Convert.ToString(row["Contrôle"]), selectedStatus, StringComparison.CurrentCultureIgnoreCase))
                {
                    continue;
                }
                if (selectedManagement != "Toutes les gérances" &&
                    !string.Equals(Convert.ToString(row["Code gérance"]), selectedManagement, StringComparison.CurrentCultureIgnoreCase))
                {
                    continue;
                }

                bool matches = search.Length == 0;
                if (!matches)
                {
                    foreach (object value in row.ItemArray)
                    {
                        if (value != DBNull.Value && Convert.ToString(value, CultureInfo.CurrentCulture)
                            .IndexOf(search, StringComparison.CurrentCultureIgnoreCase) >= 0)
                        {
                            matches = true;
                            break;
                        }
                    }
                }
                if (matches)
                {
                    filtered.ImportRow(row);
                }
            }

            _resultsGrid.DataSource = filtered;
            int flagged = 0;
            foreach (DataRow row in filtered.Rows)
            {
                if (Convert.ToString(row["Contrôle"]) == "À vérifier")
                {
                    flagged++;
                }
            }
            _summaryLabel.Text = filtered.Rows.Count.ToString("N0", CultureInfo.CurrentCulture) +
                " encaissement(s) affiché(s), dont " + flagged.ToString("N0", CultureInfo.CurrentCulture) + " à vérifier.";
            _statusLabel.Text = "Aucune donnée n'est modifiée. Les alertes signalent des vérifications à effectuer, pas une preuve d'erreur.";
        }

        private void ResultsGridDataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            if (_resultsGrid.Columns.Contains("N°"))
            {
                _resultsGrid.Columns["N°"].FillWeight = 45;
            }
            if (_resultsGrid.Columns.Contains("Anomalies à vérifier"))
            {
                _resultsGrid.Columns["Anomalies à vérifier"].FillWeight = 180;
            }
            if (_resultsGrid.Columns.Contains("Contrôle"))
            {
                _resultsGrid.Columns["Contrôle"].FillWeight = 70;
            }

            foreach (DataGridViewRow row in _resultsGrid.Rows)
            {
                if (Convert.ToString(row.Cells["Contrôle"].Value) == "À vérifier")
                {
                    row.Cells["Contrôle"].Style.ForeColor = Color.FromArgb(154, 52, 18);
                    row.Cells["Contrôle"].Style.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
                }
            }
        }

        private static string GetValue(DataRow row, string column)
        {
            return row[column] == DBNull.Value
                ? string.Empty
                : Convert.ToString(row[column], CultureInfo.CurrentCulture).Trim();
        }

        private static bool TryGetAmount(object value, out decimal amount)
        {
            string text = Convert.ToString(value, CultureInfo.CurrentCulture);
            return decimal.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out amount) ||
                decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out amount);
        }

        private static string FormatDate(object value)
        {
            if (value == null || value == DBNull.Value)
            {
                return string.Empty;
            }

            DateTime date;
            if (DateTime.TryParse(Convert.ToString(value, CultureInfo.CurrentCulture), out date))
            {
                return date.ToString("dd/MM/yyyy", CultureInfo.CurrentCulture);
            }
            return Convert.ToString(value, CultureInfo.CurrentCulture);
        }
    }
}
