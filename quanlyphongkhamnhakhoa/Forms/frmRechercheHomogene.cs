using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace PKNK_CNPM.Forms
{
    public class FormRechercheHomogene : Form
    {
        private readonly TextBox _searchBox;
        private readonly ComboBox _typeFilter;
        private readonly ComboBox _managementFilter;
        private readonly DataGridView _resultsGrid;
        private readonly Label _statusLabel;
        private DataTable _allResults;

        public FormRechercheHomogene()
        {
            Text = "Recherche et filtres";
            BackColor = Color.FromArgb(245, 248, 250);
            Font = new Font("Segoe UI", 9F);
            MinimumSize = new Size(800, 540);

            Panel header = new Panel
            {
                Dock = DockStyle.Top,
                Height = 122,
                Padding = new Padding(18, 8, 18, 8),
                BackColor = Color.White
            };
            Label title = new Label
            {
                Text = "Recherche et filtres",
                Dock = DockStyle.Top,
                Height = 34,
                Font = new Font("Segoe UI", 16F, FontStyle.Bold),
                ForeColor = Color.FromArgb(15, 70, 75)
            };
            Label description = new Label
            {
                Text = "Recherchez dans les locataires, encaissements, logements et gérances. Les résultats sont en lecture seule.",
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
                Width = 340,
                Height = 29,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.FromArgb(39, 52, 67)
            };
            _searchBox.TextChanged += (sender, e) => ApplyFilters();

            _typeFilter = new ComboBox
            {
                Width = 190,
                DropDownStyle = ComboBoxStyle.DropDownList,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9F)
            };
            _typeFilter.Items.AddRange(new object[]
            {
                "Tous les types", "Locataires", "Encaissements", "Logements", "Gérances"
            });
            _typeFilter.SelectedIndex = 0;
            _typeFilter.SelectedIndexChanged += (sender, e) => ApplyFilters();

            _managementFilter = new ComboBox
            {
                Width = 190,
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
                Width = 110,
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
            filters.Controls.Add(_typeFilter);
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

            Panel footer = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 34,
                Padding = new Padding(12, 7, 12, 4),
                BackColor = Color.White
            };
            _statusLabel = new Label
            {
                Dock = DockStyle.Fill,
                ForeColor = Color.FromArgb(71, 85, 105),
                TextAlign = ContentAlignment.MiddleLeft
            };
            footer.Controls.Add(_statusLabel);

            Controls.Add(_resultsGrid);
            Controls.Add(footer);
            Controls.Add(header);
            LoadData();
        }

        private void LoadData()
        {
            DataTable results = CreateResultsTable();
            string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + AGI_DB.PathDataBase.Trim() + ";";

            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    LoadSource(connection, results, "Locataires",
                        "SELECT [code loc], [CODE GER], [nom loc], [adresse], [datloc], [total], [impy] " +
                        "FROM [loc] WHERE [code loc] IS NOT NULL");
                    LoadSource(connection, results, "Encaissements",
                        "SELECT [code_loc], [code_ger], [Nom_Locataire], [quittance], [date_encs], [total], [intitulé] " +
                        "FROM [Encaissement]");
                    LoadSource(connection, results, "Logements",
                        "SELECT [Id_Appartement], [Type_Appartement], [Etat], [Arrondissement], [Quartier], [Prix] " +
                        "FROM [Appartement]");
                    LoadSource(connection, results, "Gérances",
                        "SELECT [CODE GER], [nom  ger], [adresse ger], [tel ger] FROM [ger]");
                }

                _allResults = results;
                PopulateManagementFilter(results);
                ApplyFilters();
            }
            catch (Exception ex)
            {
                _statusLabel.Text = "Échec de chargement des résultats.";
                MessageBox.Show("Impossible de charger la recherche globale : " + ex.Message,
                    "Erreur de chargement", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static DataTable CreateResultsTable()
        {
            DataTable results = new DataTable();
            results.Columns.Add("Type", typeof(string));
            results.Columns.Add("Code", typeof(string));
            results.Columns.Add("Code gérance", typeof(string));
            results.Columns.Add("Nom / Référence", typeof(string));
            results.Columns.Add("Adresse / Détails", typeof(string));
            results.Columns.Add("Date", typeof(string));
            results.Columns.Add("Montant / Valeur", typeof(string));
            results.Columns.Add("Informations", typeof(string));
            return results;
        }

        private static void LoadSource(OleDbConnection connection, DataTable results, string category, string query)
        {
            using (OleDbDataAdapter adapter = new OleDbDataAdapter(query, connection))
            {
                DataTable source = new DataTable();
                adapter.Fill(source);
                foreach (DataRow row in source.Rows)
                {
                    switch (category)
                    {
                        case "Locataires":
                            results.Rows.Add(category,
                                GetValue(row, "code loc"), GetValue(row, "CODE GER"), GetValue(row, "nom loc"),
                                GetValue(row, "adresse"), FormatDate(row["datloc"]), GetValue(row, "total"),
                                "Impayé enregistré : " + GetValue(row, "impy"));
                            break;
                        case "Encaissements":
                            results.Rows.Add(category,
                                GetValue(row, "code_loc"), GetValue(row, "code_ger"), GetValue(row, "Nom_Locataire"),
                                "Quittance : " + GetValue(row, "quittance"), FormatDate(row["date_encs"]),
                                GetValue(row, "total"), "Période : " + GetValue(row, "intitulé"));
                            break;
                        case "Logements":
                            results.Rows.Add(category,
                                GetValue(row, "Id_Appartement"), string.Empty,
                                GetValue(row, "Type_Appartement"),
                                GetValue(row, "Arrondissement") + " " + GetValue(row, "Quartier"),
                                string.Empty, GetValue(row, "Prix"), "État : " + GetValue(row, "Etat"));
                            break;
                        case "Gérances":
                            results.Rows.Add(category,
                                GetValue(row, "CODE GER"), GetValue(row, "CODE GER"), GetValue(row, "nom  ger"),
                                GetValue(row, "adresse ger"), string.Empty, string.Empty,
                                "Téléphone : " + GetValue(row, "tel ger"));
                            break;
                    }
                }
            }
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
                _statusLabel.Text = "Chargement des données...";
                return;
            }

            string search = _searchBox.Text.Trim();
            string selectedType = _typeFilter.SelectedItem as string;
            string selectedManagement = _managementFilter.SelectedItem as string;
            DataTable filtered = CreateResultsTable();

            foreach (DataRow row in _allResults.Rows)
            {
                if (!string.IsNullOrEmpty(selectedType) && selectedType != "Tous les types" &&
                    !string.Equals(Convert.ToString(row["Type"]), selectedType, StringComparison.CurrentCultureIgnoreCase))
                {
                    continue;
                }

                if (!string.IsNullOrEmpty(selectedManagement) && selectedManagement != "Toutes les gérances" &&
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
            _statusLabel.Text = filtered.Rows.Count.ToString("N0", CultureInfo.CurrentCulture) +
                " résultat(s) sur " + _allResults.Rows.Count.ToString("N0", CultureInfo.CurrentCulture) +
                " enregistrement(s).";
        }

        private static string GetValue(DataRow row, string column)
        {
            return row[column] == DBNull.Value
                ? string.Empty
                : Convert.ToString(row[column], CultureInfo.CurrentCulture).Trim();
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
