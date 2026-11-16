using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace PKNK_CNPM.Forms
{
    public class FormFicheLogement : Form
    {
        private readonly TextBox _searchBox;
        private readonly DataGridView _propertiesGrid;
        private readonly TableLayoutPanel _propertyDetails;
        private readonly TableLayoutPanel _managementDetails;
        private readonly DataGridView _paymentsGrid;
        private readonly DataGridView _ownersGrid;
        private readonly Label _statusLabel;
        private DataTable _properties;

        public FormFicheLogement()
        {
            Text = "Fiche unique par logement";
            BackColor = Color.FromArgb(245, 248, 250);
            Font = new Font("Segoe UI", 9F);
            MinimumSize = new Size(850, 560);

            Panel header = new Panel
            {
                Dock = DockStyle.Top,
                Height = 88,
                Padding = new Padding(18, 8, 18, 8),
                BackColor = Color.White
            };
            Label title = new Label
            {
                Text = "Fiche unique par logement",
                Dock = DockStyle.Top,
                Height = 34,
                Font = new Font("Segoe UI", 16F, FontStyle.Bold),
                ForeColor = Color.FromArgb(15, 70, 75)
            };
            Label description = new Label
            {
                Text = "Sélectionnez un dossier locataire pour consulter ses informations, ses encaissements et sa gérance.",
                Dock = DockStyle.Fill,
                ForeColor = Color.FromArgb(71, 85, 105)
            };
            header.Controls.Add(description);
            header.Controls.Add(title);

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
                ForeColor = Color.FromArgb(100, 116, 139),
                TextAlign = ContentAlignment.MiddleLeft
            };
            footer.Controls.Add(_statusLabel);

            SplitContainer mainSplit = new SplitContainer
            {
                Size = new Size(500, 650),
                Dock = DockStyle.Fill,
                SplitterDistance = 170,
                BackColor = Color.FromArgb(245, 248, 250),
                Panel1MinSize = 150,
                Panel2MinSize = 220
            };

            Panel listPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10),
                BackColor = Color.White
            };
            Label listTitle = new Label
            {
                Text = "Logements / locataires",
                Dock = DockStyle.Top,
                Height = 30,
                ForeColor = Color.FromArgb(15, 118, 110),
                Font = new Font("Segoe UI", 10F, FontStyle.Bold)
            };
            _searchBox = new TextBox
            {
                Dock = DockStyle.Top,
                Height = 28,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.FromArgb(39, 52, 67)
            };
            _searchBox.TextChanged += (sender, e) => ApplySearch();
            _propertiesGrid = CreateGrid();
            _propertiesGrid.SelectionChanged += PropertySelectionChanged;
            listPanel.Controls.Add(_propertiesGrid);
            listPanel.Controls.Add(_searchBox);
            listPanel.Controls.Add(listTitle);
            mainSplit.Panel1.Controls.Add(listPanel);

            TabControl tabs = new TabControl
            {
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Padding = new Point(12, 5)
            };
            TabPage propertyPage = CreatePage("Logement / locataire");
            TabPage paymentsPage = CreatePage("Encaissements");
            TabPage managementPage = CreatePage("Gérance");
            TabPage ownersPage = CreatePage("Copropriétaires");
            tabs.TabPages.AddRange(new[] { propertyPage, paymentsPage, managementPage, ownersPage });

            _propertyDetails = CreateDetailsLayout();
            Panel propertyScroll = new Panel { Dock = DockStyle.Fill, AutoScroll = true, BackColor = Color.White };
            propertyScroll.Controls.Add(_propertyDetails);
            propertyPage.Controls.Add(propertyScroll);

            _paymentsGrid = CreateGrid();
            paymentsPage.Controls.Add(_paymentsGrid);
            _managementDetails = CreateDetailsLayout();
            Panel managementScroll = new Panel { Dock = DockStyle.Fill, AutoScroll = true, BackColor = Color.White };
            managementScroll.Controls.Add(_managementDetails);
            managementPage.Controls.Add(managementScroll);
            _ownersGrid = CreateGrid();
            ownersPage.Controls.Add(_ownersGrid);
            mainSplit.Panel2.Controls.Add(tabs);

            Controls.Add(mainSplit);
            Controls.Add(footer);
            Controls.Add(header);
            LoadProperties();
        }

        private static TabPage CreatePage(string title)
        {
            return new TabPage(title)
            {
                BackColor = Color.White,
                Padding = new Padding(8)
            };
        }

        private static DataGridView CreateGrid()
        {
            return new DataGridView
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
                ColumnHeadersHeight = 36,
                RowTemplate = { Height = 30 },
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
        }

        private static TableLayoutPanel CreateDetailsLayout()
        {
            TableLayoutPanel details = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                ColumnCount = 2,
                Padding = new Padding(18),
                BackColor = Color.White
            };
            details.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 170));
            details.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            return details;
        }

        private void LoadProperties()
        {
            DataTable properties = new DataTable();
            string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + AGI_DB.PathDataBase.Trim() + ";";
            string query = "SELECT [code loc] AS [Code locataire], [CODE GER] AS [Code gérance], " +
                "[nom loc] AS [Locataire], [adresse] AS [Adresse], [datloc] AS [Début location], " +
                "[Loyer], [CHARGE] AS [Charges], [TAXE] AS [Taxes], [total] AS [Total enregistré], " +
                "[impy] AS [Impayé enregistré] FROM [loc] WHERE [code loc] IS NOT NULL " +
                "ORDER BY [nom loc]";

            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(query, connection))
                {
                    adapter.Fill(properties);
                }

                _properties = properties;
                _propertiesGrid.DataSource = _properties;
                if (_propertiesGrid.Columns.Contains("Adresse"))
                {
                    _propertiesGrid.Columns["Adresse"].Visible = false;
                }
                if (_propertiesGrid.Columns.Contains("Début location"))
                {
                    _propertiesGrid.Columns["Début location"].Visible = false;
                }
                foreach (string column in new[] { "Loyer", "Charges", "Taxes", "Total enregistré", "Impayé enregistré" })
                {
                    if (_propertiesGrid.Columns.Contains(column))
                    {
                        _propertiesGrid.Columns[column].Visible = false;
                    }
                }

                _statusLabel.Text = properties.Rows.Count.ToString("N0", CultureInfo.CurrentCulture) +
                    " dossiers chargés. Les logements de la table Appartement ne sont pas reliés aux locataires dans les champs disponibles.";
                ApplySearch();
            }
            catch (Exception ex)
            {
                _statusLabel.Text = "Erreur de chargement.";
                MessageBox.Show("Impossible de charger les fiches logement : " + ex.Message,
                    "Erreur de chargement", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ApplySearch()
        {
            if (_properties == null)
            {
                return;
            }

            string search = _searchBox.Text.Trim();
            DataTable filtered = _properties.Clone();
            foreach (DataRow row in _properties.Rows)
            {
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

            _propertiesGrid.DataSource = filtered;
        }

        private void PropertySelectionChanged(object sender, EventArgs e)
        {
            if (_propertiesGrid.CurrentRow == null)
            {
                ClearDetails();
                return;
            }

            DataRowView selected = _propertiesGrid.CurrentRow.DataBoundItem as DataRowView;
            if (selected == null)
            {
                ClearDetails();
                return;
            }

            DataRow property = selected.Row;
            PopulatePropertyDetails(property);
            LoadLinkedRecords(property);
        }

        private void PopulatePropertyDetails(DataRow property)
        {
            _propertyDetails.SuspendLayout();
            _propertyDetails.Controls.Clear();
            _propertyDetails.RowStyles.Clear();
            _propertyDetails.RowCount = 0;

            if (property == null)
            {
                AddDetail(_propertyDetails, "Fiche logement", "Sélectionnez un dossier dans la liste.");
                _propertyDetails.ResumeLayout(true);
                return;
            }

            AddDetail(_propertyDetails, "Code locataire", GetValue(property, "Code locataire"));
            AddDetail(_propertyDetails, "Locataire", GetValue(property, "Locataire"));
            AddDetail(_propertyDetails, "Adresse", GetValue(property, "Adresse"));
            AddDetail(_propertyDetails, "Début location", GetValue(property, "Début location"));
            AddDetail(_propertyDetails, "Code gérance", GetValue(property, "Code gérance"));
            AddDetail(_propertyDetails, "Loyer", FormatCurrency(property["Loyer"]));
            AddDetail(_propertyDetails, "Charges", FormatCurrency(property["Charges"]));
            AddDetail(_propertyDetails, "Taxes", FormatCurrency(property["Taxes"]));
            AddDetail(_propertyDetails, "Total enregistré", FormatCurrency(property["Total enregistré"]));
            AddDetail(_propertyDetails, "Impayé enregistré", GetValue(property, "Impayé enregistré"));
            _propertyDetails.ResumeLayout(true);
        }

        private void LoadLinkedRecords(DataRow property)
        {
            string codeLocataire = GetValue(property, "Code locataire");
            string codeGerance = GetValue(property, "Code gérance");
            string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + AGI_DB.PathDataBase.Trim() + ";";

            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    DataTable payments = new DataTable();
                    using (OleDbCommand command = new OleDbCommand(
                        "SELECT [quittance] AS [Quittance], [date_encs] AS [Date encaissement], " +
                        "[intitulé] AS [Période], [total] AS [Total], [def_loy] AS [Déficit enregistré] " +
                        "FROM [Encaissement] WHERE [code_ger] = ? AND [code_loc] = ? ORDER BY [date_encs] DESC", connection))
                    {
                        command.Parameters.AddWithValue("?", codeGerance);
                        command.Parameters.AddWithValue("?", codeLocataire);
                        using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                        {
                            adapter.Fill(payments);
                        }
                    }
                    _paymentsGrid.DataSource = payments;

                    DataTable management = new DataTable();
                    using (OleDbCommand command = new OleDbCommand(
                        "SELECT [CODE GER] AS [Code gérance], [nom  ger] AS [Nom gérance], " +
                        "[adresse ger] AS [Adresse], [tel ger] AS [Téléphone] FROM [ger] WHERE [CODE GER] = ?", connection))
                    {
                        command.Parameters.AddWithValue("?", codeGerance);
                        using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                        {
                            adapter.Fill(management);
                        }
                    }
                    PopulateManagementDetails(management);

                    DataTable owners = new DataTable();
                    using (OleDbCommand command = new OleDbCommand(
                        "SELECT [nom_cop] AS [Copropriétaire], [adresse1] AS [Adresse 1], " +
                        "[adresse2] AS [Adresse 2], [adresse3] AS [Adresse 3], [pqrt] AS [Part], [den] AS [Dénominateur] " +
                        "FROM [part] WHERE [code ger] = ? ORDER BY [nom_cop]", connection))
                    {
                        command.Parameters.AddWithValue("?", codeGerance);
                        using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                        {
                            adapter.Fill(owners);
                        }
                    }
                    _ownersGrid.DataSource = owners;
                }
            }
            catch (Exception ex)
            {
                _paymentsGrid.DataSource = null;
                _ownersGrid.DataSource = null;
                PopulateManagementDetails(null);
                MessageBox.Show("Impossible de charger les informations liées au logement : " + ex.Message,
                    "Erreur de chargement", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void PopulateManagementDetails(DataTable management)
        {
            _managementDetails.SuspendLayout();
            _managementDetails.Controls.Clear();
            _managementDetails.RowStyles.Clear();
            _managementDetails.RowCount = 0;
            if (management == null || management.Rows.Count == 0)
            {
                AddDetail(_managementDetails, "Gérance", "Aucune information de gérance trouvée.");
            }
            else
            {
                DataRow row = management.Rows[0];
                foreach (DataColumn column in management.Columns)
                {
                    AddDetail(_managementDetails, column.ColumnName, GetValue(row, column.ColumnName));
                }
            }
            _managementDetails.ResumeLayout(true);
        }

        private void ClearDetails()
        {
            PopulatePropertyDetails(null);
            _paymentsGrid.DataSource = null;
            _ownersGrid.DataSource = null;
            PopulateManagementDetails(null);
        }

        private static void AddDetail(TableLayoutPanel layout, string name, string value)
        {
            int row = layout.RowCount++;
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(new Label
            {
                Text = name,
                AutoSize = true,
                Padding = new Padding(4, 8, 8, 8),
                ForeColor = Color.FromArgb(71, 85, 105),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            }, 0, row);
            layout.Controls.Add(new Label
            {
                Text = string.IsNullOrWhiteSpace(value) ? "—" : value,
                AutoSize = true,
                MaximumSize = new Size(650, 0),
                Padding = new Padding(4, 8, 8, 8),
                ForeColor = Color.FromArgb(39, 52, 67)
            }, 1, row);
        }

        private static string GetValue(DataRow row, string column)
        {
            return row == null || row[column] == DBNull.Value
                ? string.Empty
                : Convert.ToString(row[column], CultureInfo.CurrentCulture);
        }

        private static string FormatCurrency(object value)
        {
            if (value == null || value == DBNull.Value)
            {
                return string.Empty;
            }

            decimal amount;
            string text = Convert.ToString(value, CultureInfo.CurrentCulture);
            if (decimal.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out amount) ||
                decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out amount))
            {
                return amount.ToString("N2", CultureInfo.CurrentCulture) + " DH";
            }
            return text;
        }
    }
}
