using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace PKNK_CNPM.Forms
{
    public class FormPilotage : Form
    {
        private readonly DataGridView _allTenantsGrid;
        private readonly DataGridView _arrearsGrid;
        private readonly TextBox _tenantSearch;
        private readonly TextBox _arrearsSearch;
        private readonly Label _tenantCountValue;
        private readonly Label _arrearsCountValue;
        private readonly Label _totalValue;
        private readonly Label _statusLabel;
        private readonly TableLayoutPanel _detailsLayout;
        private DataTable _tenants;
        private DataTable _arrears;

        public FormPilotage()
        {
            Text = "Pilotage immobilier";
            BackColor = Color.FromArgb(245, 248, 250);
            Font = new Font("Segoe UI", 9F);
            MinimumSize = new Size(760, 520);

            TabControl tabs = new TabControl
            {
                Dock = DockStyle.Fill,
                Padding = new Point(14, 6),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };
            Controls.Add(tabs);

            TabPage overviewPage = CreatePage("Vue d'ensemble");
            TabPage tenantsPage = CreatePage("Locataires");
            TabPage arrearsPage = CreatePage("Suivi des impayés");
            TabPage detailsPage = CreatePage("Fiche détaillée");
            tabs.TabPages.AddRange(new[] { overviewPage, tenantsPage, arrearsPage, detailsPage });

            FlowLayoutPanel summaryCards = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 142,
                Padding = new Padding(12),
                WrapContents = true,
                AutoScroll = true,
                BackColor = Color.FromArgb(245, 248, 250)
            };
            _tenantCountValue = AddSummaryCard(summaryCards, "LOCATAIRES / LOGEMENTS", "--");
            _arrearsCountValue = AddSummaryCard(summaryCards, "DOSSIERS AVEC IMPAYÉ", "--");
            _totalValue = AddSummaryCard(summaryCards, "TOTAL ENREGISTRÉ", "--");

            _statusLabel = new Label
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(20),
                ForeColor = Color.FromArgb(71, 85, 105),
                Text = "Les retards reprennent la valeur existante dans le champ impy. Aucun montant d'impayé n'est recalculé ici.",
                TextAlign = ContentAlignment.TopLeft
            };
            overviewPage.Controls.Add(_statusLabel);
            overviewPage.Controls.Add(summaryCards);

            _allTenantsGrid = CreateGrid();
            _tenantSearch = CreateSearchBox("Rechercher par code, gérance, locataire ou adresse...");
            tenantsPage.Controls.Add(CreateSearchBar(_tenantSearch, "Locataires enregistrés"));
            tenantsPage.Controls.Add(_allTenantsGrid);

            _arrearsGrid = CreateGrid();
            _arrearsSearch = CreateSearchBox("Rechercher dans les dossiers en retard...");
            arrearsPage.Controls.Add(CreateSearchBar(_arrearsSearch, "Suivi des valeurs d'impayé enregistrées"));
            arrearsPage.Controls.Add(_arrearsGrid);

            _detailsLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                ColumnCount = 2,
                Padding = new Padding(24),
                BackColor = Color.White
            };
            _detailsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 190));
            _detailsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            detailsPage.Controls.Add(_detailsLayout);
            SetDetails(null);

            _tenantSearch.TextChanged += (sender, e) => ApplyTenantFilter();
            _arrearsSearch.TextChanged += (sender, e) => ApplyArrearsFilter();
            _allTenantsGrid.SelectionChanged += TenantSelectionChanged;
            _arrearsGrid.SelectionChanged += TenantSelectionChanged;

            Panel toolbar = new Panel
            {
                Dock = DockStyle.Top,
                Height = 48,
                BackColor = Color.White
            };
            Button refreshButton = new Button
            {
                Text = "Actualiser",
                Dock = DockStyle.Right,
                Width = 120,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(15, 118, 110),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            refreshButton.FlatAppearance.BorderSize = 0;
            refreshButton.Click += (sender, e) => LoadData();
            toolbar.Controls.Add(refreshButton);
            Controls.Add(toolbar);
            toolbar.BringToFront();

            LoadData();
        }

        private static TabPage CreatePage(string title)
        {
            return new TabPage(title)
            {
                BackColor = Color.White,
                Padding = new Padding(8)
            };
        }

        private static Label AddSummaryCard(FlowLayoutPanel container, string title, string initialValue)
        {
            Panel card = new Panel
            {
                Width = 250,
                Height = 104,
                Margin = new Padding(8),
                Padding = new Padding(14),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };
            Label titleLabel = new Label
            {
                Dock = DockStyle.Top,
                Height = 28,
                Text = title,
                ForeColor = Color.FromArgb(100, 116, 139),
                Font = new Font("Segoe UI", 8F, FontStyle.Bold)
            };
            Label valueLabel = new Label
            {
                Dock = DockStyle.Fill,
                Text = initialValue,
                ForeColor = Color.FromArgb(15, 118, 110),
                Font = new Font("Segoe UI", 19F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true
            };
            card.Controls.Add(valueLabel);
            card.Controls.Add(titleLabel);
            container.Controls.Add(card);
            return valueLabel;
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
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(15, 118, 110),
                    ForeColor = Color.White,
                    Font = new Font("Segoe UI", 9F, FontStyle.Bold)
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
                },
                RowTemplate = { Height = 30 },
                ColumnHeadersHeight = 36
            };
        }

        private static TextBox CreateSearchBox(string placeholder)
        {
            return new TextBox
            {
                Dock = DockStyle.Right,
                Width = 330,
                Margin = new Padding(4),
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 9F),
                BackColor = Color.White,
                ForeColor = Color.FromArgb(39, 52, 67),
                Tag = placeholder
            };
        }

        private static Control CreateSearchBar(TextBox searchBox, string title)
        {
            Panel bar = new Panel
            {
                Dock = DockStyle.Top,
                Height = 48,
                Padding = new Padding(8),
                BackColor = Color.White
            };
            Label label = new Label
            {
                Dock = DockStyle.Left,
                Width = 300,
                Text = title,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.FromArgb(15, 118, 110),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };
            searchBox.Text = Convert.ToString(searchBox.Tag, CultureInfo.CurrentCulture);
            searchBox.ForeColor = Color.FromArgb(100, 116, 139);
            searchBox.Enter += (sender, e) =>
            {
                if (searchBox.Text == Convert.ToString(searchBox.Tag, CultureInfo.CurrentCulture))
                {
                    searchBox.Clear();
                    searchBox.ForeColor = Color.FromArgb(39, 52, 67);
                }
            };
            searchBox.Leave += (sender, e) =>
            {
                if (string.IsNullOrWhiteSpace(searchBox.Text))
                {
                    searchBox.Text = Convert.ToString(searchBox.Tag, CultureInfo.CurrentCulture);
                    searchBox.ForeColor = Color.FromArgb(100, 116, 139);
                }
            };
            bar.Controls.Add(searchBox);
            bar.Controls.Add(label);
            return bar;
        }

        private void LoadData()
        {
            DataTable loadedTenants = new DataTable();
            string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + AGI_DB.PathDataBase.Trim() + ";";
            string query = "SELECT [code loc] AS [Code], [CODE GER] AS [Gérance], [nom loc] AS [Locataire], " +
                "[adresse] AS [Adresse], [datloc] AS [Date d'entrée], [Loyer] AS [Loyer], [CHARGE] AS [Charges], " +
                "[TAXE] AS [Taxes], [total] AS [Total enregistré], [impy] AS [Valeur impayée] " +
                "FROM [loc] WHERE [code loc] IS NOT NULL ORDER BY [nom loc]";

            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(query, connection))
                {
                    adapter.Fill(loadedTenants);
                }

                _tenants = loadedTenants;
                _arrears = CreateArrearsTable(_tenants);
                _allTenantsGrid.DataSource = _tenants.DefaultView;
                _arrearsGrid.DataSource = _arrears.DefaultView;
                _tenantCountValue.Text = _tenants.Rows.Count.ToString("N0", CultureInfo.CurrentCulture);
                _arrearsCountValue.Text = _arrears.Rows.Count.ToString("N0", CultureInfo.CurrentCulture);
                _totalValue.Text = CalculateRegisteredTotal(_tenants).ToString("N2", CultureInfo.CurrentCulture);
                _statusLabel.Text = "Données actualisées le " + DateTime.Now.ToString("dd/MM/yyyy à HH:mm", CultureInfo.CurrentCulture) +
                    ". Le suivi des impayés repose sur la valeur enregistrée dans loc.impy ; elle est affichée telle quelle et n'est pas assimilée à un montant.";
                ApplyTenantFilter();
                ApplyArrearsFilter();
                SetDetails(null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Impossible de charger le tableau de bord immobilier : " + ex.Message,
                    "Erreur de chargement", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static DataTable CreateArrearsTable(DataTable tenants)
        {
            DataTable arrears = tenants.Clone();
            foreach (DataRow tenant in tenants.Rows)
            {
                if (HasArrears(tenant["Valeur impayée"]))
                {
                    arrears.ImportRow(tenant);
                }
            }
            return arrears;
        }

        private static bool HasArrears(object value)
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

            string normalized = text.Replace(" ", "").Replace("0", "").Replace("/", "").Replace("-", "").Replace(",", "").Replace(".", "");
            return normalized.Length > 0;
        }

        private static decimal CalculateRegisteredTotal(DataTable tenants)
        {
            decimal total = 0;
            foreach (DataRow row in tenants.Rows)
            {
                if (row["Total enregistré"] == DBNull.Value)
                {
                    continue;
                }

                decimal value;
                string text = Convert.ToString(row["Total enregistré"], CultureInfo.CurrentCulture);
                if (decimal.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out value) ||
                    decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out value))
                {
                    total += value;
                }
            }
            return total;
        }

        private void ApplyTenantFilter()
        {
            if (_tenants == null)
            {
                return;
            }

            string search = GetSearchText(_tenantSearch);
            _allTenantsGrid.DataSource = CreateFilteredView(_tenants, search);
        }

        private void ApplyArrearsFilter()
        {
            if (_arrears == null)
            {
                return;
            }

            string search = GetSearchText(_arrearsSearch);
            _arrearsGrid.DataSource = CreateFilteredView(_arrears, search);
        }

        private static string GetSearchText(TextBox searchBox)
        {
            string text = searchBox.Text;
            if (text == Convert.ToString(searchBox.Tag, CultureInfo.CurrentCulture))
            {
                return string.Empty;
            }
            return text.Trim();
        }

        private static DataView CreateFilteredView(DataTable source, string search)
        {
            DataTable filtered = source.Clone();
            foreach (DataRow row in source.Rows)
            {
                bool matches = string.IsNullOrEmpty(search);
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
            return filtered.DefaultView;
        }

        private void TenantSelectionChanged(object sender, EventArgs e)
        {
            DataGridView grid = sender as DataGridView;
            if (grid == null || grid.CurrentRow == null)
            {
                SetDetails(null);
                return;
            }

            DataRowView rowView = grid.CurrentRow.DataBoundItem as DataRowView;
            if (rowView == null)
            {
                SetDetails(null);
                return;
            }

            SetDetails(rowView.Row);
        }

        private void SetDetails(DataRow row)
        {
            _detailsLayout.SuspendLayout();
            _detailsLayout.Controls.Clear();
            _detailsLayout.RowStyles.Clear();
            _detailsLayout.RowCount = 0;

            string[] fields = { "Code", "Gérance", "Locataire", "Adresse", "Date d'entrée", "Loyer", "Charges", "Taxes", "Total enregistré", "Valeur impayée" };
            foreach (string field in fields)
            {
                string value = row == null || row[field] == DBNull.Value
                    ? "Sélectionnez un dossier dans l'onglet Locataires ou Suivi des impayés."
                    : Convert.ToString(row[field], CultureInfo.CurrentCulture);
                AddDetailRow(field, value);
            }

            _detailsLayout.ResumeLayout(true);
        }

        private void AddDetailRow(string title, string value)
        {
            int rowIndex = _detailsLayout.RowCount++;
            _detailsLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Label titleLabel = new Label
            {
                Text = title,
                AutoSize = true,
                Padding = new Padding(4, 8, 8, 8),
                ForeColor = Color.FromArgb(71, 85, 105),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };
            Label valueLabel = new Label
            {
                Text = value,
                AutoSize = true,
                MaximumSize = new Size(700, 0),
                Padding = new Padding(4, 8, 8, 8),
                ForeColor = Color.FromArgb(39, 52, 67)
            };
            _detailsLayout.Controls.Add(titleLabel, 0, rowIndex);
            _detailsLayout.Controls.Add(valueLabel, 1, rowIndex);
        }
    }
}
