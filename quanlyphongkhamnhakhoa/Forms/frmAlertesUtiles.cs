using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace PKNK_CNPM.Forms
{
    public class FormAlertesUtiles : Form
    {
        private readonly TextBox _searchBox;
        private readonly ComboBox _categoryFilter;
        private readonly ComboBox _priorityFilter;
        private readonly DataGridView _alertsGrid;
        private readonly Label _summaryLabel;
        private DataTable _allAlerts;

        public FormAlertesUtiles()
        {
            Text = "Alertes utiles";
            BackColor = Color.FromArgb(245, 248, 250);
            Font = new Font("Segoe UI", 9F);
            MinimumSize = new Size(850, 530);

            Panel header = new Panel
            {
                Dock = DockStyle.Top,
                Height = 145,
                Padding = new Padding(18, 8, 18, 8),
                BackColor = Color.White
            };
            Label title = new Label
            {
                Text = "Alertes utiles",
                Dock = DockStyle.Top,
                Height = 34,
                Font = new Font("Segoe UI", 16F, FontStyle.Bold),
                ForeColor = Color.FromArgb(15, 70, 75)
            };
            Label description = new Label
            {
                Text = "Repérez les impayés déclarés, les fiches locataires incomplètes et les encaissements nécessitant une vérification.",
                Dock = DockStyle.Top,
                Height = 28,
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
                Width = 310,
                Height = 29,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.FromArgb(39, 52, 67)
            };
            _searchBox.TextChanged += (sender, e) => ApplyFilters();

            _categoryFilter = new ComboBox
            {
                Width = 190,
                DropDownStyle = ComboBoxStyle.DropDownList,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9F)
            };
            _categoryFilter.Items.AddRange(new object[]
            {
                "Toutes les catégories", "Impayés", "Fiches locataires", "Encaissements"
            });
            _categoryFilter.SelectedIndex = 0;
            _categoryFilter.SelectedIndexChanged += (sender, e) => ApplyFilters();

            _priorityFilter = new ComboBox
            {
                Width = 150,
                DropDownStyle = ComboBoxStyle.DropDownList,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9F)
            };
            _priorityFilter.Items.AddRange(new object[] { "Toutes les priorités", "À traiter", "À contrôler" });
            _priorityFilter.SelectedIndex = 0;
            _priorityFilter.SelectedIndexChanged += (sender, e) => ApplyFilters();

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
            refreshButton.Click += (sender, e) => LoadAlerts();
            filters.Controls.Add(_searchBox);
            filters.Controls.Add(_categoryFilter);
            filters.Controls.Add(_priorityFilter);
            filters.Controls.Add(refreshButton);
            header.Controls.Add(filters);
            header.Controls.Add(description);
            header.Controls.Add(title);
            ResponsiveScreenLayout.ConfigureFilterHeader(header, filters, title.Height + description.Height);

            _alertsGrid = new DataGridView
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
                RowTemplate = { Height = 34 },
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
            _alertsGrid.DataBindingComplete += AlertsGridDataBindingComplete;

            Panel footer = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 42,
                Padding = new Padding(12, 6, 12, 4),
                BackColor = Color.White
            };
            _summaryLabel = new Label
            {
                Dock = DockStyle.Fill,
                ForeColor = Color.FromArgb(39, 52, 67),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft
            };
            footer.Controls.Add(_summaryLabel);

            Controls.Add(_alertsGrid);
            Controls.Add(footer);
            Controls.Add(header);
            LoadAlerts();
        }

        private void LoadAlerts()
        {
            DataTable tenants = new DataTable();
            DataTable payments = new DataTable();
            string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + AGI_DB.PathDataBase.Trim() + ";";

            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(
                        "SELECT [code loc], [CODE GER], [nom loc], [adresse], [impy] FROM [loc] WHERE [code loc] IS NOT NULL", connection))
                    {
                        adapter.Fill(tenants);
                    }
                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(
                        "SELECT [N], [quittance], [code_ger], [code_loc], [Nom_Locataire], [date_encs], [total] FROM [Encaissement]", connection))
                    {
                        adapter.Fill(payments);
                    }
                }

                _allAlerts = BuildAlerts(tenants, payments);
                ApplyFilters();
            }
            catch (Exception ex)
            {
                _allAlerts = null;
                _alertsGrid.DataSource = null;
                _summaryLabel.Text = "Échec du chargement des alertes.";
                MessageBox.Show("Impossible de charger les alertes utiles : " + ex.Message,
                    "Erreur de chargement", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static DataTable BuildAlerts(DataTable tenants, DataTable payments)
        {
            DataTable alerts = new DataTable();
            alerts.Columns.Add("Priorité", typeof(string));
            alerts.Columns.Add("Catégorie", typeof(string));
            alerts.Columns.Add("Référence", typeof(string));
            alerts.Columns.Add("Locataire", typeof(string));
            alerts.Columns.Add("Gérance", typeof(string));
            alerts.Columns.Add("Alerte", typeof(string));
            alerts.Columns.Add("Valeur", typeof(string));

            foreach (DataRow tenant in tenants.Rows)
            {
                string code = GetValue(tenant, "code loc");
                string name = GetValue(tenant, "nom loc");
                string management = GetValue(tenant, "CODE GER");
                string arrears = GetValue(tenant, "impy");
                if (HasOutstandingValue(arrears))
                {
                    alerts.Rows.Add("À traiter", "Impayés", code, name, management,
                        "Valeur d'impayé enregistrée (loc.impy)", arrears);
                }
                if (name.Length == 0)
                {
                    alerts.Rows.Add("À traiter", "Fiches locataires", code, name, management,
                        "Nom du locataire manquant", "Compléter la fiche");
                }
                if (management.Length == 0)
                {
                    alerts.Rows.Add("À traiter", "Fiches locataires", code, name, management,
                        "Code de gérance manquant", "Compléter la fiche");
                }
                if (GetValue(tenant, "adresse").Length == 0)
                {
                    alerts.Rows.Add("À contrôler", "Fiches locataires", code, name, management,
                        "Adresse absente", "Vérifier la fiche");
                }
            }

            Dictionary<string, int> receiptCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (DataRow payment in payments.Rows)
            {
                string receipt = GetValue(payment, "quittance");
                string management = GetValue(payment, "code_ger");
                if (receipt.Length > 0 && management.Length > 0)
                {
                    string key = management + "|" + receipt;
                    int count;
                    receiptCounts.TryGetValue(key, out count);
                    receiptCounts[key] = count + 1;
                }
            }

            foreach (DataRow payment in payments.Rows)
            {
                string receipt = GetValue(payment, "quittance");
                string management = GetValue(payment, "code_ger");
                string tenantCode = GetValue(payment, "code_loc");
                string tenantName = GetValue(payment, "Nom_Locataire");
                string reference = "N° " + GetValue(payment, "N") + " / quittance " + (receipt.Length == 0 ? "(absente)" : receipt);

                if (receipt.Length == 0)
                {
                    AddPaymentAlert(alerts, "À traiter", reference, tenantCode, tenantName, management,
                        "Numéro de quittance absent");
                }
                else if (management.Length > 0 && receiptCounts[management + "|" + receipt] > 1)
                {
                    AddPaymentAlert(alerts, "À contrôler", reference, tenantCode, tenantName, management,
                        "Quittance répétée pour cette gérance");
                }
                if (management.Length == 0)
                {
                    AddPaymentAlert(alerts, "À traiter", reference, tenantCode, tenantName, management,
                        "Code de gérance absent");
                }
                if (tenantCode.Length == 0)
                {
                    AddPaymentAlert(alerts, "À traiter", reference, tenantCode, tenantName, management,
                        "Code locataire absent");
                }
                if (payment["date_encs"] == DBNull.Value || GetValue(payment, "date_encs").Length == 0)
                {
                    AddPaymentAlert(alerts, "À traiter", reference, tenantCode, tenantName, management,
                        "Date d'encaissement absente");
                }
                decimal amount;
                if (payment["total"] == DBNull.Value || !TryGetAmount(payment["total"], out amount) || amount <= 0)
                {
                    AddPaymentAlert(alerts, "À contrôler", reference, tenantCode, tenantName, management,
                        "Montant nul, négatif ou non numérique : " + GetValue(payment, "total"));
                }
            }
            return alerts;
        }

        private static void AddPaymentAlert(DataTable alerts, string priority, string reference,
            string tenantCode, string tenantName, string management, string message)
        {
            alerts.Rows.Add(priority, "Encaissements", reference, tenantCode + " - " + tenantName,
                management, message, string.Empty);
        }

        private void ApplyFilters()
        {
            if (_allAlerts == null)
            {
                return;
            }

            string search = _searchBox.Text.Trim();
            string category = _categoryFilter.SelectedItem as string;
            string priority = _priorityFilter.SelectedItem as string;
            DataTable filtered = _allAlerts.Clone();
            foreach (DataRow row in _allAlerts.Rows)
            {
                if (category != "Toutes les catégories" &&
                    !string.Equals(Convert.ToString(row["Catégorie"]), category, StringComparison.CurrentCultureIgnoreCase))
                {
                    continue;
                }
                if (priority != "Toutes les priorités" &&
                    !string.Equals(Convert.ToString(row["Priorité"]), priority, StringComparison.CurrentCultureIgnoreCase))
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

            _alertsGrid.DataSource = filtered;
            int urgent = 0;
            foreach (DataRow row in filtered.Rows)
            {
                if (Convert.ToString(row["Priorité"]) == "À traiter")
                {
                    urgent++;
                }
            }
            _summaryLabel.Text = filtered.Rows.Count.ToString("N0", CultureInfo.CurrentCulture) +
                " alerte(s) affichée(s), dont " + urgent.ToString("N0", CultureInfo.CurrentCulture) +
                " à traiter. Les valeurs sont signalées pour vérification ; aucune donnée n'est modifiée.";
        }

        private void AlertsGridDataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            if (_alertsGrid.Columns.Contains("Alerte"))
            {
                _alertsGrid.Columns["Alerte"].FillWeight = 155;
            }
            if (_alertsGrid.Columns.Contains("Priorité"))
            {
                _alertsGrid.Columns["Priorité"].FillWeight = 65;
            }
            foreach (DataGridViewRow row in _alertsGrid.Rows)
            {
                if (Convert.ToString(row.Cells["Priorité"].Value) == "À traiter")
                {
                    row.Cells["Priorité"].Style.ForeColor = Color.FromArgb(154, 52, 18);
                    row.Cells["Priorité"].Style.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
                }
            }
        }

        private static string GetValue(DataRow row, string column)
        {
            return row[column] == DBNull.Value
                ? string.Empty
                : Convert.ToString(row[column], CultureInfo.CurrentCulture).Trim();
        }

        private static bool HasOutstandingValue(string value)
        {
            string text = (value ?? string.Empty).Trim();
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

        private static bool TryGetAmount(object value, out decimal amount)
        {
            string text = Convert.ToString(value, CultureInfo.CurrentCulture);
            return decimal.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out amount) ||
                decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out amount);
        }
    }
}
