using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace PKNK_CNPM.Forms
{
    public class FormHistoriqueActions : Form
    {
        private readonly TextBox _searchBox;
        private readonly ComboBox _actionFilter;
        private readonly DataGridView _actionsGrid;
        private readonly Label _statusLabel;
        private DataTable _allActions;

        public FormHistoriqueActions()
        {
            Text = "Historique des actions";
            BackColor = Color.FromArgb(245, 248, 250);
            Font = new Font("Segoe UI", 9F);
            MinimumSize = new Size(800, 520);

            Panel header = new Panel
            {
                Dock = DockStyle.Top,
                Height = 122,
                Padding = new Padding(18, 8, 18, 8),
                BackColor = Color.White
            };
            Label title = new Label
            {
                Text = "Historique des actions",
                Dock = DockStyle.Top,
                Height = 34,
                Font = new Font("Segoe UI", 16F, FontStyle.Bold),
                ForeColor = Color.FromArgb(15, 70, 75)
            };
            Label description = new Label
            {
                Text = "Journal en lecture seule des créations, modifications et suppressions des encaissements et locataires.",
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
                Width = 390,
                Height = 29,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.FromArgb(39, 52, 67)
            };
            _searchBox.TextChanged += (sender, e) => ApplyFilters();

            _actionFilter = new ComboBox
            {
                Width = 180,
                DropDownStyle = ComboBoxStyle.DropDownList,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9F)
            };
            _actionFilter.Items.AddRange(new object[] { "Toutes les actions", "Création", "Modification", "Suppression" });
            _actionFilter.SelectedIndex = 0;
            _actionFilter.SelectedIndexChanged += (sender, e) => ApplyFilters();

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
            refreshButton.Click += (sender, e) => LoadActions();
            filters.Controls.Add(_searchBox);
            filters.Controls.Add(_actionFilter);
            filters.Controls.Add(refreshButton);
            header.Controls.Add(filters);
            header.Controls.Add(description);
            header.Controls.Add(title);
            ResponsiveScreenLayout.ConfigureFilterHeader(header, filters, title.Height + description.Height);

            _actionsGrid = new DataGridView
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
                Height = 36,
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

            Controls.Add(_actionsGrid);
            Controls.Add(footer);
            Controls.Add(header);
            LoadActions();
        }

        private void LoadActions()
        {
            try
            {
                ActionJournal.EnsureTable();
                DataTable actions = new DataTable();
                string connectionString = ActionJournal.GetConnectionString();
                string query = "SELECT [IdAction] AS [N°], [DateAction] AS [Date et heure], " +
                    "[Utilisateur] AS [Utilisateur], [ActionEffectuee] AS [Action], [Entite] AS [Élément], " +
                    "[Reference] AS [Référence], [Details] AS [Détails] FROM [JournalActions] ORDER BY [DateAction] DESC, [IdAction] DESC";
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(query, connection))
                {
                    adapter.Fill(actions);
                }

                _allActions = actions;
                ApplyFilters();
                _statusLabel.Text = actions.Rows.Count.ToString("N0", CultureInfo.CurrentCulture) +
                    " action(s) journalisée(s). Le journal débute à son activation et ne reconstitue pas les opérations antérieures.";
            }
            catch (Exception ex)
            {
                _allActions = null;
                _actionsGrid.DataSource = null;
                _statusLabel.Text = "Impossible de charger l'historique.";
                MessageBox.Show("Impossible de charger l'historique des actions : " + ex.Message,
                    "Erreur de chargement", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ApplyFilters()
        {
            if (_allActions == null)
            {
                return;
            }

            string search = _searchBox.Text.Trim();
            string selectedAction = _actionFilter.SelectedItem as string;
            DataTable filtered = _allActions.Clone();
            foreach (DataRow row in _allActions.Rows)
            {
                string action = Convert.ToString(row["Action"], CultureInfo.CurrentCulture);
                if (selectedAction != "Toutes les actions" &&
                    !string.Equals(action, selectedAction, StringComparison.CurrentCultureIgnoreCase))
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
            _actionsGrid.DataSource = filtered;
            if (_actionsGrid.Columns.Contains("N°"))
            {
                _actionsGrid.Columns["N°"].FillWeight = 45;
            }
        }
    }

    internal static class ActionJournal
    {
        internal static string GetConnectionString()
        {
            return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + AGI_DB.PathDataBase.Trim() + ";";
        }

        internal static void EnsureTable()
        {
            using (OleDbConnection connection = new OleDbConnection(GetConnectionString()))
            {
                connection.Open();
                DataTable schema = connection.GetSchema("Tables");
                foreach (DataRow row in schema.Rows)
                {
                    if (string.Equals(Convert.ToString(row["TABLE_NAME"], CultureInfo.CurrentCulture),
                        "JournalActions", StringComparison.OrdinalIgnoreCase))
                    {
                        return;
                    }
                }

                using (OleDbCommand command = new OleDbCommand(
                    "CREATE TABLE [JournalActions] ([IdAction] COUNTER PRIMARY KEY, [DateAction] DATETIME NOT NULL, " +
                    "[Utilisateur] TEXT(100), [ActionEffectuee] TEXT(100), [Entite] TEXT(100), " +
                    "[Reference] TEXT(255), [Details] MEMO)", connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }

        internal static void RecordOrWarn(IWin32Window owner, string action, string entity, string reference, string details)
        {
            string error;
            if (!TryRecord(action, entity, reference, details, out error))
            {
                MessageBox.Show(owner, "L'opération a été effectuée, mais son inscription dans l'historique a échoué : " + error,
                    "Avertissement du journal", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        internal static bool TryRecord(string action, string entity, string reference, string details, out string error)
        {
            error = string.Empty;
            try
            {
                EnsureTable();
                using (OleDbConnection connection = new OleDbConnection(GetConnectionString()))
                using (OleDbCommand command = new OleDbCommand(
                    "INSERT INTO [JournalActions] ([DateAction], [Utilisateur], [ActionEffectuee], [Entite], [Reference], [Details]) " +
                    "VALUES (?, ?, ?, ?, ?, ?)", connection))
                {
                    command.Parameters.Add("?", OleDbType.Date).Value = DateTime.Now;
                    command.Parameters.Add("?", OleDbType.VarWChar, 100).Value = Environment.UserName;
                    command.Parameters.Add("?", OleDbType.VarWChar, 100).Value = action;
                    command.Parameters.Add("?", OleDbType.VarWChar, 100).Value = entity;
                    command.Parameters.Add("?", OleDbType.VarWChar, 255).Value = reference ?? string.Empty;
                    command.Parameters.Add("?", OleDbType.LongVarWChar).Value = details ?? string.Empty;
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }
    }
}
