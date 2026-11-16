using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

namespace PKNK_CNPM.Forms
{
    public class FormDocumentsFiables : Form
    {
        private readonly DataGridView _documentsGrid;
        private readonly Label _summaryLabel;
        private readonly Dictionary<string, string> _documentTypes;

        public FormDocumentsFiables()
        {
            Text = "Documents plus fiables";
            BackColor = Color.FromArgb(245, 248, 250);
            Font = new Font("Segoe UI", 9F);
            MinimumSize = new Size(850, 520);
            _documentTypes = new Dictionary<string, string>
            {
                { "Quittance", "Quittances" },
                { "EtatDetaille", "États détaillés" },
                { "ListeRetard", "Listes des retards" },
                { "RetardLocataire", "Retards locataires" },
                { "OrderVirement", "Ordres de virement" },
                { "Releve", "Relevés" }
            };

            Panel header = new Panel
            {
                Dock = DockStyle.Top,
                Height = 105,
                Padding = new Padding(18, 8, 18, 8),
                BackColor = Color.White
            };
            Label title = new Label
            {
                Text = "Documents plus fiables",
                Dock = DockStyle.Top,
                Height = 34,
                Font = new Font("Segoe UI", 16F, FontStyle.Bold),
                ForeColor = Color.FromArgb(15, 70, 75)
            };
            Label description = new Label
            {
                Text = "Vérifiez les chemins configurés, retrouvez les PDF générés et repérez les fichiers vides ou invalides. Aucun fichier n'est modifié.",
                Dock = DockStyle.Fill,
                ForeColor = Color.FromArgb(71, 85, 105)
            };
            header.Controls.Add(description);
            header.Controls.Add(title);

            _documentsGrid = new DataGridView
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
            _documentsGrid.DataBindingComplete += DocumentsGridDataBindingComplete;
            _documentsGrid.CellDoubleClick += (sender, e) => OpenSelectedDocument();

            Panel footer = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 70,
                Padding = new Padding(12, 4, 12, 4),
                BackColor = Color.White
            };
            FlowLayoutPanel actions = new FlowLayoutPanel
            {
                Dock = DockStyle.Right,
                Width = 270,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false
            };
            Button openFileButton = CreateButton("Ouvrir le PDF", Color.FromArgb(15, 118, 110));
            openFileButton.Click += (sender, e) => OpenSelectedDocument();
            Button openFolderButton = CreateButton("Ouvrir le dossier", Color.FromArgb(51, 65, 85));
            openFolderButton.Click += (sender, e) => OpenSelectedFolder();
            Button refreshButton = CreateButton("Vérifier", Color.FromArgb(71, 85, 105));
            refreshButton.Click += (sender, e) => LoadDocuments();
            actions.Controls.Add(openFileButton);
            actions.Controls.Add(openFolderButton);
            actions.Controls.Add(refreshButton);

            _summaryLabel = new Label
            {
                Dock = DockStyle.Fill,
                ForeColor = Color.FromArgb(39, 52, 67),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft
            };
            footer.Controls.Add(actions);
            footer.Controls.Add(_summaryLabel);
            ResponsiveScreenLayout.ConfigureActionFooter(footer, actions, 340, 70);

            Controls.Add(_documentsGrid);
            Controls.Add(footer);
            Controls.Add(header);
            LoadDocuments();
        }

        private static Button CreateButton(string text, Color color)
        {
            Button button = new Button
            {
                Text = text,
                Width = 105,
                Height = 31,
                FlatStyle = FlatStyle.Flat,
                BackColor = color,
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Cursor = Cursors.Hand,
                UseVisualStyleBackColor = false,
                Margin = new Padding(3, 5, 3, 3)
            };
            button.FlatAppearance.BorderSize = 0;
            return button;
        }

        private void LoadDocuments()
        {
            DataTable rows = CreateDocumentTable();
            string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + AGI_DB.PathDataBase.Trim() + ";";

            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                using (OleDbCommand command = new OleDbCommand(
                    "SELECT [Id Param], [Valeur] FROM [Params] WHERE [Id Param] IN " +
                    "('Quittance', 'EtatDetaille', 'ListeRetard', 'RetardLocataire', 'OrderVirement', 'Releve')", connection))
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                {
                    DataTable settings = new DataTable();
                    adapter.Fill(settings);
                    Dictionary<string, string> configuredPaths = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    foreach (DataRow setting in settings.Rows)
                    {
                        string key = Convert.ToString(setting["Id Param"], CultureInfo.CurrentCulture).Trim();
                        string value = setting["Valeur"] == DBNull.Value
                            ? string.Empty
                            : Convert.ToString(setting["Valeur"], CultureInfo.CurrentCulture).Trim();
                        configuredPaths[key] = value;
                    }

                    foreach (KeyValuePair<string, string> documentType in _documentTypes)
                    {
                        string path;
                        if (!configuredPaths.TryGetValue(documentType.Key, out path) || string.IsNullOrWhiteSpace(path))
                        {
                            AddDocumentRow(rows, documentType.Value, "", "", "", "", "Chemin non configuré");
                            continue;
                        }

                        ScanDirectory(rows, documentType.Value, path);
                    }
                }

                _documentsGrid.DataSource = rows;
                int attentionCount = 0;
                foreach (DataRow row in rows.Rows)
                {
                    string status = Convert.ToString(row["Contrôle"], CultureInfo.CurrentCulture);
                    if (status != "PDF valide" && status != "Dossier vide")
                    {
                        attentionCount++;
                    }
                }
                _summaryLabel.Text = rows.Rows.Count.ToString("N0", CultureInfo.CurrentCulture) +
                    " élément(s) vérifié(s) ; " + attentionCount.ToString("N0", CultureInfo.CurrentCulture) +
                    " à corriger ou confirmer. Double-cliquez sur un PDF valide pour l'ouvrir.";
            }
            catch (Exception ex)
            {
                _documentsGrid.DataSource = null;
                _summaryLabel.Text = "La vérification des documents a échoué.";
                MessageBox.Show("Impossible de vérifier les documents configurés : " + ex.Message,
                    "Erreur de vérification", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static DataTable CreateDocumentTable()
        {
            DataTable rows = new DataTable();
            rows.Columns.Add("Type de document", typeof(string));
            rows.Columns.Add("Fichier", typeof(string));
            rows.Columns.Add("Taille", typeof(string));
            rows.Columns.Add("Dernière modification", typeof(string));
            rows.Columns.Add("Contrôle", typeof(string));
            rows.Columns.Add("Chemin", typeof(string));
            return rows;
        }

        private static void ScanDirectory(DataTable rows, string type, string path)
        {
            try
            {
                if (!Directory.Exists(path))
                {
                    AddDocumentRow(rows, type, "", "", "", path, "Dossier introuvable");
                    return;
                }

                string[] files = Directory.GetFiles(path, "*.pdf", SearchOption.TopDirectoryOnly);
                if (files.Length == 0)
                {
                    AddDocumentRow(rows, type, "", "", "", path, "Dossier vide");
                    return;
                }

                foreach (string filePath in files)
                {
                    try
                    {
                        FileInfo file = new FileInfo(filePath);
                        string status = ValidatePdf(filePath, file.Length);
                        AddDocumentRow(rows, type, file.Name, FormatSize(file.Length),
                            file.LastWriteTime.ToString("dd/MM/yyyy HH:mm", CultureInfo.CurrentCulture), filePath, status);
                    }
                    catch (Exception ex)
                    {
                        AddDocumentRow(rows, type, Path.GetFileName(filePath), "", "", filePath,
                            "Lecture impossible : " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                AddDocumentRow(rows, type, "", "", "", path, "Accès impossible : " + ex.Message);
            }
        }

        private static string ValidatePdf(string filePath, long length)
        {
            if (length == 0)
            {
                return "Fichier vide";
            }

            byte[] signature = new byte[5];
            using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                if (stream.Read(signature, 0, signature.Length) != signature.Length)
                {
                    return "Fichier trop court pour être un PDF";
                }
            }

            return signature[0] == 0x25 && signature[1] == 0x50 && signature[2] == 0x44 &&
                signature[3] == 0x46 && signature[4] == 0x2D
                ? "PDF valide"
                : "Signature PDF invalide";
        }

        private static void AddDocumentRow(DataTable rows, string type, string file, string size,
            string modified, string path, string status)
        {
            rows.Rows.Add(type, file, size, modified, status, path);
        }

        private void DocumentsGridDataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            if (_documentsGrid.Columns.Contains("Chemin"))
            {
                _documentsGrid.Columns["Chemin"].Visible = false;
            }
            if (_documentsGrid.Columns.Contains("Fichier"))
            {
                _documentsGrid.Columns["Fichier"].FillWeight = 130;
            }
            if (_documentsGrid.Columns.Contains("Contrôle"))
            {
                _documentsGrid.Columns["Contrôle"].FillWeight = 135;
            }
            foreach (DataGridViewRow row in _documentsGrid.Rows)
            {
                string status = Convert.ToString(row.Cells["Contrôle"].Value, CultureInfo.CurrentCulture);
                if (status != "PDF valide" && status != "Dossier vide")
                {
                    row.Cells["Contrôle"].Style.ForeColor = Color.FromArgb(154, 52, 18);
                    row.Cells["Contrôle"].Style.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
                }
            }
        }

        private void OpenSelectedDocument()
        {
            string path = GetSelectedPath();
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                MessageBox.Show("Sélectionnez un fichier existant.", "Ouvrir le PDF",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string status = Convert.ToString(_documentsGrid.CurrentRow.Cells["Contrôle"].Value, CultureInfo.CurrentCulture);
            if (status != "PDF valide")
            {
                MessageBox.Show("Le fichier n'a pas passé le contrôle PDF ; vérifiez-le avant de l'ouvrir.",
                    "Document à vérifier", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                Process.Start(new ProcessStartInfo { FileName = path, UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show("Impossible d'ouvrir le document : " + ex.Message,
                    "Erreur d'ouverture", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OpenSelectedFolder()
        {
            string path = GetSelectedPath();
            if (string.IsNullOrWhiteSpace(path))
            {
                MessageBox.Show("Sélectionnez un chemin configuré.", "Ouvrir le dossier",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string directory = File.Exists(path) ? Path.GetDirectoryName(path) : path;
            if (string.IsNullOrWhiteSpace(directory) || !Directory.Exists(directory))
            {
                MessageBox.Show("Le dossier sélectionné n'existe pas : " + directory,
                    "Dossier introuvable", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                Process.Start(new ProcessStartInfo { FileName = directory, UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show("Impossible d'ouvrir le dossier : " + ex.Message,
                    "Erreur d'ouverture", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string GetSelectedPath()
        {
            if (_documentsGrid.CurrentRow == null || _documentsGrid.Columns["Chemin"] == null)
            {
                return string.Empty;
            }
            object value = _documentsGrid.CurrentRow.Cells["Chemin"].Value;
            return value == null || value == DBNull.Value
                ? string.Empty
                : Convert.ToString(value, CultureInfo.CurrentCulture);
        }

        private static string FormatSize(long bytes)
        {
            double size = bytes;
            string[] units = { "o", "Ko", "Mo", "Go" };
            int unitIndex = 0;
            while (size >= 1024 && unitIndex < units.Length - 1)
            {
                size /= 1024;
                unitIndex++;
            }
            return size.ToString("N1", CultureInfo.CurrentCulture) + " " + units[unitIndex];
        }
    }
}
