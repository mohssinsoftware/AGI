using System;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Windows.Forms;

namespace PKNK_CNPM.Forms
{
    public class FormSauvegardeRassurante : Form
    {
        private readonly TextBox _sourcePathBox;
        private readonly TextBox _backupFolderBox;
        private readonly DataGridView _backupsGrid;
        private readonly Label _statusLabel;
        private readonly Button _backupButton;

        public FormSauvegardeRassurante()
        {
            Text = "Sauvegarde rassurante";
            BackColor = Color.FromArgb(245, 248, 250);
            Font = new Font("Segoe UI", 9F);
            MinimumSize = new Size(850, 530);

            Panel header = new Panel
            {
                Dock = DockStyle.Top,
                Height = 165,
                Padding = new Padding(18, 8, 18, 8),
                BackColor = Color.White
            };
            Label title = new Label
            {
                Text = "Sauvegarde rassurante",
                Dock = DockStyle.Top,
                Height = 34,
                Font = new Font("Segoe UI", 16F, FontStyle.Bold),
                ForeColor = Color.FromArgb(15, 70, 75)
            };
            Label description = new Label
            {
                Text = "Crée une copie horodatée de la base configurée, puis vérifie sa taille et son empreinte. Aucune restauration automatique ni suppression de sauvegarde.",
                Dock = DockStyle.Top,
                Height = 36,
                ForeColor = Color.FromArgb(71, 85, 105)
            };
            TableLayoutPanel pathLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 2,
                Padding = new Padding(0, 3, 0, 0),
                BackColor = Color.White
            };
            pathLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 115));
            pathLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            pathLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 115));
            pathLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 50));
            pathLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 50));

            Label sourceLabel = CreatePathLabel("Base à sauvegarder");
            _sourcePathBox = new TextBox
            {
                ReadOnly = true,
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.FromArgb(39, 52, 67)
            };
            _sourcePathBox.Text = AGI_DB.PathDataBase.Trim();
            Label sourceEmpty = new Label { Dock = DockStyle.Fill };

            Label folderLabel = CreatePathLabel("Dossier de copie");
            _backupFolderBox = new TextBox
            {
                ReadOnly = true,
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.FromArgb(39, 52, 67)
            };
            _backupFolderBox.Text = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Sauvegardes");
            Button chooseFolderButton = CreateButton("Choisir...", Color.FromArgb(51, 65, 85), 105);
            chooseFolderButton.Click += (sender, e) => ChooseBackupFolder();

            pathLayout.Controls.Add(sourceLabel, 0, 0);
            pathLayout.Controls.Add(_sourcePathBox, 1, 0);
            pathLayout.Controls.Add(sourceEmpty, 2, 0);
            pathLayout.Controls.Add(folderLabel, 0, 1);
            pathLayout.Controls.Add(_backupFolderBox, 1, 1);
            pathLayout.Controls.Add(chooseFolderButton, 2, 1);
            header.Controls.Add(pathLayout);
            header.Controls.Add(description);
            header.Controls.Add(title);

            _backupsGrid = new DataGridView
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
                Height = 72,
                Padding = new Padding(12, 4, 12, 4),
                BackColor = Color.White
            };
            FlowLayoutPanel actions = new FlowLayoutPanel
            {
                Dock = DockStyle.Right,
                Width = 235,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false
            };
            _backupButton = CreateButton("Sauvegarder", Color.FromArgb(15, 118, 110), 125);
            _backupButton.Click += (sender, e) => CreateBackup();
            Button refreshButton = CreateButton("Actualiser", Color.FromArgb(71, 85, 105), 100);
            refreshButton.Click += (sender, e) => LoadBackups();
            actions.Controls.Add(_backupButton);
            actions.Controls.Add(refreshButton);
            _statusLabel = new Label
            {
                Dock = DockStyle.Fill,
                ForeColor = Color.FromArgb(39, 52, 67),
                TextAlign = ContentAlignment.MiddleLeft
            };
            footer.Controls.Add(actions);
            footer.Controls.Add(_statusLabel);
            ResponsiveScreenLayout.ConfigureActionFooter(footer, actions, 250, 72);

            Controls.Add(_backupsGrid);
            Controls.Add(footer);
            Controls.Add(header);
            LoadBackups();
        }

        private static Label CreatePathLabel(string text)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                ForeColor = Color.FromArgb(39, 52, 67),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };
        }

        private static Button CreateButton(string text, Color color, int width)
        {
            Button button = new Button
            {
                Text = text,
                Width = width,
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

        private void ChooseBackupFolder()
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Choisissez le dossier dans lequel conserver les copies de sauvegarde.";
                dialog.SelectedPath = Directory.Exists(_backupFolderBox.Text)
                    ? _backupFolderBox.Text
                    : AppDomain.CurrentDomain.BaseDirectory;
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    _backupFolderBox.Text = dialog.SelectedPath;
                    LoadBackups();
                }
            }
        }

        private void CreateBackup()
        {
            string sourcePath = _sourcePathBox.Text.Trim();
            string backupFolder = _backupFolderBox.Text.Trim();
            if (sourcePath.Length == 0 || !File.Exists(sourcePath))
            {
                MessageBox.Show("La base configurée est introuvable : " + sourcePath,
                    "Source introuvable", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (backupFolder.Length == 0)
            {
                MessageBox.Show("Choisissez un dossier de destination avant de lancer la sauvegarde.",
                    "Dossier requis", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _backupButton.Enabled = false;
            Cursor previousCursor = Cursor;
            Cursor = Cursors.WaitCursor;
            string destinationPath = null;
            try
            {
                string sourceFullPath = Path.GetFullPath(sourcePath);
                string destinationFullFolder = Path.GetFullPath(backupFolder);
                if (string.Equals(Path.GetDirectoryName(sourceFullPath).TrimEnd(Path.DirectorySeparatorChar),
                    destinationFullFolder.TrimEnd(Path.DirectorySeparatorChar), StringComparison.OrdinalIgnoreCase))
                {
                    throw new IOException("Le dossier de sauvegarde ne peut pas être le dossier de la base source.");
                }

                Directory.CreateDirectory(destinationFullFolder);
                string extension = Path.GetExtension(sourceFullPath);
                string baseName = Path.GetFileNameWithoutExtension(sourceFullPath);
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff", CultureInfo.InvariantCulture);
                destinationPath = Path.Combine(destinationFullFolder, baseName + "_backup_" + timestamp + extension);
                int suffix = 1;
                while (File.Exists(destinationPath))
                {
                    destinationPath = Path.Combine(destinationFullFolder,
                        baseName + "_backup_" + timestamp + "_" + suffix.ToString(CultureInfo.InvariantCulture) + extension);
                    suffix++;
                }

                long sourceLengthBefore = new FileInfo(sourceFullPath).Length;
                File.Copy(sourceFullPath, destinationPath, false);
                long sourceLengthAfter = new FileInfo(sourceFullPath).Length;
                long backupLength = new FileInfo(destinationPath).Length;
                if (sourceLengthBefore != sourceLengthAfter || sourceLengthAfter != backupLength ||
                    !FilesHaveSameHash(sourceFullPath, destinationPath))
                {
                    throw new IOException("La copie n'a pas passé la vérification d'intégrité ; aucune sauvegarde n'est déclarée valide.");
                }

                _statusLabel.Text = "Sauvegarde vérifiée : " + destinationPath;
                LoadBackups();
                MessageBox.Show("La copie a été créée et vérifiée (taille et empreinte identiques) :\r\n" + destinationPath +
                    "\r\n\r\nPour une base Access utilisée par d'autres postes, effectuez de préférence la copie quand aucun utilisateur n'écrit de données.",
                    "Sauvegarde vérifiée", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                if (destinationPath != null && File.Exists(destinationPath))
                {
                    try
                    {
                        File.Delete(destinationPath);
                    }
                    catch (Exception cleanupException)
                    {
                        _statusLabel.Text = "La sauvegarde a échoué et le fichier partiel n'a pas pu être supprimé : " + destinationPath;
                        MessageBox.Show("La sauvegarde a échoué : " + ex.Message +
                            "\r\nLe fichier partiel reste à cet emplacement et doit être supprimé manuellement : " +
                            destinationPath + "\r\nErreur de nettoyage : " + cleanupException.Message,
                            "Erreur de sauvegarde", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                _statusLabel.Text = "Échec de la sauvegarde ; aucune copie valide n'a été confirmée.";
                MessageBox.Show("La sauvegarde n'a pas pu être créée ou vérifiée : " + ex.Message,
                    "Erreur de sauvegarde", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = previousCursor;
                _backupButton.Enabled = true;
            }
        }

        private void LoadBackups()
        {
            DataTable backups = new DataTable();
            backups.Columns.Add("Fichier", typeof(string));
            backups.Columns.Add("Taille", typeof(string));
            backups.Columns.Add("Dernière modification", typeof(string));
            backups.Columns.Add("Chemin", typeof(string));

            string sourcePath = _sourcePathBox.Text.Trim();
            string backupFolder = _backupFolderBox.Text.Trim();
            if (sourcePath.Length == 0 || backupFolder.Length == 0)
            {
                _backupsGrid.DataSource = backups;
                _statusLabel.Text = "Source ou dossier de sauvegarde non défini.";
                return;
            }

            try
            {
                if (!Directory.Exists(backupFolder))
                {
                    _backupsGrid.DataSource = backups;
                    _statusLabel.Text = "Le dossier de sauvegarde n'existe pas encore. Il sera créé lors de la prochaine copie.";
                    return;
                }

                string pattern = Path.GetFileNameWithoutExtension(sourcePath) + "_backup_*" + Path.GetExtension(sourcePath);
                string[] files = Directory.GetFiles(backupFolder, pattern, SearchOption.TopDirectoryOnly);
                Array.Sort(files, StringComparer.OrdinalIgnoreCase);
                Array.Reverse(files);
                foreach (string filePath in files)
                {
                    FileInfo info = new FileInfo(filePath);
                    backups.Rows.Add(info.Name, FormatSize(info.Length),
                        info.LastWriteTime.ToString("dd/MM/yyyy HH:mm:ss", CultureInfo.CurrentCulture), filePath);
                }

                _backupsGrid.DataSource = backups;
                if (_backupsGrid.Columns.Contains("Chemin"))
                {
                    _backupsGrid.Columns["Chemin"].Visible = false;
                }
                if (backups.Rows.Count > 0)
                {
                    _statusLabel.Text = backups.Rows.Count.ToString("N0", CultureInfo.CurrentCulture) +
                        " copie(s) trouvée(s). La date seule ne garantit pas l'intégrité : les copies créées depuis cet écran sont vérifiées par empreinte.";
                }
                else
                {
                    _statusLabel.Text = "Aucune copie horodatée trouvée dans le dossier sélectionné.";
                }
            }
            catch (Exception ex)
            {
                _backupsGrid.DataSource = null;
                _statusLabel.Text = "Impossible de lire le dossier de sauvegarde.";
                MessageBox.Show("Impossible de lister les sauvegardes : " + ex.Message,
                    "Erreur de lecture", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static bool FilesHaveSameHash(string firstPath, string secondPath)
        {
            byte[] firstHash;
            byte[] secondHash;
            using (SHA256 algorithm = SHA256.Create())
            using (FileStream first = new FileStream(firstPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                firstHash = algorithm.ComputeHash(first);
            }
            using (SHA256 algorithm = SHA256.Create())
            using (FileStream second = new FileStream(secondPath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                secondHash = algorithm.ComputeHash(second);
            }
            if (firstHash.Length != secondHash.Length)
            {
                return false;
            }
            for (int i = 0; i < firstHash.Length; i++)
            {
                if (firstHash[i] != secondHash[i])
                {
                    return false;
                }
            }
            return true;
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
