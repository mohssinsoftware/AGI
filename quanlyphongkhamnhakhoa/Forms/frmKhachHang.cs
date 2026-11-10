using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;

using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms; 

namespace PKNK_CNPM.Forms
{
    public partial class FormCustomer : Form
    {
        
        private OleDbConnection _connection;
        private OleDbDataAdapter _dataAdapter;
        private OleDbDataAdapter _dataAdapter_Retard;
        private DataTable _dataTable;
        private DataTable _dataTableRetard;
        private string PathListeRetard = "";
        private string PathRetardLocataire = "";

        public FormCustomer()
        {
            InitializeComponent();
            this.Icon = new Icon("AGI.ico");
            InitializeDatabase();
            LoadData();
            FnGetPath();
            LoadGerData();
            LoadGerDataMAJ();
        }

        /****************************************************************************************************/
        private void InitializeDatabase()
        {
            string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+ AGI_DB.PathDataBase + "";
            _connection = new OleDbConnection(connectionString);
        }

        private void LoadData()
        {   
            try 
            {   
                if (_connection.State != ConnectionState.Open)
                {   
                    _connection.Open();
                }

                // Load data from the database
                _dataAdapter = new OleDbDataAdapter(
                    "SELECT [code loc] AS [Code Locataire], [nom loc] AS [Nom Locataire], " +
                    "[CODE GER] AS [Code Gerance], [total] AS [Loyer], [datloc] AS [Date Location] " +
                    "FROM loc WHERE [code loc] IS NOT NULL " +
                    "ORDER BY IIF(IsNumeric([code loc] & ''), Val([code loc]), NULL), [code loc]",
                    _connection);

                OleDbCommandBuilder commandBuilder = new OleDbCommandBuilder(_dataAdapter);

                _dataTable = new DataTable();
                _dataAdapter.Fill(_dataTable);

                if (_dataTable.Columns.Contains("Loyer"))
                {
                    // Add a new column for the formatted Loyer
                    _dataTable.Columns.Add("Formatted Loyer", typeof(string));

                    foreach (DataRow row in _dataTable.Rows)
                    {
                        if (double.TryParse(row["Loyer"]?.ToString(), out double LoyerValue))
                        {
                            row["Formatted Loyer"] = $"{LoyerValue:0.00} DH"; // Format with 2 decimal places
                        }
                        else
                        {
                            row["Formatted Loyer"] = "Invalid Value"; // Handle non-numeric cases
                        }
                    }

                    // Optional: Remove the original 'Loyer' column if not needed
                    _dataTable.Columns.Remove("Loyer");
                    _dataTable.Columns["Formatted Loyer"].ColumnName = "Loyer";
                }

                DataGridViewListLocataires.DataSource = _dataTable;

                //DEBUT : Ajustement automatique pour les colonnes restantes-------------------------------------------------
                DataGridViewListLocataires.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                // Ajustement spécifique pour la colonne "Nom Locataire"
                DataGridViewListLocataires.Columns["Nom Locataire"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; // Désactiver l'ajustement automatique
                DataGridViewListLocataires.Columns["Nom Locataire"].Width = 400; // Définir la largeur souhaitée
                //FIN : Ajustement automatique pour les colonnes restantes-------------------------------------------------

                _dataAdapter_Retard = new OleDbDataAdapter("SELECT [code loc] AS [Code Locataire], [nom loc] AS [Nom Locataire]," +
                " [CODE GER] AS [Code Gerance], [Loyer] AS [Loyer], [impy]  AS [Retards] " +
                "  FROM loc  WHERE [code loc] IS NOT NULL ORDER BY IIF(IsNumeric([code loc] & ''), Val([code loc]), NULL), [code loc]", _connection);

                OleDbCommandBuilder commandBuilder_Retard = new OleDbCommandBuilder(_dataAdapter_Retard);

                _dataTableRetard = new DataTable();
                _dataAdapter_Retard.Fill(_dataTableRetard);
                if (_dataTableRetard.Columns.Contains("Loyer"))
                {
                    // Add a new column for the formatted Loyer
                    _dataTableRetard.Columns.Add("Formatted Loyer", typeof(string));

                    foreach (DataRow row in _dataTableRetard.Rows)
                    {
                        if (double.TryParse(row["Loyer"]?.ToString(), out double LoyerValue))
                        {
                            row["Formatted Loyer"] = $"{LoyerValue:0.00} DH"; // Format with 2 decimal places
                        }
                        else
                        {
                            row["Formatted Loyer"] = "Invalid Value"; // Handle non-numeric cases
                        }
                    }

                    // Optional: Remove the original 'Loyer' column if not needed
                    _dataTableRetard.Columns.Remove("Loyer");
                    _dataTableRetard.Columns["Formatted Loyer"].ColumnName = "Loyer";
                }
                GridViewLocataireRetard.DataSource = _dataTableRetard;

                //DEBUT : Ajustement automatique pour les colonnes restantes-------------------------------------------------
                GridViewLocataireRetard.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                // Ajustement spécifique pour la colonne "Nom Locataire"
                GridViewLocataireRetard.Columns["Nom Locataire"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; // Désactiver l'ajustement automatique
                GridViewLocataireRetard.Columns["Nom Locataire"].Width = 400; // Définir la largeur souhaitée
                //FIN : Ajustement automatique pour les colonnes restantes-------------------------------------------------

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading data: " + ex.Message);
            }
            finally
            {
                _connection.Close();
            }
        }
        private void FnGetPath()
        {
            try
            {
                // Adjust the SQL query to match the exact table and column names
                string selectQuery = "SELECT [Valeur], [Id Param] FROM [Params] WHERE [Id Param] in ('ListeRetard', 'RetardLocataire')";

                // Open the connection
                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open();
                }

                using (OleDbCommand command = new OleDbCommand(selectQuery, _connection))
                {
                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                    {
                        DataTable dataTable = new DataTable();

                        adapter.Fill(dataTable);

                        for(int i = 0; i < dataTable.Rows.Count; i++)
                        {
                            DataRow row = dataTable.Rows[i];

                            if(row["Id Param"].ToString() == "ListeRetard")
                            {
                                PathListeRetard = row["Valeur"].ToString();
                            }
                            else
                            {
                                PathRetardLocataire = row["Valeur"].ToString();
                            }
                            //PathRetardLocataire  PathListeRetard
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading Affaires: " + ex.Message + "\n\n" + ex.StackTrace);
            }
            finally
            {
                if (_connection.State == ConnectionState.Open)
                {
                    _connection.Close();
                }
            }
        }

        private void LoadGerDataMAJ()
        {
            
            try
            {
                string selectQueryLocataire = @" SELECT [code loc] FROM [loc] WHERE [code loc] IS NOT NULL ORDER BY IIF(IsNumeric([code loc] & ''), Val([code loc]), NULL), [code loc]";

                // Open the connection
                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open(); // Open the connection
                }
                
                // Fetch the data for loc
                using (OleDbDataAdapter adapter1 = new OleDbDataAdapter(selectQueryLocataire, _connection))
                {
                    DataTable dataTable1 = new DataTable();
                    adapter1.Fill(dataTable1);

                    DataRow emptyRow = dataTable1.NewRow();
                    emptyRow["code loc"] = DBNull.Value; // or string.Empty if you prefer empty strings
                    emptyRow["code loc"] = string.Empty; // or set a placeholder text like "<Select>"
                    dataTable1.Rows.InsertAt(emptyRow, 0);

                    MAJCodeLocataire.DisplayMember = "code loc"; // Display this column in ComboBox
                    MAJCodeLocataire.ValueMember = "code loc"; // Use this column as the value

                    MAJCodeLocataire.DataSource = dataTable1;
                }
                
            }
            catch (Exception ex)
            {
                // Output the full exception details to help diagnose the issue
                MessageBox.Show("Error loading ger data: " + ex.Message + "\n\n" + ex.StackTrace);
            }
            finally
            {
                // Ensure the connection is closed even if an error occurs
                if (_connection.State == ConnectionState.Open)
                {
                    _connection.Close();
                }
            }
            
        }

        private void LoadGerData()
        {
            try
            {
                //Adjust the SQL query to match the exact table and column names
                //string selectQuery = "SELECT [CODE GER], [nom  ger] FROM [ger]";
                string selectQuery = "SELECT [CODE GER] FROM [ger] order by [CODE GER]";

                // Open the connection
                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open(); // Open the connection
                }

                // Fetch the data for ger
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(selectQuery, _connection))
                {
                    DataTable dataTable = new DataTable();
                    adapter.Fill(dataTable);

                    DataRow emptyRow = dataTable.NewRow();
                    emptyRow["CODE GER"] = DBNull.Value; // or string.Empty if you prefer empty strings
                    dataTable.Rows.InsertAt(emptyRow, 0);

                    txtCodeGerance.DisplayMember = "CODE GER"; // Display this column in ComboBox
                    txtCodeGerance.ValueMember = "CODE GER"; // Use this column as the value
                    txtCodeGerance.DataSource = dataTable;

                    MAJCodeGerance.DisplayMember = "CODE GER"; // Display this column in ComboBox
                    MAJCodeGerance.ValueMember = "CODE GER"; // Use this column as the value
                    MAJCodeGerance.DataSource = dataTable;
                }
                /*
                // Fetch the data for loc
                using (OleDbDataAdapter adapter1 = new OleDbDataAdapter(selectQueryLocataire, _connection))
                {
                    DataTable dataTable1 = new DataTable();
                    adapter1.Fill(dataTable1);

                    MAJNomLocataire.DisplayMember = "nom loc"; // Display this column in ComboBox
                    MAJNomLocataire.ValueMember = "code loc"; // Use this column as the value

                    MAJNomLocataire.DataSource = dataTable1;
                }
                */
            }
            catch (Exception ex)
            {
                // Output the full exception details to help diagnose the issue
                MessageBox.Show("Error loading ger data: " + ex.Message + "\n\n" + ex.StackTrace);
            }
            finally
            {
                // Ensure the connection is closed even if an error occurs
                if (_connection.State == ConnectionState.Open)
                {
                    _connection.Close();
                }
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (DataGridViewListLocataires.SelectedRows.Count > 0)
            {
                try
                {
                    // Get the selected row
                    DataGridViewRow selectedRow = DataGridViewListLocataires.SelectedRows[0];

                    // Get the unique value (e.g., code loc) of the row to delete
                    string codeLoc = selectedRow.Cells["Code Locataire"].Value.ToString();

                    // Remove the row from DataGridView
                    DataGridViewListLocataires.Rows.Remove(selectedRow);

                    // Delete the corresponding row from the database
                    string deleteQuery = "DELETE FROM loc WHERE [code loc] = @codeLoc";

                    using (OleDbCommand cmd = new OleDbCommand(deleteQuery, _connection))
                    {
                        cmd.Parameters.AddWithValue("@codeLoc", codeLoc);

                        if (_connection.State != ConnectionState.Open)
                        {
                            _connection.Open();
                        }

                        int rowsAffected = cmd.ExecuteNonQuery();
                        if (rowsAffected > 0)
                        {
                            MessageBox.Show("     Le locataire a été supprimé avec succès.     ");
                        }
                        else
                        {
                            MessageBox.Show("     Erreur lors de la suppression du locataire.      ");
                        }
                    }

                    // Refresh the DataGridView after deletion
                    LoadData();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("        Erreur du supprission dans la ligne : " + ex.Message        );
                }
                finally
                {
                    if (_connection.State == ConnectionState.Open)
                    {
                        _connection.Close();
                    }
                }
            }
            else
            {
                MessageBox.Show("Veuillez sélectionner une ligne à supprimer.");
            }
        }


        private void SaveChanges()
        {
            try
            {
                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open(); // Open the connection
                }
                _dataAdapter.Update(_dataTable);
                MessageBox.Show("Changes saved successfully!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error saving changes: " + ex.Message);
            }
            finally
            {
                _connection.Close();
            }
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            LoadData();
        }

        private void btnTaiLai_Click(object sender, EventArgs e)
        {
            
            if (DataGridViewListLocataires.SelectedRows.Count > 0)
            {
                DataGridViewRow selectedRow = DataGridViewListLocataires.SelectedRows[0];
                selectedRow.Cells["nom loc"].Value = txtAdresse.Text;
                selectedRow.Cells["CODE GER"].Value = txtLoyer.Text;
                // Update other fields as necessary

                SaveChanges();
            }
            else
            {
                MessageBox.Show("Please select a row to update.");
            }

            
            //BindGrid(khachHangService.GetAll());
            //clearValue();
        }

        private void btnCanLamSan_Click_1(object sender, EventArgs e)
        {
            try
            {
                string TypeHabitation = "H";
                if (radiobtnCommercial.Checked)
                {
                    TypeHabitation = "C";
                }
                if (radiobtnDepot.Checked)
                {
                    TypeHabitation = "D";
                }

                string textCodeLocataire_ = String.IsNullOrWhiteSpace(textCodeLocataire.Text) ? String.Empty : textCodeLocataire.Text;
                string txtCodeGerance_ = txtCodeGerance.SelectedValue != null ? txtCodeGerance.SelectedValue.ToString() : String.Empty;
                float textCharge_ = 0;
                if (!String.IsNullOrWhiteSpace(textCharge.Text))
                {
                    string formattedText = textCharge.Text.Replace(',', '.');
                    if (!float.TryParse(formattedText, NumberStyles.Float, CultureInfo.InvariantCulture, out textCharge_))
                    {
                        MessageBox.Show("Veuillez entrer un montant valide !");
                        return;
                    }
                }
                float textLoyer_ = 0;
                if (!String.IsNullOrWhiteSpace(txtLoyer.Text))
                {
                    string formattedText = txtLoyer.Text.Replace(',', '.');
                    if (!float.TryParse(formattedText, NumberStyles.Float, CultureInfo.InvariantCulture, out textLoyer_))
                    {
                        MessageBox.Show("Veuillez entrer un montant valide !");
                        return;
                    }
                }
                float textTotal_ = 0;
                if (!String.IsNullOrWhiteSpace(txtLoyer.Text))
                {
                    string formattedText = txtLoyer.Text.Replace(',', '.');
                    if (!float.TryParse(formattedText, NumberStyles.Float, CultureInfo.InvariantCulture, out textTotal_))
                    {
                        MessageBox.Show("Veuillez entrer un montant valide !");
                        return;
                    }
                }
                float textTaxt_ = 0;
                if (!String.IsNullOrWhiteSpace(textTaxt.Text))
                {
                    string formattedText = textTaxt.Text.Replace(',', '.');
                    if (!float.TryParse(formattedText, NumberStyles.Float, CultureInfo.InvariantCulture, out textTaxt_))
                    {
                        MessageBox.Show("Veuillez entrer un montant valide !");
                        return;
                    }
                }
                string datloc_ = String.IsNullOrWhiteSpace(textDateLoyer.Text) ? DateTime.Today.ToString() : DateTime.Parse(textDateLoyer.Text).ToString();
                string textNomLocataire_ = String.IsNullOrWhiteSpace(textNomLocataire.Text) ? String.Empty : textNomLocataire.Text;
                string txtAdresse_ = String.IsNullOrWhiteSpace(txtAdresse.Text) ? String.Empty : txtAdresse.Text;

                string insertQuery = "INSERT INTO loc ([code loc], [CODE GER], [CHARGE], [Loyer], [datloc], [nom loc], [TAXE], [adresse1], [Champ1], [total]) " +
                                     "VALUES (@codeLoc, @codeGer, @charge, @Loyer, @datloc, @nomLoc, @taxe, @adresse1, @TypeHabitation, @Loyer)";

                using (OleDbCommand cmd = new OleDbCommand(insertQuery, _connection))
                {
                    cmd.Parameters.AddWithValue("@codeLoc", textCodeLocataire_);
                    cmd.Parameters.AddWithValue("@codeGer", txtCodeGerance_);
                    cmd.Parameters.AddWithValue("@charge", textCharge_);
                    cmd.Parameters.AddWithValue("@Loyer", textLoyer_);
                    cmd.Parameters.AddWithValue("@datloc", datloc_) ;
                    cmd.Parameters.AddWithValue("@nomLoc", textNomLocataire_);
                    cmd.Parameters.AddWithValue("@taxe", textTaxt_);
                    cmd.Parameters.AddWithValue("@adresse1", txtAdresse_);
                    cmd.Parameters.AddWithValue("@TypeHabitation", TypeHabitation);
                    // Add more parameters as needed

                    if (_connection.State != ConnectionState.Open)
                    {
                        _connection.Open(); // Open the connection
                    }
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Locataire Bien Ajouté ");
                    Refresh();
                    LoadData();
                    LoadGerDataMAJ();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Veuillez verifier les champs de saisie ! : ");
            }
            finally
            {
                _connection.Close();
            }
            
        }

        private void Refresh()
        {
            try
            {
                textCodeLocataire.Text = null;
                txtCodeGerance.SelectedValue = DBNull.Value;
                textCharge.Text = null;
                txtLoyer.Text = null;
                textNomLocataire.Text = null;
                textTaxt.Text = null;
                txtAdresse.Text = null;
            }
            catch (Exception ex)
            {

            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                string TypeHabitation = "H";
                if (MAJradiobtnCommercial.Checked)
                {
                    TypeHabitation = "C";
                }
                if (MAJradiobtnDepot.Checked)
                {
                    TypeHabitation = "D";
                }

                if (!int.TryParse(txtCodeGerance.SelectedValue?.ToString(), out int codeGer))
                {
                    return;
                }

                string textCodeLocataire_ = String.IsNullOrWhiteSpace(MAJCodeLocataire.Text) ? String.Empty : MAJCodeLocataire.Text;
                float textCharge_ = 0, textLoyer_ = 0, textTaxt_ = 0;

                if (!String.IsNullOrWhiteSpace(MAJCharge.Text) &&
                    !float.TryParse(MAJCharge.Text.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out textCharge_))
                {
                    MessageBox.Show("Veuillez entrer un montant valide pour les charges !");
                    return;
                }
                if (!String.IsNullOrWhiteSpace(MAJLoyer.Text) &&
                    !float.TryParse(MAJLoyer.Text.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out textLoyer_))
                {
                    MessageBox.Show("Veuillez entrer un montant valide pour le loyer !");
                    return;
                }
                if (!String.IsNullOrWhiteSpace(MAJTaxt.Text) &&
                    !float.TryParse(MAJTaxt.Text.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out textTaxt_))
                {
                    MessageBox.Show("Veuillez entrer un montant valide pour la taxe !");
                    return;
                }

                string datloc_ = DateTime.TryParse(MAJDateLoyer.Text, out var parsedDate) ? parsedDate.ToString("yyyy-MM-dd") : DateTime.Today.ToString("yyyy-MM-dd");
                string textNomLocataire_ = String.IsNullOrWhiteSpace(MAJNomLocataire.Text) ? String.Empty : MAJNomLocataire.Text;
                string txtAdresse_ = String.IsNullOrWhiteSpace(MAJAdresse.Text) ? String.Empty : MAJAdresse.Text;

                string updateQuery = "UPDATE loc SET [CODE GER] = @codeGer, [CHARGE] = @charge, [LOYER] = @loyer, " +
                                     "[datloc] = @datloc, [nom loc] = @nomLoc, [TAXE] = @taxe, [adresse1] = @adresse1, [Champ1] = @Champ1 " +
                                     "WHERE [code loc] = @codeLoc";

                using (OleDbCommand cmd = new OleDbCommand(updateQuery, _connection))
                {
                    cmd.Parameters.AddWithValue("@codeGer", codeGer);
                    cmd.Parameters.AddWithValue("@charge", textCharge_);
                    cmd.Parameters.AddWithValue("@loyer", textLoyer_);
                    cmd.Parameters.AddWithValue("@datloc", datloc_);
                    cmd.Parameters.AddWithValue("@nomLoc", textNomLocataire_);
                    cmd.Parameters.AddWithValue("@taxe", textTaxt_);
                    cmd.Parameters.AddWithValue("@adresse1", txtAdresse_);
                    cmd.Parameters.AddWithValue("@Champ1", TypeHabitation); // Add this here
                    cmd.Parameters.AddWithValue("@codeLoc", textCodeLocataire_);

                    if (_connection.State != ConnectionState.Open)
                    {
                        _connection.Open();
                    }

                    int rowsAffected = cmd.ExecuteNonQuery();
                    if (rowsAffected > 0)
                    {
                        MessageBox.Show("Locataire bien modifié !");
                        LoadData();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Veuillez verifier les champs de saisie ! : {ex.Message}");
            }
            finally
            {
                _connection.Close();
            }
        }

        private void MAJCodeGerance_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        /****************************************************************************************************/

        private void FormCustomer_Load(object sender, EventArgs e)
        {
            LoadTheme();
            setCount();
        }
        private void LoadTheme()
        {
            foreach (Control btns in this.Controls)
            {
                if (btns.GetType() == typeof(Button))
                {   
                    Button btn= (Button)btns;
                    btns.BackColor = Themecolor.PrimaryColor;
                    btns.ForeColor = Color.White;
                    btn.FlatAppearance.BorderColor = Themecolor.SecondaryColor;
                    
                }
            }
        }

        private void setCount()
        {
         //   txtTong.Text = khachHangService.GetAll().Count.ToString();
        }

        private void clearValue()
        {
        }




        private void btnSua_Click(object sender, EventArgs e)
        {

        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
 
        }

        private void btnTimKiem_Click(object sender, EventArgs e)
        {
            
            
        }

        private void btnKhamLamSan_Click(object sender, EventArgs e)
        {
        }

        private void btnCanLamSan_Click(object sender, EventArgs e)
        {
        }


        private void panel7_Paint(object sender, PaintEventArgs e)
        {

        }

        private void MAJCodeLocatire_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void txtRecherch_TextChanged(object sender, EventArgs e)
        {
            DataView dv = _dataTable.DefaultView;
            dv.RowFilter = string.Format("[Nom Locataire] LIKE '%{0}%'", txtRecherch.Text.Trim());
            DataGridViewListLocataires.DataSource = dv;
        }

        private void MAJNomLocataire_SelectedIndexChanged(object sender, EventArgs e)
        {
            ChargerDonneesLocataireParCode();
        }

        private void ChargerDonneesLocataireParCode()
        {
            try
            {
                // Get the selected value from the ComboBox, assuming the ValueMember is set to "code loc"
                string selectedCodeLoc = MAJCodeLocataire.SelectedValue?.ToString();

                if (string.IsNullOrEmpty(selectedCodeLoc) && string.IsNullOrEmpty(MAJCodeLocataire.Text))
                {
                    return;
                }

                if (string.IsNullOrEmpty(selectedCodeLoc))
                {
                    selectedCodeLoc = MAJCodeLocataire.Text;
                }

                // Define the SQL query with a WHERE clause to filter by the selected "code loc"
                string selectQuery = "SELECT [code loc], [CODE GER], [CHARGE], [Loyer], [total], [datloc], [nom loc], [TAXE], [adresse], [Champ1] " +
                                     "FROM [loc] WHERE [code loc] = ?";

                // Check if the connection is not already open
                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open(); // Open the connection
                }

                // Use OleDbCommand to fetch the data with parameterized query
                using (OleDbCommand command = new OleDbCommand(selectQuery, _connection))
                {
                    // Add parameter to the query to prevent SQL injection
                    command.Parameters.AddWithValue("?", selectedCodeLoc);

                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                    {
                        // Fill a DataTable with the results
                        DataTable dataTable = new DataTable();

                        adapter.Fill(dataTable); // Fill the DataTable

                        if (dataTable.Rows.Count > 0)
                        {

                            DataRow row = dataTable.Rows[0];

                            MAJNomLocataire.Text = row["nom loc"].ToString();
                            MAJCharge.Text = row["CHARGE"].ToString();
                            MAJLoyer.Text = row["Loyer"].ToString();
                            MAJDateLoyer.Text = row["datloc"].ToString();
                            MAJTaxt.Text = row["TAXE"].ToString();
                            MAJTotal.Text = row["total"].ToString();
                            MAJAdresse.Text = row["adresse"].ToString();
                            MAJCodeGerance.SelectedValue = row["CODE GER"].ToString();

                            if (row["Champ1"].ToString() == "D")
                            {
                                MAJradiobtnDepot.Checked = true;
                            }
                            else if (row["Champ1"].ToString() == "C")
                            {
                                MAJradiobtnCommercial.Checked = true;
                            }
                            else
                            {
                                MAJradiobtnHabitation.Checked = true;
                            }
                        }
                        else
                        {
                            MessageBox.Show("No data found for the selected locataire.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading locataire data: " + ex.Message);
            }
            finally
            {
                // Ensure the connection is closed even if an error occurs
                if (_connection.State == ConnectionState.Open)
                {
                    _connection.Close();
                }
            }

        }

        private void btnRechercheCodeLocataire_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtRecherchCodeLocataire.Text != "")
                {
                    DataView dv = _dataTable.DefaultView;
                    dv.RowFilter = string.Format("[Code Locataire] = '{0}'", txtRecherchCodeLocataire.Text);
                    DataGridViewListLocataires.DataSource = dv;
                }
                else
                {
                    DataView dv = _dataTable.DefaultView;
                    dv.RowFilter = string.Format("","");
                    DataGridViewListLocataires.DataSource = dv;
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void btnRechercheCodeGerance_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtRecherchCodeGerance.Text != "")
                {
                    DataView dv = _dataTable.DefaultView;
                    dv.RowFilter = string.Format("[Code Gerance] = '{0}'", txtRecherchCodeGerance.Text);
                    DataGridViewListLocataires.DataSource = dv;
                }
                else
                {
                    DataView dv = _dataTable.DefaultView;
                    dv.RowFilter = string.Format("", "");
                    DataGridViewListLocataires.DataSource = dv;
                }
            }
            catch
            {

            }
        }

        private void txtCodeGerance_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                string selectedCodeGerance = txtCodeGerance.SelectedValue?.ToString();

                if (string.IsNullOrEmpty(selectedCodeGerance))
                {
                    return;
                }

                string selectQuery = " SELECT [adresse1] FROM [ger] WHERE [CODE GER] = @CodeGerance";

                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open(); // Open the connection
                }

                // Use OleDbCommand to fetch the data with parameterized query
                using (OleDbCommand command = new OleDbCommand(selectQuery, _connection))
                {
                    // Add parameter to the query to prevent SQL injection
                    command.Parameters.AddWithValue("@CodeGerance", selectedCodeGerance);

                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                    {
                        // Fill a DataTable with the results
                        DataTable dataTable = new DataTable();

                        adapter.Fill(dataTable); // Fill the DataTable

                        if (dataTable.Rows.Count > 0)
                        {
                            // Assuming you want to show the first row's data
                            DataRow row = dataTable.Rows[0];

                            txtAdresse.Text = row["adresse1"].ToString();
                            MAJAdresse.Text = row["adresse1"].ToString();
                        }
                        else
                        {
                            MessageBox.Show("Merci de vérifier si le nom et l'adresse de ce gérance est bien enregistrer au niveau du base de donnée !");
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading Adresse Gérance data: " + ex.Message);
            }
            finally
            {
                // Ensure the connection is closed even if an error occurs
                if (_connection.State == ConnectionState.Open)
                {
                    _connection.Close();
                }
            }

        }

        private void MAJCodeGerance_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            txtCodeGerance_SelectedIndexChanged(sender, e);
        }

        private void textRechercheLocNomRetard_TextChanged(object sender, EventArgs e)
        {
            DataView dv = _dataTableRetard.DefaultView;
            dv.RowFilter = string.Format("[Nom Locataire] LIKE '%{0}%'", textRechercheLocNomRetard.Text.Trim());
            GridViewLocataireRetard.DataSource = dv;
        }

        private void txtRechrcheCodeLocRetard_TextChanged(object sender, EventArgs e)
        {
            DataView dv = _dataTableRetard.DefaultView;
            //dv.RowFilter = string.Format("[Code Locataire] LIKE '%{0}%'", txtRechrcheCodeLocRetard.Text.Trim());
            GridViewLocataireRetard.DataSource = dv;
        }

        private void RecCodeGerRetard_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtRecherchCodeGerRetard.Text != "")
                {
                    DataView dv = _dataTableRetard.DefaultView;
                    dv.RowFilter = string.Format("[Code Gerance] = '{0}'", txtRecherchCodeGerRetard.Text);
                    GridViewLocataireRetard.DataSource = dv;
                }
                else
                {
                    DataView dv = _dataTableRetard.DefaultView;
                    dv.RowFilter = string.Format("", "");
                    GridViewLocataireRetard.DataSource = dv;
                }
            }
            catch
            {

            }
        }

        private void textCodeLocataire_TextChanged(object sender, EventArgs e)
        {
            try
            {
                // Check if the input value is null or empty
                if (string.IsNullOrWhiteSpace(textCodeLocataire.Text))
                {
                    // Clear the error message and reset the background color
                    textShowError.Text = string.Empty;
                    textCodeLocataire.BackColor = SystemColors.Window; // Default background color
                    _dataTable.DefaultView.RowFilter = string.Empty;
                    DataGridViewListLocataires.DataSource = _dataTable.DefaultView;
                    return; // Exit the method
                }

                // Filter the DataView based on the search term
                DataView dv = _dataTable.DefaultView;
                dv.RowFilter = string.Format("[Code Locataire] LIKE '%{0}%'", textCodeLocataire.Text.Trim());

                if (dv.Count > 0)
                {
                    // Match found, show error message and change background color to red
                    textShowError.Text = "Code Locataire déjà existe !";
                    textCodeLocataire.BackColor = Color.Red; // Set the background color to red
                }
                else
                {
                    // No match found, clear the error message and reset the background color
                    textShowError.Text = string.Empty;
                    textCodeLocataire.BackColor = SystemColors.Window; // Default background color
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur lors de la recherche : " + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                if (RechtxtRetard.Text != "")
                {
                    DataView dv = _dataTableRetard.DefaultView;
                    dv.RowFilter = string.Format("[Retards] >= '{0}'", RechtxtRetard.Text);
                    GridViewLocataireRetard.DataSource = dv;
                }
                else
                {
                    DataView dv = _dataTableRetard.DefaultView;
                    dv.RowFilter = string.Format("", "");
                    GridViewLocataireRetard.DataSource = dv;
                }
            }
            catch
            {

            }
        }

        private void ConvertDataViewToPDF(DataView dataView, string filePath, string imagePath)
        {
            // Create a document
            Document document = new Document(PageSize.A4, 30, 30, 10, 10);
            PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(filePath, FileMode.Create));
            document.Open();

            try
            {
                iTextSharp.text.Font headerFont = FontFactory.GetFont(FontFactory.COURIER_BOLD, 8);
                iTextSharp.text.Font regularFont = FontFactory.GetFont(FontFactory.COURIER, 8);

                // Load and place the image in the top-right corner
                iTextSharp.text.Image logo = iTextSharp.text.Image.GetInstance(imagePath);
                logo.ScaleToFit(80, 50); // Adjust the size of the image
                logo.SetAbsolutePosition(document.PageSize.Width - logo.ScaledWidth -500, document.PageSize.Height - logo.ScaledHeight - 10);
                document.Add(logo);

                // Header
                Paragraph header = new Paragraph("AGI - agence de gestion immobilière - SARL au capital de 240.000 dhs\n" +
                                                 "9, rue Oued Zam Casablanca 20.000\n" +
                                                 "Tel : 0522 263924 - 0522 275296 - Fax : 0522 223020\n" +
                                                 "C.N.S.S : 1540500 - R.C : 76.923 - Patente : 35501274\n\n",
                                                 headerFont);
                header.Alignment = Element.ALIGN_CENTER;
                document.Add(header);

                // Date and recipient details
                Paragraph dateAndRecipient = new Paragraph("Casablanca, le " + DateTime.Today.ToShortDateString() + "\n\n"
                                                           , regularFont);
                dateAndRecipient.Alignment = Element.ALIGN_RIGHT;
                document.Add(dateAndRecipient);


                PdfPTable pdfTable = new PdfPTable(dataView.Table.Columns.Count);


                foreach (DataColumn column in dataView.Table.Columns)
                {
                    PdfPCell cell = new PdfPCell(new Phrase(column.ColumnName, FontFactory.GetFont(FontFactory.COURIER_BOLD, 6)));
                    cell.BackgroundColor = BaseColor.LIGHT_GRAY;
                    pdfTable.AddCell(cell);
                }

                foreach (DataRowView rowView in dataView)
                {
                    foreach (var item in rowView.Row.ItemArray)
                    {
                        PdfPCell cell = new PdfPCell(new Phrase(item.ToString(), FontFactory.GetFont(FontFactory.COURIER, 6)));
                        pdfTable.AddCell(cell);
                    }
                }

                document.Add(pdfTable);
                document.Close();
                writer.Close();
            }
            catch (Exception ex)
            {

            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string PathComplet = PathListeRetard + $"/Liste_Retards_{DateTime.Today.ToString("yyyyMMdd")}.pdf";

            ConvertDataViewToPDF(_dataTableRetard.DefaultView , PathComplet, "Logo_AGI_page-0001-removebg-preview.png");

            ShowFile(PathComplet);
        }

        private void ShowFile(string pdfPath)
        {
            try
            {
                WebBrowser WebBrowser1 = new WebBrowser();
                WebBrowser1.Navigate(pdfPath);
            }
            catch
            {

            }
        }

        private void Imprimer_Click(object sender, EventArgs e)
        {
            if (GridViewLocataireRetard.SelectedRows.Count > 0)
            {
                try
                {
                    DataTable dataTable = new DataTable();

                    // Get the selected row
                    DataGridViewRow selectedRow1 = GridViewLocataireRetard.SelectedRows[0];

                    // Get the unique value (e.g., code loc) of the row to delete
                    string codeLoc = selectedRow1.Cells["Code Locataire"].Value.ToString();
                    string NomLoc = selectedRow1.Cells["Nom Locataire"].Value.ToString();
                    string Loyer   = selectedRow1.Cells["Loyer"].Value.ToString();
                    string Retard  = selectedRow1.Cells["Retards"].Value.ToString();

                    // Remove the row from DataGridView
                    GridViewLocataireRetard.Rows.Remove(selectedRow1);

                    // Define the SQL query with a WHERE clause to filter by the selected "code loc"
                    string selectQuery = "SELECT [CODE GER], [CHARGE], [Loyer], [datloc], [nom loc], [TAXE], [adresse], [Champ1] " +
                                         "FROM [loc] WHERE [code loc] = ?";

                    // Check if the connection is not already open
                    if (_connection.State != ConnectionState.Open)
                    {
                        _connection.Open(); // Open the connection
                    }

                    // Use OleDbCommand to fetch the data with parameterized query
                    using (OleDbCommand command = new OleDbCommand(selectQuery, _connection))
                    {
                        // Add parameter to the query to prevent SQL injection
                        command.Parameters.AddWithValue("?", codeLoc);

                        using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                        {
                            // Fill a DataTable with the results
                            DataTable dataTable_1 = new DataTable();

                            adapter.Fill(dataTable_1); // Fill the DataTable

                            if (dataTable_1.Rows.Count > 0)
                            {
                                // Assuming you want to show the first row's data
                                DataRow row = dataTable_1.Rows[0];
                                string CodeGerance = row["CODE GER"].ToString();
                                string Usage = row["Champ1"].ToString();

                                if(Usage == "H")
                                {
                                    Usage = "habitation";
                                }
                                else if (Usage == "C")
                                {
                                    Usage = "Comercial";
                                }
                                else if (Usage == "D")
                                {
                                    Usage = "Dépot";
                                }

                                string adresse = row["adresse"].ToString();
                                string dateLocation = Convert.ToDateTime(row["datloc"].ToString()).ToString("dd/MM/yyyy");


                                dataTable.Columns.Add("Code Locataire", typeof(string));
                                dataTable.Columns.Add("Nom Loctaire", typeof(string));
                                dataTable.Columns.Add("Code Gerance", typeof(string));
                                dataTable.Columns.Add("Usage", typeof(string));
                                dataTable.Columns.Add("adresse", typeof(string));
                                dataTable.Columns.Add("Loyer", typeof(string));
                                dataTable.Columns.Add("Retards", typeof(string));
                                dataTable.Columns.Add("Date Engagement", typeof(string));

                                dataTable.Rows.Add(codeLoc, NomLoc, CodeGerance, Usage, adresse, Loyer, Retard, dateLocation);

                            }
                            else
                            {
                                MessageBox.Show("No data found for the selected locataire.");
                            }
                        }
                    }

                    DataView dataView = new DataView(dataTable);
                    DataRowView selectedRow = dataView[0];

                    string outputPath = PathRetardLocataire+$"/Locataire_{NomLoc}_Retards{DateTime.Today.ToString("yyyyMMdd")}.pdf";
                    ConvertRowToPDF(selectedRow, outputPath, "Logo_AGI_page-0001-removebg-preview.png");

                    ShowFile(outputPath);
                }
                catch (Exception ex)
                {
                    //MessageBox.Show();
                }
                finally
                {
                    if (_connection.State == ConnectionState.Open)
                    {
                        _connection.Close();
                    }
                }
            }
            else
            {
                MessageBox.Show("Veuillez sélectionner une ligne pour générer le fichier.");
            }
        }

        private void ConvertRowToPDF(DataRowView row, string filePath, string imagePath)
        {
            // Create a document
            Document document = new Document(PageSize.A4, 30, 30, 10, 10);
            PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(filePath, FileMode.Create));
            document.Open();

            try
            {
                iTextSharp.text.Font headerFont = FontFactory.GetFont(FontFactory.COURIER_BOLD, 8);
                iTextSharp.text.Font regularFont = FontFactory.GetFont(FontFactory.COURIER, 8);

                // Load and place the image in the top-right corner
                iTextSharp.text.Image logo = iTextSharp.text.Image.GetInstance(imagePath);
                logo.ScaleToFit(80, 50); // Adjust the size of the image
                logo.SetAbsolutePosition(document.PageSize.Width - logo.ScaledWidth - 500, document.PageSize.Height - logo.ScaledHeight - 10);
                document.Add(logo);

                // Header
                Paragraph header = new Paragraph("AGI - agence de gestion immobilière - SARL au capital de 240.000 dhs\n" +
                                                 "9, rue Oued Zam Casablanca 20.000\n" +
                                                 "Tel : 0522 263924 - 0522 275296 - Fax : 0522 223020\n" +
                                                 "C.N.S.S : 1540500 - R.C : 76.923 - Patente : 35501274\n\n",
                                                 headerFont);
                header.Alignment = Element.ALIGN_CENTER;
                document.Add(header);

                // Date and recipient details
                Paragraph dateAndRecipient = new Paragraph("Casablanca, le " + DateTime.Today.ToShortDateString() + "\n\n"
                                                           , regularFont);
                dateAndRecipient.Alignment = Element.ALIGN_RIGHT;
                document.Add(dateAndRecipient);


                PdfPTable pdfTable = new PdfPTable(2); // Key-Value pairs

                // Add each column as a key-value pair
                foreach (DataColumn column in row.DataView.Table.Columns)
                {
                    // Key (Column Name)
                    PdfPCell keyCell = new PdfPCell(new Phrase(column.ColumnName, regularFont));
                    keyCell.BackgroundColor = BaseColor.LIGHT_GRAY;
                    pdfTable.AddCell(keyCell);

                    // Value (Row Value)
                    PdfPCell valueCell = new PdfPCell(new Phrase(row[column.ColumnName].ToString(), regularFont));
                    pdfTable.AddCell(valueCell);
                }

                document.Add(pdfTable);
                document.Close();
                writer.Close();
            }
            catch (Exception ex)
            {

            }
        }

        private void MAJCodeLocataire_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) // Si l'utilisateur appuie sur Entrée
            {
                ChargerDonneesLocataireParCode(); // Appelle votre méthode pour charger les données
            }
        }
        
        private void MAJCodeLocataire_TextChanged(object sender, EventArgs e)
        {

        }
        private void MAJCodeLocataire_Leave(object sender, EventArgs e)
        {
            // Appelé lorsque le contrôle perd le focus
            ChargerDonneesLocataireParCode();
        }
    }
}

