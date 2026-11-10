
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;

using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms; 

namespace PKNK_CNPM.Forms
{
    public partial class FormFrais : Form
    {
        private OleDbConnection _connection;
        private OleDbDataAdapter _dataAdapter;
        private DataTable _dataTable;

        public FormFrais()
        {
            InitializeComponent();
            this.Icon = new Icon("AGI.ico");
            InitializeDatabase();
            LoadData();
            LoadGerData();
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

                _dataAdapter = new OleDbDataAdapter("SELECT [N°] AS [Numero Frais],[code_ger] AS [Code Gerance]," +
                "[nom_coper] AS [Nom Coproprietaire], [description] AS [Description], [montant] AS [Montant]" +
                "  FROM frais_part", _connection);
                OleDbCommandBuilder commandBuilder = new OleDbCommandBuilder(_dataAdapter);

                _dataTable = new DataTable();
                _dataAdapter.Fill(_dataTable);
                
                if (_dataTable.Columns.Contains("Montant"))
                {
                    // Add a new column for the formatted Montant
                    _dataTable.Columns.Add("Formatted Montant", typeof(string));

                    foreach (DataRow row in _dataTable.Rows)
                    {
                        if (double.TryParse(row["Montant"]?.ToString(), out double Montant))
                        {
                            row["Formatted Montant"] = $"{Montant:0.00} DH"; // Format with 2 decimal places
                        }
                        else
                        {
                            row["Formatted Montant"] = "Invalid Value"; // Handle non-numeric cases
                        }
                    }

                    // Optional: Remove the original 'Montant' column if not needed
                    _dataTable.Columns.Remove("Montant");
                    _dataTable.Columns["Formatted Montant"].ColumnName = "Montant";
                }

                DataGridViewListfrais.DataSource = _dataTable;

                //DEBUT : Ajustement automatique pour les colonnes restantes-------------------------------------------------
                DataGridViewListfrais.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                // Ajustement spécifique pour la colonne "Nom Locataire"
                DataGridViewListfrais.Columns["Nom Coproprietaire"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; // Désactiver l'ajustement automatique
                DataGridViewListfrais.Columns["Nom Coproprietaire"].Width = 400; // Définir la largeur souhaitée
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

        private void LoadGerData()
        {
            try
            {
                // Adjust the SQL query to match the exact table and column names
                string selectQuery = "SELECT [CODE GER] FROM [ger] ORDER BY [CODE GER]";

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

        private void btnDelete_Click_________(object sender, EventArgs e)
        {
            if (DataGridViewListfrais.SelectedRows.Count > 0)
            {
                foreach (DataGridViewRow row in DataGridViewListfrais.SelectedRows)
                {
                    DataGridViewListfrais.Rows.Remove(row);
                }
                SaveChanges();
            }
            else
            {
                MessageBox.Show("Please select a row to delete.");
            }
            
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (DataGridViewListfrais.SelectedRows.Count > 0)
            {
                try
                {
                    // Get the selected row
                    DataGridViewRow selectedRow = DataGridViewListfrais.SelectedRows[0];

                    // Get the unique value (e.g., code loc) of the row to delete
                    string NumeroFrais = selectedRow.Cells["Numero Frais"].Value.ToString();

                    // Remove the row from DataGridView
                    DataGridViewListfrais.Rows.Remove(selectedRow);

                    // Delete the corresponding row from the database
                    string deleteQuery = "DELETE FROM frais_part WHERE [N°] = @NumeroFrais";

                    using (OleDbCommand cmd = new OleDbCommand(deleteQuery, _connection))
                    {
                        cmd.Parameters.AddWithValue("@NumeroFrais", NumeroFrais);

                        if (_connection.State != ConnectionState.Open)
                        {
                            _connection.Open();
                        }

                        int rowsAffected = cmd.ExecuteNonQuery();
                        if (rowsAffected > 0)
                        {
                            MessageBox.Show("Frai a été supprimé avec succès.");
                        }
                        else
                        {
                            MessageBox.Show("Erreur lors de la suppression du Frais.");
                        }
                    }

                    // Refresh the DataGridView after deletion
                    LoadData();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error deleting row: " + ex.Message);
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
                MessageBox.Show("Please select a row to delete.");
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
            
            if (DataGridViewListfrais.SelectedRows.Count > 0)
            {
                DataGridViewRow selectedRow = DataGridViewListfrais.SelectedRows[0];
                selectedRow.Cells["code_ger"].Value = txtCodeGerance.Text;
                //selectedRow.Cells["CODE GER"].Value = txtLoyer.Text;
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
            
            if (string.IsNullOrEmpty(txtMontant.Text))
            {
                MessageBox.Show("Veillez saisie un montant frais.");
                return;
            }

            string MAX_ID_frais = "1";

            try {
                string Query_MAX_ID_frais = "SELECT MAX([N°]) AS N° FROM frais_part";

                // Check if the connection is not already open
                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open(); // Open the connection
                }

                // Use OleDbCommand to fetch the data with parameterized query
                using (OleDbCommand command = new OleDbCommand(Query_MAX_ID_frais, _connection))
                {   
                    using (OleDbDataAdapter adapter_ID_frais = new OleDbDataAdapter(command))
                    {
                        // Fill a DataTable with the results
                        DataTable dataTable_ID_frais = new DataTable();

                        adapter_ID_frais.Fill(dataTable_ID_frais); // Fill the DataTable

                        if (dataTable_ID_frais.Rows.Count > 0)
                        {
                            // Assuming you want to show the first row's data
                            DataRow row = dataTable_ID_frais.Rows[0];
                            int Res_Int = Convert.ToInt32(row["N°"].ToString());
                            Res_Int++;
                            MAX_ID_frais = Res_Int.ToString();
                        }
                        else
                        {
                            // MessageBox.Show("Impossible de récupérer le MAX N° Frais !!");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("La table des Frais dans la base de données est vide, le numéro de Frais sera être 1 " + ex.Message);
            }
            finally
            {
                // Ensure the connection is closed even if an error occurs
                if (_connection.State == ConnectionState.Open)
                {
                    _connection.Close();
                }
            }

            try
            {
                string txtNomCoprop_ = txtNomCoprop.SelectedValue != null ? txtNomCoprop.SelectedValue.ToString() : String.Empty;
                string txtCodeGerance_ = txtCodeGerance.SelectedValue != null ? txtCodeGerance.SelectedValue.ToString() : String.Empty;
                string txtDescCredit_ = String.IsNullOrWhiteSpace(txtDescfrais.Text) ? String.Empty : txtDescfrais.Text;
                float txtMontant_ = 0;
                if (!String.IsNullOrWhiteSpace(txtMontant.Text))
                {
                    string formattedText = txtMontant.Text.Replace(',', '.');
                    if (!float.TryParse(formattedText, NumberStyles.Float, CultureInfo.InvariantCulture, out txtMontant_))
                    {
                        MessageBox.Show("Veuillez entrer un montant valide !");
                        return;
                    }
                }

                string insertQuery = "INSERT INTO frais_part ([N°], [code_ger], [nom_coper], [description], [montant]) " +
                                     "VALUES (@N, @codeGer, @nomCoper, @description, @montant)";

                using (OleDbCommand cmd = new OleDbCommand(insertQuery, _connection))
                {
                    cmd.Parameters.AddWithValue("@N", MAX_ID_frais);
                    cmd.Parameters.AddWithValue("@codeGer", txtCodeGerance_);
                    cmd.Parameters.AddWithValue("@nomCoper", txtNomCoprop_);
                    cmd.Parameters.AddWithValue("@description", txtDescCredit_);
                    cmd.Parameters.AddWithValue("@montant", txtMontant_);
                    //Add more parameters as needed

                    if (_connection.State != ConnectionState.Open)
                    {
                        _connection.Open(); // Open the connection
                    }
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("frai Bien Ajouté ");

                    LoadData(); // Refresh the DataGridView after adding a new record
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Veuillez verifier les champs de saisie ! : " + ex.Message);
            }
            finally
            {
                _connection.Close();
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(MAJtxtMontant.Text))
            {
                MessageBox.Show("Veillez saisie un montant frais.");
                return;
            }

            try
            {
                string MAJtxtNomCoprop_ = MAJtxtNomCoprop.SelectedValue != null ? MAJtxtNomCoprop.SelectedValue.ToString() : String.Empty;
                string MAJNumeroFrais_ = MAJNumeroFrais.SelectedValue != null ? MAJNumeroFrais.SelectedValue.ToString() : String.Empty;
                string txtCodeGerance_ = MAJCodeGerance.SelectedValue != null ? MAJCodeGerance.SelectedValue.ToString() : String.Empty;
                string MAJtxtDescfrais_ = String.IsNullOrWhiteSpace(MAJtxtDescfrais.Text) ? String.Empty : MAJtxtDescfrais.Text;
                float MAJtxtMontant_ = 0;
                if (!String.IsNullOrWhiteSpace(MAJtxtMontant.Text))
                {
                    string formattedText = MAJtxtMontant.Text.Replace(',', '.');
                    if (!float.TryParse(formattedText, NumberStyles.Float, CultureInfo.InvariantCulture, out MAJtxtMontant_))
                    {
                        MessageBox.Show("Veuillez entrer un montant valide !");
                        return;
                    }
                }

                string updateQuery = "UPDATE frais_part SET [nom_coper]=@nomCoper, " +
                                     "[description]=@description, [montant]=@montant " +
                                     "WHERE [N°] = @NumeroFrais";

                using (OleDbCommand cmd = new OleDbCommand(updateQuery, _connection))
                {   
                    cmd.Parameters.AddWithValue("@nomCoper", MAJtxtNomCoprop_);
                    cmd.Parameters.AddWithValue("@description", MAJtxtDescfrais_);
                    cmd.Parameters.AddWithValue("@montant", MAJtxtMontant_);
                    cmd.Parameters.AddWithValue("@NumeroFrais", MAJNumeroFrais_);
                    //cmd.Parameters.AddWithValue("@N", );
                    cmd.Parameters.AddWithValue("@codeGer", txtCodeGerance_);

                    if (_connection.State != ConnectionState.Open)
                    {
                        _connection.Open(); // Open the connection
                    }
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Frais Bien Modifié"); // Updated message

                    LoadData(); // Refresh the DataGridView after updating a record
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Veuillez verifier les champs de saisie ! : " + ex.Message);
            }
            finally
            {
                _connection.Close();
            }
        }

        /****************************************************************************************************/

        private void FormFrais_Load(object sender, EventArgs e)
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


        private void panel7_Paint(object sender, PaintEventArgs e)
        {

        }

        private void txtRecherch_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                if(txtRecherch.Text != "")
                {
                    DataView dv = _dataTable.DefaultView;
                    dv.RowFilter = string.Format("[Code Gerance] ='{0}'", txtRecherch.Text);
                    DataGridViewListfrais.DataSource = dv;
                }
                else
                {
                    DataView dv = _dataTable.DefaultView;
                    dv.RowFilter = string.Format("", "");
                    DataGridViewListfrais.DataSource = dv;
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void txtCodeGerance_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                string selectedCodeGer = txtCodeGerance.SelectedValue?.ToString();

                if (string.IsNullOrEmpty(selectedCodeGer))
                {
                    // If no valid selection is made, exit early
                    return;
                }

                string selectQuery = "SELECT [nom_cop] FROM [part] WHERE [code ger] = @CodeGer ORDER BY [nom_cop]";

                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open();
                }

                using (OleDbCommand command = new OleDbCommand(selectQuery, _connection))
                {
                    command.Parameters.AddWithValue("@CodeGer", selectedCodeGer);

                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                    {
                        DataTable _dataTable_Coprop = new DataTable();
                        adapter.Fill(_dataTable_Coprop);

                        DataRow emptyRow = _dataTable_Coprop.NewRow();
                        emptyRow["nom_cop"] = DBNull.Value;
                        _dataTable_Coprop.Rows.InsertAt(emptyRow, 0);

                        txtNomCoprop.DisplayMember = "nom_cop";
                        txtNomCoprop.ValueMember = "nom_cop";
                        txtNomCoprop.DataSource = _dataTable_Coprop;

                        MAJtxtNomCoprop.DisplayMember = "nom_cop";
                        MAJtxtNomCoprop.ValueMember = "nom_cop";
                        MAJtxtNomCoprop.DataSource = _dataTable_Coprop;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading crédit data: " + ex.Message);
            }
            finally
            {
                if (_connection.State == ConnectionState.Open)
                {
                    _connection.Close();
                }
            }

        }

        private void MAJCodeGerance_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                string selectedCodeGer = MAJCodeGerance.SelectedValue?.ToString();

                if (string.IsNullOrEmpty(selectedCodeGer))
                {
                    // If no valid selection is made, exit early
                    return;
                }

                // Correct the query syntax
                string selectQuery = "SELECT [N°] FROM [frais_part] WHERE [code_ger] = @CodeGer ORDER BY [N°]";

                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open();
                }

                using (OleDbCommand command = new OleDbCommand(selectQuery, _connection))
                {
                    // Add parameter value for the query
                    command.Parameters.AddWithValue("@CodeGer", selectedCodeGer);

                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                    {
                        DataTable dataTable = new DataTable();
                        adapter.Fill(dataTable);

                        if (dataTable.Rows.Count > 0)
                        {
                            // Bind data to MAJNumeroFrais ComboBox
                            MAJNumeroFrais.DisplayMember = "N°"; // Column to display
                            MAJNumeroFrais.ValueMember = "N°";           // Column to use as value
                            MAJNumeroFrais.DataSource = dataTable;

                            // Optionally set the selected value to the first row
                            MAJNumeroFrais.SelectedValue = dataTable.Rows[0]["N°"];
                        }
                        else
                        {
                            // Clear ComboBox if no data is found
                            MAJNumeroFrais.DataSource = null;
                            MAJNumeroFrais.Items.Clear();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading data: " + ex.Message);
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

        private void MAJNumeroFrais_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                string NumeroFrais = MAJCodeGerance.SelectedValue?.ToString();

                if (string.IsNullOrEmpty(NumeroFrais))
                {
                    //MessageBox.Show("No valid locataire selected.");
                    return;
                }

                string selectQuery = "SELECT [nom_coper], [description], [montant]" +
                                     "FROM [frais_part] WHERE [N°] = @NumeroFrais";

                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open();
                }

                using (OleDbCommand command = new OleDbCommand(selectQuery, _connection))
                {
                    command.Parameters.AddWithValue("@NumeroFrais", NumeroFrais);

                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                    {
                        DataTable dataTable = new DataTable();

                        adapter.Fill(dataTable);

                        if (dataTable.Rows.Count > 0)
                        {
                            DataRow row = dataTable.Rows[0];
                            MAJtxtDescfrais.Text = row["description"].ToString();
                            MAJtxtMontant.Text = row["montant"].ToString();
                            MAJtxtNomCoprop.SelectedValue = row["nom_coper"].ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading crédit data: " + ex.Message);
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
    }
}
