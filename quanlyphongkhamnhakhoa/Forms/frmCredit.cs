
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
    public partial class FormCredit : Form
    {
        
        private OleDbConnection _connection;
        private OleDbDataAdapter _dataAdapter;
        private DataTable _dataTable;

        public FormCredit()
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
                
                _dataAdapter = new OleDbDataAdapter("SELECT [N°] AS [Numero Credit],[code_ger] AS [Code Gerance]," +
                    "[nom_coper] AS [Nom Coproprietaire],[description] AS [Description],[montant] AS [Montant]" +
                    " FROM credit", _connection);
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

                GridViewCredit.DataSource = _dataTable;

                //DEBUT : Ajustement automatique pour les colonnes restantes-------------------------------------------------
                GridViewCredit.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                // Ajustement spécifique pour la colonne "Nom Locataire"
                GridViewCredit.Columns["Nom Coproprietaire"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; // Désactiver l'ajustement automatique
                GridViewCredit.Columns["Nom Coproprietaire"].Width = 400; // Définir la largeur souhaitée
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
                string selectQuery = "SELECT [CODE GER], [nom  ger] FROM [ger] ORDER BY [CODE GER]";

                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open();
                }

                using (OleDbDataAdapter adapter = new OleDbDataAdapter(selectQuery, _connection))
                {
                    DataTable dataTable = new DataTable();
                    adapter.Fill(dataTable);

                    DataRow emptyRow = dataTable.NewRow();
                    emptyRow["CODE GER"] = DBNull.Value; // or string.Empty if you prefer empty strings
                    dataTable.Rows.InsertAt(emptyRow, 0);

                    txtCodeGerance.DisplayMember = "CODE GER";
                    txtCodeGerance.ValueMember = "CODE GER";
                    txtCodeGerance.DataSource = dataTable;

                    MAJCodeGerance.DisplayMember = "CODE GER";
                    MAJCodeGerance.ValueMember = "CODE GER";
                    MAJCodeGerance.DataSource = dataTable;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading ger data: " + ex.Message + "\n\n" + ex.StackTrace);
            }
            finally
            {
                if (_connection.State == ConnectionState.Open)
                {
                    _connection.Close();
                }
            }
        }

        private void btnDelete_Click_________(object sender, EventArgs e)
        {
            if (GridViewCredit.SelectedRows.Count > 0)
            {
                foreach (DataGridViewRow row in GridViewCredit.SelectedRows)
                {
                    GridViewCredit.Rows.Remove(row);
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
            if (GridViewCredit.SelectedRows.Count > 0)
            {
                try
                {
                    // Get the selected row
                    DataGridViewRow selectedRow = GridViewCredit.SelectedRows[0];

                    // Get the unique value (e.g., code loc) of the row to delete
                    string NumeroCredit = selectedRow.Cells["Numero Credit"].Value.ToString();

                    // Remove the row from DataGridView
                    GridViewCredit.Rows.Remove(selectedRow);

                    // Delete the corresponding row from the database
                    string deleteQuery = "DELETE FROM credit WHERE [N°] = @NumeroCredit";

                    using (OleDbCommand cmd = new OleDbCommand(deleteQuery, _connection))
                    {
                        cmd.Parameters.AddWithValue("@NumeroCredit", NumeroCredit);

                        if (_connection.State != ConnectionState.Open)
                        {
                            _connection.Open();
                        }

                        int rowsAffected = cmd.ExecuteNonQuery();
                        if (rowsAffected > 0)
                        {
                            MessageBox.Show("Le créedit a été supprimé avec succès.");
                        }
                        else
                        {
                            MessageBox.Show("Erreur lors de la suppression du crédit.");
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
            
            if (GridViewCredit.SelectedRows.Count > 0)
            {
                DataGridViewRow selectedRow = GridViewCredit.SelectedRows[0];
                selectedRow.Cells["code_ger"].Value = txtCodeGerance.Text;
                //selectedRow.Cells["CODE GER"].Value = txtLoyer.Text;
                // Update other fields as necessary

                SaveChanges();
            }
            else
            {
                MessageBox.Show("Veuillez sélectionner une ligne à mettre à jour.");
            }

            
            //BindGrid(khachHangService.GetAll());
            //clearValue();
        }

        private void btnCanLamSan_Click_1(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtMontant.Text))
            {
                MessageBox.Show("Veillez saisie un montant Crédit.");
                return;
            }

            string MAX_ID_Credit = "1";

            try {
                string Query_MAX_ID_Credit = "SELECT MAX([N°]) AS N° FROM credit";

                // Check if the connection is not already open
                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open(); // Open the connection
                }

                // Use OleDbCommand to fetch the data with parameterized query
                using (OleDbCommand command = new OleDbCommand(Query_MAX_ID_Credit, _connection))
                {   
                    using (OleDbDataAdapter adapter_ID_Credit = new OleDbDataAdapter(command))
                    {
                        // Fill a DataTable with the results
                        DataTable dataTable_ID_Credit = new DataTable();

                        adapter_ID_Credit.Fill(dataTable_ID_Credit); // Fill the DataTable

                        if (dataTable_ID_Credit.Rows.Count > 0)
                        {
                            // Assuming you want to show the first row's data
                            DataRow row = dataTable_ID_Credit.Rows[0];
                            int Res_Int = Convert.ToInt32(row["N°"].ToString());
                            Res_Int++;
                            MAX_ID_Credit = Res_Int.ToString();
                        }
                        else
                        {
                            // MessageBox.Show("Impossible de récupérer le MAX N° Crédit !!");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("La table des Crédit dans la base de données est vide, le numéro de Crédit sera être 1 " + ex.Message);
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
                string txtDescCredit_ = String.IsNullOrWhiteSpace(txtDescCredit.Text) ? String.Empty : txtDescCredit.Text;

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

                string insertQuery = "INSERT INTO credit ([N°], [code_ger], [nom_coper], [description], [montant]) " +
                                     "VALUES (@N, @codeGer, @nomCoper, @description, @montant)";

                using (OleDbCommand cmd = new OleDbCommand(insertQuery, _connection))
                {
                    cmd.Parameters.AddWithValue("@N", MAX_ID_Credit);
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
                    MessageBox.Show("Credit Bien Ajouté ");

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
                MessageBox.Show("Veillez saisie un montant Crédit.");
                return;
            }

            try
            {
                string MAJNumeroFrais_ = MAJNumeroFrais.SelectedValue != null ? MAJNumeroFrais.SelectedValue.ToString() : String.Empty;
                string MAJtxtNomCoprop_ = MAJtxtNomCoprop.SelectedValue != null ? MAJtxtNomCoprop.SelectedValue.ToString() : String.Empty;
                string MAJtxtDescCredit_ = String.IsNullOrWhiteSpace(MAJtxtDescCredit.Text) ? String.Empty : MAJtxtDescCredit.Text;

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

                string updateQuery = "UPDATE credit SET [nom_coper]=@nomCoper, " +
                                     "[description]=@description, [montant]=@montant" +
                                     " WHERE [N°] = @NumeroCredit";

                using (OleDbCommand cmd = new OleDbCommand(updateQuery, _connection))
                {   
                    cmd.Parameters.AddWithValue("@nomCoper", MAJtxtNomCoprop_);
                    cmd.Parameters.AddWithValue("@description", MAJtxtDescCredit_);
                    cmd.Parameters.AddWithValue("@montant", MAJtxtMontant_);
                    cmd.Parameters.AddWithValue("@NumeroCredit", MAJNumeroFrais_);

                    if (_connection.State != ConnectionState.Open)
                    {
                        _connection.Open();
                    }
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Crédit Bien Modifié");

                    LoadData();
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
                string selectQuery = "SELECT [N°] FROM [credit] WHERE [code_ger] = @CodeGer ORDER BY [N°]";

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
                            MAJNumeroFrais.DisplayMember = "N°"; // Column to display
                            MAJNumeroFrais.ValueMember = "N°";           // Column to use as value
                            MAJNumeroFrais.DataSource = dataTable;

                            MAJNumeroFrais.SelectedValue = dataTable.Rows[0]["N°"];
                        }
                        else
                        {
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

        /****************************************************************************************************/

        private void FormCredit_Load(object sender, EventArgs e)
        {
            LoadTheme();
            setCount();
        }
        private void LoadTheme()
        {
            
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
            
        }

        private void MAJNomLocataire_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                // Get the selected value from the ComboBox, assuming the ValueMember is set to "code loc"
                string selectedCodeLoc = "";//MAJNomLocataire.SelectedValue?.ToString();

                if (string.IsNullOrEmpty(selectedCodeLoc))
                {
                    MessageBox.Show("No valid locataire selected.");
                    return;
                }

                // Define the SQL query with a WHERE clause to filter by the selected "code loc"
                string selectQuery = "SELECT [code loc], [CODE GER], [CHARGE], [total], [datloc], [nom loc], [TAXE], [adresse1] " +
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
                            // Assuming you want to show the first row's data
                            DataRow row = dataTable.Rows[0];
                            //MAJCodeLocataire.Text = row["code loc"].ToString();
                            //MAJCharge.Text = row["CHARGE"].ToString();
                            //MAJTotal.Text = row["total"].ToString();
                            //MAJDateLoyer.Text = row["datloc"].ToString();
                            //MAJTaxt.Text = row["TAXE"].ToString();
                            //MAJAdresse.Text = row["adresse1"].ToString();
                            MAJCodeGerance.SelectedValue = row["CODE GER"].ToString();
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

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                if(txtRecherch.Text != "")
                {
                    DataView dv = _dataTable.DefaultView;
                    dv.RowFilter = string.Format("[Code Gerance] ='{0}'", txtRecherch.Text);
                    GridViewCredit.DataSource = dv;
                }
                else
                {
                    DataView dv = _dataTable.DefaultView;
                    dv.RowFilter = string.Format("", "");
                    GridViewCredit.DataSource = dv;
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
                    // Add parameter value
                    command.Parameters.AddWithValue("@CodeGer", selectedCodeGer);

                    // Use OleDbDataAdapter with the properly configured OleDbCommand
                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                    {
                        DataTable _dataTable_Coprop = new DataTable();
                        adapter.Fill(_dataTable_Coprop);

                        // Insert an empty row for display purposes
                        DataRow emptyRow = _dataTable_Coprop.NewRow();
                        emptyRow["nom_cop"] = DBNull.Value;
                        _dataTable_Coprop.Rows.InsertAt(emptyRow, 0);

                        // Bind data to txtNomCoprop
                        txtNomCoprop.DisplayMember = "nom_cop";
                        txtNomCoprop.ValueMember = "nom_cop";
                        txtNomCoprop.DataSource = _dataTable_Coprop;

                        // Bind data to MAJtxtNomCoprop
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

        private void MAJNumeroFrais_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                string NumeroFrais = MAJNumeroFrais.SelectedValue?.ToString();

                if (string.IsNullOrEmpty(NumeroFrais))
                {
                    //MessageBox.Show("No valid locataire selected.");
                    return;
                }

                string selectQuery = "SELECT [nom_coper], [description], [montant]" +
                                     "FROM [credit] WHERE [N°] = @NumeroFrais";

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
                            MAJtxtDescCredit.Text = row["description"].ToString();
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
