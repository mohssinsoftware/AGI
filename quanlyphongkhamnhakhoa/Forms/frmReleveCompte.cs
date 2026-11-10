
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;

using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms; 

namespace PKNK_CNPM.Forms
{
    public partial class FormReleveCompte : Form
    {   
        private OleDbConnection _connection;
        private OleDbDataAdapter _dataAdapter;
        private DataTable _dataTable;

        public FormReleveCompte()
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

                // Note that Champ1 is included in the SELECT for internal logic, but will not be displayed in the DataGridView
                _dataAdapter = new OleDbDataAdapter("SELECT [quittance] AS [Quittance], [code_ger] AS [Code Gerance]," +
                    "[caution] AS [Caution] FROM Encaissement", _connection);
                OleDbCommandBuilder commandBuilder = new OleDbCommandBuilder(_dataAdapter);

                _dataTable = new DataTable();
                _dataAdapter.Fill(_dataTable);

                // Bind the updated DataTable to the DataGridView
                DataGridViewListGerance.DataSource = _dataTable;

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
                string selectQuery = "SELECT [CODE GER], [nom  ger] FROM [ger]";

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

                    //txtCodeGerance.DisplayMember = "nom  ger"; // Display this column in ComboBox
                    //txtCodeGerance.ValueMember = "CODE GER"; // Use this column as the value
                    //txtCodeGerance.DataSource = dataTable;


                    //MAJCodeGerance.DisplayMember = "nom  ger"; // Display this column in ComboBox
                    //MAJCodeGerance.ValueMember = "CODE GER"; // Use this column as the value
                    //MAJCodeGerance.DataSource = dataTable;
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

        private void btnDelete_Click_________(object sender, EventArgs e)
        {
            if (DataGridViewListGerance.SelectedRows.Count > 0)
            {
                foreach (DataGridViewRow row in DataGridViewListGerance.SelectedRows)
                {
                    DataGridViewListGerance.Rows.Remove(row);
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
            if (DataGridViewListGerance.SelectedRows.Count > 0)
            {
                try
                {
                    // Get the selected row
                    DataGridViewRow selectedRow = DataGridViewListGerance.SelectedRows[0];

                    // Get the unique value (e.g., code loc) of the row to delete
                    string codeLoc = selectedRow.Cells["Code Gerance"].Value.ToString();

                    // Remove the row from DataGridView
                    DataGridViewListGerance.Rows.Remove(selectedRow);

                    // Delete the corresponding row from the database
                    string deleteQuery = "DELETE FROM Gerance WHERE [Code Gerance] = @codeGer";

                    using (OleDbCommand cmd = new OleDbCommand(deleteQuery, _connection))
                    {
                        cmd.Parameters.AddWithValue("@codeGer", codeLoc);

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
                            MessageBox.Show("Erreur lors de la suppression du Gerance.");
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
            
            if (DataGridViewListGerance.SelectedRows.Count > 0)
            {
                DataGridViewRow selectedRow = DataGridViewListGerance.SelectedRows[0];
                //selectedRow.Cells["Code Gerance"].Value = txtCodeGerance.Text;
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
            string MAX_ID_Gerance = "1";

            try {
                string Query_MAX_ID_Gerance = "SELECT MAX([N°]) AS N° FROM Gerance_part";

                // Check if the connection is not already open
                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open(); // Open the connection
                }

                // Use OleDbCommand to fetch the data with parameterized query
                using (OleDbCommand command = new OleDbCommand(Query_MAX_ID_Gerance, _connection))
                {   
                    using (OleDbDataAdapter adapter_ID_Gerance = new OleDbDataAdapter(command))
                    {
                        // Fill a DataTable with the results
                        DataTable dataTable_ID_Gerance = new DataTable();

                        adapter_ID_Gerance.Fill(dataTable_ID_Gerance); // Fill the DataTable

                        if (dataTable_ID_Gerance.Rows.Count > 0)
                        {
                            // Assuming you want to show the first row's data
                            DataRow row = dataTable_ID_Gerance.Rows[0];
                            int Res_Int = Convert.ToInt32(row["N°"].ToString());
                            Res_Int++;
                            MAX_ID_Gerance = Res_Int.ToString();
                        }
                        else
                        {
                            // MessageBox.Show("Impossible de récupérer le MAX N° Gerance !!");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading Gerance data: " + ex.Message);
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
                string insertQuery = "INSERT INTO ger ([CODE GER], [nom ger], [c,i,n], [tel ger]" +
                                     ", [comission], [compte], [banque], [agence]) " +
                                     "VALUES (@codeGer, @nomGer, @cin, @telGer, @comission, @compte," +
                                     "@banque, @agence)";

                using (OleDbCommand cmd = new OleDbCommand(insertQuery, _connection))
                {
                    cmd.Parameters.AddWithValue("@codeGer", txtCodeGerance.Text);
                    cmd.Parameters.AddWithValue("@nomGer", txtNomGerance.Text);
                    cmd.Parameters.AddWithValue("@cin", txtCIN.Text);
                    cmd.Parameters.AddWithValue("@telGer", txtTelephone1.Text);
                    cmd.Parameters.AddWithValue("@comission", txtComission.Text);
                    cmd.Parameters.AddWithValue("@compte", txtCompte.Text);
                    cmd.Parameters.AddWithValue("@banque", txtBanque.Text);
                    cmd.Parameters.AddWithValue("@agence", txtAgence.Text);

                    if (_connection.State != ConnectionState.Open)
                    {
                        _connection.Open(); // Open the connection
                    }
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("        Gérance Bien Ajouté        ");

                    LoadData();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur : " + ex.Message);
            }
            finally
            {
                _connection.Close();
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {   
                string updateQuery = "UPDATE ger SET [nom_coper]=@nomCoper, " +
                                     "[description]=@description, [montant]=@montant" +
                                     "WHERE [Code Gerance] = @codeGer";

                /*
                using (OleDbCommand cmd = new OleDbCommand(updateQuery, _connection))
                {   
                    cmd.Parameters.AddWithValue("@nomCoper", txtNomCoprop.Text);
                    cmd.Parameters.AddWithValue("@description", txtDescGerance.Text);
                    cmd.Parameters.AddWithValue("@montant", txtMontant.Text);

                    //cmd.Parameters.AddWithValue("@N", );
                    cmd.Parameters.AddWithValue("@codeGer", MAJCodeGerance.SelectedValue);

                    if (_connection.State != ConnectionState.Open)
                    {
                        _connection.Open(); // Open the connection
                    }
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Gerance Bien Modifié"); // Updated message

                    LoadData(); // Refresh the DataGridView after updating a record
                }
                */
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur : " + ex.Message);
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
                /*
                // Get the selected value from the ComboBox, assuming the ValueMember is set to "code loc"
                string selectedCodeGer = MAJCodeGerance.SelectedValue?.ToString();

                if (string.IsNullOrEmpty(selectedCodeGer))
                {
                    MessageBox.Show("No valid locataire selected.");
                    return;
                }

                // Define the SQL query with a WHERE clause to filter by the selected "code loc"
                string selectQuery = "SELECT [nom_coper], [description], [montant]" +
                                     "FROM [Gerance_part] WHERE [Code Gerance] = ?";

                // Check if the connection is not already open
                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open(); // Open the connection
                }

                // Use OleDbCommand to fetch the data with parameterized query
                using (OleDbCommand command = new OleDbCommand(selectQuery, _connection))
                {
                    // Add parameter to the query to prevent SQL injection
                    command.Parameters.AddWithValue("?", selectedCodeGer);

                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                    {
                        // Fill a DataTable with the results
                        DataTable dataTable = new DataTable();

                        adapter.Fill(dataTable); // Fill the DataTable

                        if (dataTable.Rows.Count > 0)
                        {
                            // Assuming you want to show the first row's data
                            DataRow row = dataTable.Rows[0];
                            MAJtxtDescGerance.Text = row["description"].ToString();
                            MAJtxtMontant.Text = row["montant"].ToString();
                            MAJtxtNomCoprop.SelectedValue = row["nom_coper"].ToString();
                        }
                        else
                        {
                            MessageBox.Show("No data found for the selected Gerance.");
                        }
                    }
                }
                */
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading Gerance data: " + ex.Message);
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

        private void FormReleveCompte_Load(object sender, EventArgs e)
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
                            //MAJCodeGerance.SelectedValue = row["CODE GER"].ToString();
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
                    dv.RowFilter = string.Format("[Quittance] ='{0}'", txtRecherch.Text);
                    DataGridViewListGerance.DataSource = dv;
                }
                else
                {
                    DataView dv = _dataTable.DefaultView;
                    dv.RowFilter = string.Format("", "");
                    DataGridViewListGerance.DataSource = dv;
                }
            }
            catch (Exception ex)
            {

            }
        }
    }
}
