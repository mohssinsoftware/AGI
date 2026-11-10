
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
    public partial class FormGerance : Form
    {
        
        private OleDbConnection _connection;
        private OleDbDataAdapter _dataAdapter;
        private DataTable _dataTable;
        private OleDbDataAdapter _dataAdapterCoprop;
        private DataTable _dataTableCoprop;

        public FormGerance()
        {
            InitializeComponent();
            this.Icon = new Icon("AGI.ico");
            InitializeDatabase();
            LoadData();
            LoadCoprop();
            LoadGerData();
            FnLoadDebutFin();
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

                _dataAdapter = new OleDbDataAdapter("SELECT [CODE GER] AS [Code Gerance], [nom  ger] AS [Nom Gerance]," +
                    " [adresse ger] AS [Adresse Gerance], [tel ger] AS [Telephone Gerance] FROM ger ORDER BY [CODE GER] ASC", _connection);
                OleDbCommandBuilder commandBuilder = new OleDbCommandBuilder(_dataAdapter);

                _dataTable = new DataTable();
                _dataAdapter.Fill(_dataTable);

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

        private void LoadCoprop()
        {
            try
            {
                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open();
                }

                _dataAdapterCoprop = new OleDbDataAdapter("SELECT [nom_cop] AS [Nom Coproprietaire], [code ger] AS [Code Gerance], " +
                "[adresse1] AS [Adresse]," +
                " [autre_frais] AS [Autre frais] FROM part ORDER BY [code ger] ASC", _connection);
                OleDbCommandBuilder commandBuilder = new OleDbCommandBuilder(_dataAdapterCoprop);

                _dataTableCoprop = new DataTable();
                _dataAdapterCoprop.Fill(_dataTableCoprop);

                dataGridViewCoprop.DataSource = _dataTableCoprop;
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
                string selectQuery = "SELECT [CODE GER], [nom  ger] FROM [ger] order by [CODE GER] ASC";

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
                    emptyRow["nom  ger"] = string.Empty; // or set a placeholder text like "<Select>"
                    dataTable.Rows.InsertAt(emptyRow, 0);

                    MAJtxtCodeGer.DisplayMember = "CODE GER"; // Display this column in ComboBox
                    MAJtxtCodeGer.ValueMember = "CODE GER"; // Use this column as the value
                    MAJtxtCodeGer.DataSource = dataTable;

                    txtCodeGerance.DisplayMember = "CODE GER"; // Display this column in ComboBox
                    txtCodeGerance.ValueMember = "CODE GER"; // Use this column as the value
                    txtCodeGerance.DataSource = dataTable;

                    MAJtxtCodeGerance.DisplayMember = "CODE GER"; // Display this column in ComboBox
                    MAJtxtCodeGerance.ValueMember = "CODE GER"; // Use this column as the value
                    MAJtxtCodeGerance.DataSource = dataTable;
                    
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

        private void FnLoadDebutFin()
        {
            try
            {
                DataTable dataTable = new DataTable();

                dataTable.Columns.Add("MoisPaye");

                for (int i = 1; i <= 12; i++)
                {
                    dataTable.Rows.Add(i);
                }

                DataRow emptyRow = dataTable.NewRow();
                emptyRow["MoisPaye"] = DBNull.Value;
                dataTable.Rows.InsertAt(emptyRow, 0);

                txtGenereReleve.DisplayMember = "MoisPaye";
                txtGenereReleve.ValueMember = "MoisPaye";
                txtGenereReleve.DataSource = dataTable;

                MAJtxtGenereReleve.DisplayMember = "MoisPaye";
                MAJtxtGenereReleve.ValueMember = "MoisPaye";
                MAJtxtGenereReleve.DataSource = dataTable;

            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void LoadNomCoprop()
        {
            try
            {
                string selectedCodeGer = MAJtxtCodeGerance.SelectedValue?.ToString();

                if (string.IsNullOrEmpty(selectedCodeGer))
                {
                    return;
                }

                // Adjust the SQL query to match the exact table and column names
                string selectQuery = "SELECT [nom_cop], [nom_cop] FROM [part] where [code ger] = "+ MAJtxtCodeGerance.SelectedValue + ""+ 
                " order by [nom_cop] ASC";

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
                    emptyRow["nom_cop"] = DBNull.Value; // or string.Empty if you prefer empty strings
                    emptyRow["nom_cop"] = string.Empty; // or set a placeholder text like "<Select>"
                    dataTable.Rows.InsertAt(emptyRow, 0);

                    MAJtxtNomCoprop.DisplayMember = "nom_cop"; // Display this column in ComboBox
                    MAJtxtNomCoprop.ValueMember = "nom_cop"; // Use this column as the value
                    MAJtxtNomCoprop.DataSource = dataTable;
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
                    string deleteQuery = "DELETE FROM ger WHERE [code ger] = @codeGer";

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
                            MessageBox.Show("Gérance "+ codeLoc + " a été supprimé avec succès.");
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
                    MessageBox.Show("Erreur lors de la suppression de la ligne: " + ex.Message);
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
                MessageBox.Show("Veuillez sélectionner une ligne à supprimer !.");
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
                //selectedRow.Cells["code_ger"].Value = txtCodeGerance.Text;
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
            try
            {
                string txtCodeGer_ = String.IsNullOrWhiteSpace(txtCodeGer.Text) ? String.Empty : txtCodeGer.Text;
                string txtNomGerance_ = String.IsNullOrWhiteSpace(txtNomGerance.Text) ? "--" : txtNomGerance.Text;
                string txtCIN_ = String.IsNullOrWhiteSpace(txtCIN.Text) ? "--" : txtCIN.Text;
                string txtCompte_ = String.IsNullOrWhiteSpace(txtCompte.Text) ? String.Empty : txtCompte.Text;
                string textBanque_ = String.IsNullOrWhiteSpace(textBanque.Text) ? String.Empty : textBanque.Text;
                string textAgence_ = String.IsNullOrWhiteSpace(textAgence.Text) ? String.Empty : textAgence.Text;
                string textRapportMen_ = String.IsNullOrWhiteSpace(textRapportMen.Text) ? String.Empty : textRapportMen.Text;
                string textMois_ = String.IsNullOrWhiteSpace(textMois.Text) ? String.Empty : textMois.Text;
                int textDoit_ = String.IsNullOrWhiteSpace(textDoit.Text) ? 0 : Convert.ToInt32(textDoit.Text);
                int textComission_ = String.IsNullOrWhiteSpace(textComission.Text) ? 0 : Convert.ToInt32(textComission.Text);
                int textEncaissement_ = String.IsNullOrWhiteSpace(textDoit.Text) ? 0 : Convert.ToInt32(textDoit.Text);
                int textSolde_ = String.IsNullOrWhiteSpace(textSolde.Text) ? 0 : Convert.ToInt32(textSolde.Text);
                string txtGerAdresse1_ = String.IsNullOrWhiteSpace(txtGerAdresse1.Text) ? String.Empty : txtGerAdresse1.Text;
                string txtGerAdresse2_ = String.IsNullOrWhiteSpace(txtGerAdresse2.Text) ? String.Empty : txtGerAdresse2.Text;
                string txtGerAdresse3_ = String.IsNullOrWhiteSpace(txtGerAdresse3.Text) ? String.Empty : txtGerAdresse3.Text;
                string textEmail1_ = String.IsNullOrWhiteSpace(textEmail1.Text) ? "--" : textEmail1.Text;
                string textEmail2_ = String.IsNullOrWhiteSpace(textEmail2.Text) ? "--" : textEmail2.Text;
                string textEmail3_ = String.IsNullOrWhiteSpace(textEmail3.Text) ? "--" : textEmail3.Text;
                string textEmail4_ = String.IsNullOrWhiteSpace(textEmail4.Text) ? "--" : textEmail4.Text;
                string txtTeleGer1_ = String.IsNullOrWhiteSpace(txtTeleGer1.Text) ? "--" : txtTeleGer1.Text;
                string txtTeleGer2_ = String.IsNullOrWhiteSpace(txtTeleGer2.Text) ? "--" : txtTeleGer2.Text;
                string txtTeleGer3_ = String.IsNullOrWhiteSpace(txtTeleGer3.Text) ? "--" : txtTeleGer3.Text;
                string txtGenereReleve_ = txtGenereReleve.SelectedValue != null ? txtGenereReleve.SelectedValue.ToString() : String.Empty;

                string insertQuery = "INSERT INTO ger ([CODE GER], [nom  ger], [c,i,n], [compte], [banque], [agence]," +
                                                      "[RAPPORTS MENS], [doit], [MOIS], [comission], [encaiss], [solde]," +
                                                      "[adresse1], [adresse2], [adresse3], [m1], [m2], [m3], [m4]," +
                                                      "[tel ger], [m11], [m12], [m5])" +
                    "VALUES (@txtCodeGer, @txtNomGerance, @txtCIN, @txtCompte, @textBanque, @textAgence," +
                           " @textRapportMen, @textDoit, @textMois, @textComission, @textEncaissement, @textSolde," +
                           " @txtGerAdresse1, @txtGerAdresse2, @txtGerAdresse3, @textEmail1, @textEmail2, @textEmail3, @textEmail4," +
                           " @txtTeleGer1, @txtTeleGer2, @txtTeleGer3, @txtGenereReleve)";

                using (OleDbCommand cmd = new OleDbCommand(insertQuery, _connection))
                {   
                    cmd.Parameters.AddWithValue("@txtCodeGer", txtCodeGer_);
                    cmd.Parameters.AddWithValue("@txtNomGerance", txtNomGerance_);
                    cmd.Parameters.AddWithValue("@txtCIN", txtCIN_);
                    cmd.Parameters.AddWithValue("@txtCompte", txtCompte_);
                    cmd.Parameters.AddWithValue("@textBanque", textBanque_);
                    cmd.Parameters.AddWithValue("@textAgence", textAgence_);
                    cmd.Parameters.AddWithValue("@textRapportMen", textRapportMen_);
                    cmd.Parameters.AddWithValue("@textDoit", textDoit_);
                    cmd.Parameters.AddWithValue("@textMois", textMois_);
                    cmd.Parameters.AddWithValue("@textComission", textComission_);
                    cmd.Parameters.AddWithValue("@textEncaissement", textEncaissement_);
                    cmd.Parameters.AddWithValue("@textSolde", textSolde_);
                    cmd.Parameters.AddWithValue("@txtGerAdresse1", txtGerAdresse1_);
                    cmd.Parameters.AddWithValue("@txtGerAdresse2", txtGerAdresse2_);
                    cmd.Parameters.AddWithValue("@txtGerAdresse3", txtGerAdresse3_);
                    cmd.Parameters.AddWithValue("@textEmail1", textEmail1_);
                    cmd.Parameters.AddWithValue("@textEmail2", textEmail2_);
                    cmd.Parameters.AddWithValue("@textEmail3", textEmail3_);
                    cmd.Parameters.AddWithValue("@textEmail4", textEmail4_);
                    cmd.Parameters.AddWithValue("@txtTeleGer1", txtTeleGer1_);
                    cmd.Parameters.AddWithValue("@txtTeleGer2", txtTeleGer2_);
                    cmd.Parameters.AddWithValue("@txtTeleGer3", txtTeleGer3_);
                    cmd.Parameters.AddWithValue("@txtGenereReleve", txtGenereReleve_);

                    if (_connection.State != ConnectionState.Open)
                    {
                        _connection.Open();
                    }
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Gérance Bien Ajouté ");

                    LoadData();
                    Refresh();
                    LoadGerData();
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

        private void Refresh()
        {
            try
            {
                //txtCodeGer.SelectedValue = DBNull.Value;


                txtCodeGer.Text = null;
                txtNomGerance.Text = null;
                txtCIN.Text = null;
                txtCompte.Text = null;
                textBanque.Text = null;
                textAgence.Text = null;
                textRapportMen.Text = null;
                textMois.Text = null;
                textDoit.Text = null;
                textComission.Text = null;
                textSolde.Text = null;
                txtGerAdresse1.Text = null;
                txtGerAdresse2.Text = null;
                txtGerAdresse3.Text = null;
                textEmail1.Text = null;
                textEmail2.Text = null;
                textEmail3.Text = null;
                textEmail4.Text = null;

                txtTeleGer1.Text = null;
                txtTeleGer2.Text = null;
                txtTeleGer3.Text = null;
                txtGenereReleve.Text = null;
            }
            catch (Exception ex)
            {

            }
        }


        private void button5_Click(object sender, EventArgs e)
        {
            try
            {   
                string MAJtxtCodeGer_ = MAJtxtCodeGer.SelectedValue != null ? MAJtxtCodeGer.SelectedValue.ToString() : String.Empty;
                string MAJtxtNomGer_ = String.IsNullOrWhiteSpace(MAJtxtNomGer.Text) ? String.Empty : MAJtxtNomGer.Text;
                string MAJtxtCIN_ = String.IsNullOrWhiteSpace(MAJtxtCIN.Text) ? String.Empty : MAJtxtCIN.Text;
                string MAJtxtCompte_ = String.IsNullOrWhiteSpace(MAJtxtCompte.Text) ? String.Empty : MAJtxtCompte.Text;
                string MAJtxtBanque_ = String.IsNullOrWhiteSpace(MAJtxtBanque.Text) ? String.Empty : MAJtxtBanque.Text;
                string MAJtxtAgence_ = String.IsNullOrWhiteSpace(MAJtxtAgence.Text) ? String.Empty : MAJtxtAgence.Text;
                string MAJtxtRapportMen_ = String.IsNullOrWhiteSpace(MAJtxtRapportMen.Text) ? String.Empty : MAJtxtRapportMen.Text;
                string MAJtxtMois_ = String.IsNullOrWhiteSpace(MAJtxtMois.Text) ? String.Empty : MAJtxtMois.Text;
                int MAJtxtDoit_ = String.IsNullOrWhiteSpace(MAJtxtDoit.Text) ? 0 : Convert.ToInt32(MAJtxtDoit.Text);
                int MAJtxtComission_ = String.IsNullOrWhiteSpace(MAJtxtComission.Text) ? 0 : Convert.ToInt32(MAJtxtComission.Text);
                int MAJtxtEncaissement_ = String.IsNullOrWhiteSpace(MAJtxtEncaissement.Text) ? 0 : Convert.ToInt32(MAJtxtEncaissement.Text);
                int MAJtxtSolde_ = String.IsNullOrWhiteSpace(MAJtxtSolde.Text) ? 0 : Convert.ToInt32(MAJtxtSolde.Text);
                string MAJtxtGerAdresse1_ = String.IsNullOrWhiteSpace(MAJtxtGerAdresse1.Text) ? String.Empty : MAJtxtGerAdresse1.Text;
                string MAJtxtGerAdresse2_ = String.IsNullOrWhiteSpace(MAJtxtGerAdresse2.Text) ? String.Empty : MAJtxtGerAdresse2.Text;
                string MAJtxtGerAdresse3_ = String.IsNullOrWhiteSpace(MAJtxtGerAdresse3.Text) ? String.Empty : MAJtxtGerAdresse3.Text;
                string MAJtxtEmail1_ = String.IsNullOrWhiteSpace(MAJtxtEmail1.Text) ? String.Empty : MAJtxtEmail1.Text;
                string MAJtxtEmail2_ = String.IsNullOrWhiteSpace(MAJtxtEmail2.Text) ? String.Empty : MAJtxtEmail2.Text;
                string MAJtxtEmail3_ = String.IsNullOrWhiteSpace(MAJtxtEmail3.Text) ? String.Empty : MAJtxtEmail3.Text;
                string MAJtxtEmail4_ = String.IsNullOrWhiteSpace(MAJtxtEmail4.Text) ? String.Empty : MAJtxtEmail4.Text;
                string MAJtxtTeleGer1_ = String.IsNullOrWhiteSpace(MAJtxtTeleGer1.Text) ? String.Empty : MAJtxtTeleGer1.Text;
                string MAJtxtTeleGer2_ = String.IsNullOrWhiteSpace(MAJtxtTeleGer2.Text) ? String.Empty : MAJtxtTeleGer2.Text;
                string MAJtxtTeleGer3_ = String.IsNullOrWhiteSpace(MAJtxtTeleGer3.Text) ? String.Empty : MAJtxtTeleGer3.Text;
                string MAJtxtGenereReleve_ = MAJtxtGenereReleve.SelectedValue != null ? MAJtxtGenereReleve.SelectedValue.ToString() : String.Empty;


                string updateQuery = @"UPDATE ger SET  [nom  ger] = @MAJtxtNomGer, [c,i,n] = @MAJtxtCIN, [compte] = @MAJtxtCompte,
                [banque] = @MAJtxtBanque, [agence] = @MAJtxtAgence,  [RAPPORTS MENS] = @MAJtxtRapportMen, [doit] = @MAJtxtDoit,
                [MOIS] = @MAJtxtMois, [comission] = @MAJtxtComission, [encaiss] = @MAJtxtEncaissement, [solde] = @MAJtxtSolde,
                [adresse1] = @MAJtxtGerAdresse1, [adresse2] = @MAJtxtGerAdresse2, [adresse3] = @MAJtxtGerAdresse3,
                [m1] = @MAJtxtEmail1, [m2] = @MAJtxtEmail2, [m3] = @MAJtxtEmail3, [m4] = @MAJtxtEmail4, [tel ger] = @MAJtxtTeleGer1,
                [m11] = @MAJtxtTeleGer2, [m12] = @MAJtxtTeleGer3, [m5] = @MAJtxtGenereReleve  WHERE  [CODE GER] = @MAJtxtCodeGer";

                using (OleDbCommand cmd = new OleDbCommand(updateQuery, _connection))
                {
                    // Add parameters to the command
                    cmd.Parameters.AddWithValue("@MAJtxtNomGer", MAJtxtNomGer_);
                    cmd.Parameters.AddWithValue("@MAJtxtCIN", MAJtxtCIN_);
                    cmd.Parameters.AddWithValue("@MAJtxtCompte", MAJtxtCompte_);
                    cmd.Parameters.AddWithValue("@MAJtxtBanque", MAJtxtBanque_);
                    cmd.Parameters.AddWithValue("@MAJtxtAgence", MAJtxtAgence_);
                    cmd.Parameters.AddWithValue("@MAJtxtRapportMen", MAJtxtRapportMen_);
                    cmd.Parameters.AddWithValue("@MAJtxtDoit", MAJtxtDoit_);
                    cmd.Parameters.AddWithValue("@MAJtxtMois", MAJtxtMois_);
                    cmd.Parameters.AddWithValue("@MAJtxtComission", MAJtxtComission_);
                    cmd.Parameters.AddWithValue("@MAJtxtEncaissement", MAJtxtEncaissement_);
                    cmd.Parameters.AddWithValue("@MAJtxtSolde", MAJtxtSolde_);
                    cmd.Parameters.AddWithValue("@MAJtxtGerAdresse1", MAJtxtGerAdresse1_);
                    cmd.Parameters.AddWithValue("@MAJtxtGerAdresse2", MAJtxtGerAdresse2_);
                    cmd.Parameters.AddWithValue("@MAJtxtGerAdresse3", MAJtxtGerAdresse3_);
                    cmd.Parameters.AddWithValue("@MAJtxtEmail1", MAJtxtEmail1_);
                    cmd.Parameters.AddWithValue("@MAJtxtEmail2", MAJtxtEmail2_);
                    cmd.Parameters.AddWithValue("@MAJtxtEmail3", MAJtxtEmail3_);
                    cmd.Parameters.AddWithValue("@MAJtxtEmail4", MAJtxtEmail4_);
                    cmd.Parameters.AddWithValue("@MAJtxtTeleGer1", MAJtxtTeleGer1_);
                    cmd.Parameters.AddWithValue("@MAJtxtTeleGer2", MAJtxtTeleGer2_);
                    cmd.Parameters.AddWithValue("@MAJtxtTeleGer3", MAJtxtTeleGer3_);
                    cmd.Parameters.AddWithValue("@MAJtxtGenereReleve", MAJtxtGenereReleve_);
                    cmd.Parameters.AddWithValue("@MAJtxtCodeGer", MAJtxtCodeGer_);

                    if (_connection.State != ConnectionState.Open)
                    {
                        _connection.Open();
                    }

                    // Execute the update command
                    int rowsAffected = cmd.ExecuteNonQuery();
                    if (rowsAffected > 0)
                    {
                        MessageBox.Show("Gérance modifié avec succès.");
                        LoadData(); // Refresh the DataGridView
                    }
                    else
                    {
                        MessageBox.Show("Aucune ligne mise à jour. Vérifiez le Code Gérance.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Veuillez verifier les champs de saisie ! : " + ex.Message);
            }
            finally
            {
                // Ensure the connection is closed
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
                string selectedCodeGer = MAJtxtCodeGer.SelectedValue?.ToString();

                if (string.IsNullOrEmpty(selectedCodeGer))
                {
                    return;
                }

                string selectQuery = "SELECT [CODE GER], [nom  ger], [c,i,n], [compte], [banque], [agence]," +
                                                      "[RAPPORTS MENS], [doit], [MOIS], [comission], [encaiss], [solde]," +
                                                      "[adresse1], [adresse2], [adresse3], [m1], [m2], [m3], [m4]," +
                                                      "[tel ger], [m11], [m12], [m5]" +
                                     "FROM [ger] WHERE [CODE GER] = ?";

                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open();
                }

                using (OleDbCommand command = new OleDbCommand(selectQuery, _connection))
                {
                    command.Parameters.AddWithValue("?", selectedCodeGer);

                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                    {
                        DataTable dataTable = new DataTable();

                        adapter.Fill(dataTable);

                        if (dataTable.Rows.Count > 0)
                        {
                            DataRow row = dataTable.Rows[0];
                            
                            MAJtxtNomGer.Text = row["nom  ger"].ToString();
                            MAJtxtCIN.Text = row["c,i,n"].ToString();
                            MAJtxtGerAdresse1.Text = row["adresse1"].ToString();
                            MAJtxtGerAdresse2.Text = row["adresse2"].ToString();
                            MAJtxtGerAdresse3.Text = row["adresse3"].ToString();
                            MAJtxtTeleGer1.Text = row["tel ger"].ToString();
                            MAJtxtTeleGer2.Text = row["m11"].ToString();
                            MAJtxtTeleGer3.Text = row["m12"].ToString();
                            MAJtxtGenereReleve.Text = row["m5"].ToString();
                            MAJtxtCompte.Text = row["compte"].ToString();
                            MAJtxtBanque.Text = row["banque"].ToString();
                            MAJtxtAgence.Text = row["agence"].ToString();
                            MAJtxtRapportMen.Text = row["RAPPORTS MENS"].ToString();
                            MAJtxtDoit.Text = row["doit"].ToString();
                            MAJtxtMois.Text = row["MOIS"].ToString();
                            MAJtxtEmail1.Text = row["m1"].ToString();
                            MAJtxtEmail2.Text = row["m2"].ToString();
                            MAJtxtEmail3.Text = row["m3"].ToString();
                            MAJtxtEmail4.Text = row["m4"].ToString();
                            MAJtxtComission.Text = row["comission"].ToString();
                            MAJtxtEncaissement.Text = row["encaiss"].ToString();
                            MAJtxtSolde.Text = row["solde"].ToString();
                        }
                        else
                        {
                            MessageBox.Show("Aucune Gerance Séléctionné !");
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
                if (_connection.State == ConnectionState.Open)
                {
                    _connection.Close();
                }
            }
        }

        /****************************************************************************************************/

        private void FormGerance_Load(object sender, EventArgs e)
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

        private void MAJCodeGer_SelectedIndexChanged(object sender, EventArgs e)
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

        private void AjouterCoprop_Click(object sender, EventArgs e)
        {
            try
            {   
                string txtCodeGerance_ = txtCodeGerance.SelectedValue != null ? txtCodeGerance.SelectedValue.ToString() : String.Empty;
                string txtCoprop_ = String.IsNullOrWhiteSpace(txtCoprop.Text) ? String.Empty : txtCoprop.Text;
                string txtAdresse1_ = String.IsNullOrWhiteSpace(txtAdresse1.Text) ? String.Empty : txtAdresse1.Text;
                string txtAdresse2_ = String.IsNullOrWhiteSpace(txtAdresse2.Text) ? String.Empty : txtAdresse2.Text;
                string txtAdresse3_ = String.IsNullOrWhiteSpace(txtAdresse3.Text) ? String.Empty : txtAdresse3.Text;
                string txtTelephone_ = String.IsNullOrWhiteSpace(txtTelephone.Text) ? String.Empty : txtTelephone.Text;
                string txtCompte_Coprop_ = String.IsNullOrWhiteSpace(txtCompte_Coprop.Text) ? String.Empty : txtCompte_Coprop.Text;
                string txtBanque_ = String.IsNullOrWhiteSpace(txtBanque.Text) ? String.Empty : txtBanque.Text;
                string txtAgence_ = String.IsNullOrWhiteSpace(txtAgence.Text) ? String.Empty : txtAgence.Text;
                int txtAutreFrais_ = String.IsNullOrWhiteSpace(txtAutreFrais.Text) ? 0 : Convert.ToInt32(txtAutreFrais.Text);
                int txtPorcentagePQRT_ = String.IsNullOrWhiteSpace(txtPorcentagePQRT.Text) ? 0 : Convert.ToInt32(txtPorcentagePQRT.Text);
                int txtPorcentageDEN_ = String.IsNullOrWhiteSpace(txtPorcentageDEN.Text) ? 0 : Convert.ToInt32(txtPorcentageDEN.Text);

                string insertQuery = "INSERT INTO part ([nom_cop], [code ger], [pqrt], [den], [adresse1], [adresse2]" +
                                    ", [adresse3], [tele_cop], [compte], [banque], [agence], [autre_frais]) " +
                                     "VALUES (@NtxtCoprop, @txtCodeGerance, @txtPorcentagePQRT, @txtPorcentageDEN, @txtAdresse1, @txtAdresse2" +
                                     ", @txtAdresse3, @txtTelephone, @txtCompte_Coprop, @txtBanque, @txtAgence, @txtAutreFrais)";

                using (OleDbCommand cmd = new OleDbCommand(insertQuery, _connection))
                {
                    cmd.Parameters.AddWithValue("@txtCoprop", txtCoprop_);
                    cmd.Parameters.AddWithValue("@txtCodeGerance", txtCodeGerance_);
                    cmd.Parameters.AddWithValue("@txtPorcentagePQRT", txtPorcentagePQRT_);
                    cmd.Parameters.AddWithValue("@txtPorcentageDEN", txtPorcentageDEN_);
                    cmd.Parameters.AddWithValue("@txtAdresse1", txtAdresse1_);
                    cmd.Parameters.AddWithValue("@txtAdresse2", txtAdresse2_);
                    cmd.Parameters.AddWithValue("@txtAdresse3", txtAdresse3_);
                    cmd.Parameters.AddWithValue("@txtTelephone", txtTelephone_);
                    cmd.Parameters.AddWithValue("@txtCompte_Coprop", txtCompte_Coprop_);
                    cmd.Parameters.AddWithValue("@txtBanque", txtBanque_);
                    cmd.Parameters.AddWithValue("@txtAgence", txtAgence_);
                    cmd.Parameters.AddWithValue("@txtAutreFrais", txtAutreFrais_);
                    //Add more parameters as needed

                    if (_connection.State != ConnectionState.Open)
                    {
                        _connection.Open(); // Open the connection
                    }
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Copropriétaire "+ txtCoprop.Text + " Bien Ajouté ");

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

        private void ModifierCoprop_Click(object sender, EventArgs e)
        {       
            try 
            {
                string MAJtxtNomCoprop_ = String.IsNullOrWhiteSpace(MAJtxtNomCoprop.Text) ? String.Empty : MAJtxtNomCoprop.Text;
                string MAJtxtAdresse1_ = String.IsNullOrWhiteSpace(MAJtxtAdresse1.Text) ? String.Empty : MAJtxtAdresse1.Text;
                string MAJtxtAdresse2_ = String.IsNullOrWhiteSpace(MAJtxtAdresse2.Text) ? String.Empty : MAJtxtAdresse2.Text;
                string MAJtxtAdresse3_ = String.IsNullOrWhiteSpace(MAJtxtAdresse3.Text) ? String.Empty : MAJtxtAdresse3.Text;
                string MAJtxtTelephone_ = String.IsNullOrWhiteSpace(MAJtxtTelephone.Text) ? String.Empty : MAJtxtTelephone.Text;
                string MAJtxtCompte_Coprop_ = String.IsNullOrWhiteSpace(MAJtxtCompte_Coprop.Text) ? String.Empty : MAJtxtCompte_Coprop.Text;
                string MAJtxtBanque_Coprop_ = String.IsNullOrWhiteSpace(MAJtxtBanque_Coprop.Text) ? String.Empty : MAJtxtBanque_Coprop.Text;
                string MAJtxtAgence_Coprop_ = String.IsNullOrWhiteSpace(MAJtxtAgence_Coprop.Text) ? String.Empty : MAJtxtAgence_Coprop.Text;
                int MAJtxtAutreFrais_ = String.IsNullOrWhiteSpace(MAJtxtAutreFrais.Text) ? 0 : Convert.ToInt32(MAJtxtAutreFrais.Text);
                int MAJtxtPorcentagePQRT_ = String.IsNullOrWhiteSpace(MAJtxtPorcentagePQRT.Text) ? 0 : Convert.ToInt32(MAJtxtPorcentagePQRT.Text);
                int MAJtxtPorcentageDEN_ = String.IsNullOrWhiteSpace(MAJtxtPorcentageDEN.Text) ? 0 : Convert.ToInt32(MAJtxtPorcentageDEN.Text);
                

                string updateQuery = @"UPDATE part  SET  [pqrt] = @MAJtxtPorcentagePQRT, [den] = @MAJtxtPorcentageDEN, 
                [adresse1] = @MAJtxtAdresse1, [adresse2] = @MAJtxtAdresse2, [adresse3] = @MAJtxtAdresse3, 
                [tele_cop] = @MAJtxtTelephone, [compte] = @MAJtxtCompte_Coprop, [banque] = @MAJtxtBanque_Coprop, 
                [agence] = @MAJtxtAgence_Coprop, [autre_frais] = @MAJtxtAutreFrais WHERE [nom_cop] = @MAJtxtNomCoprop";

                using (OleDbCommand cmd = new OleDbCommand(updateQuery, _connection))
                {
                    // Add parameters in the same order as in the query
                    cmd.Parameters.AddWithValue("@MAJtxtPorcentagePQRT", MAJtxtPorcentagePQRT_);
                    cmd.Parameters.AddWithValue("@MAJtxtPorcentageDEN", MAJtxtPorcentageDEN_);
                    cmd.Parameters.AddWithValue("@MAJtxtAdresse1", MAJtxtAdresse1_);
                    cmd.Parameters.AddWithValue("@MAJtxtAdresse2", MAJtxtAdresse2_);
                    cmd.Parameters.AddWithValue("@MAJtxtAdresse3", MAJtxtAdresse3_);
                    cmd.Parameters.AddWithValue("@MAJtxtTelephone", MAJtxtTelephone_);
                    cmd.Parameters.AddWithValue("@MAJtxtCompte_Coprop", MAJtxtCompte_Coprop_);
                    cmd.Parameters.AddWithValue("@MAJtxtBanque_Coprop", MAJtxtBanque_Coprop_);
                    cmd.Parameters.AddWithValue("@MAJtxtAgence_Coprop", MAJtxtAgence_Coprop_);
                    cmd.Parameters.AddWithValue("@MAJtxtAutreFrais", MAJtxtAutreFrais_);
                    cmd.Parameters.AddWithValue("@MAJtxtNomCoprop", MAJtxtNomCoprop_);

                    if (_connection.State != ConnectionState.Open)
                    {
                        _connection.Open();
                    }

                    int rowsAffected = cmd.ExecuteNonQuery();

                    if (rowsAffected > 0)
                    {
                        MessageBox.Show($"Copropriétaire {MAJtxtNomCoprop.SelectedValue} Bien Modifié");
                    }
                    else
                    {
                        MessageBox.Show("Aucune modification effectuée. Vérifiez les conditions de la requête.");
                    }

                    LoadData();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur : " + ex.Message);
            }
            finally
            {
                if (_connection.State == ConnectionState.Open)
                {
                    _connection.Close();
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtRechercheGerance_Copro.Text != "")
                {
                    DataView dv = _dataTableCoprop.DefaultView;
                    dv.RowFilter = string.Format("[Code Gerance] ='{0}'", txtRechercheGerance_Copro.Text);
                    dataGridViewCoprop.DataSource = dv;
                }
                else
                {
                    DataView dv = _dataTableCoprop.DefaultView;
                    dv.RowFilter = string.Format("", "");
                    dataGridViewCoprop.DataSource = dv;
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            DataView dv = _dataTableCoprop.DefaultView;
            dv.RowFilter = string.Format("[Nom Coproprietaire] LIKE '%{0}%'", textNomCprop.Text.Trim());
            dataGridViewCoprop.DataSource = dv;
        }

        private void MAJtxtCodeGerance_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadNomCoprop();
        }

        private void MAJtxtNomCoprop_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {

                // Get the selected value from the ComboBox, assuming the ValueMember is set to "code loc"
                string selectedCodeGer = MAJtxtCodeGer.SelectedValue?.ToString();
                string selectedNomCopro = MAJtxtNomCoprop.SelectedValue?.ToString();

                if (string.IsNullOrEmpty(selectedCodeGer) || string.IsNullOrEmpty(selectedNomCopro))
                {
                    return;
                }

                // Define the SQL query with a WHERE clause to filter by the selected "code loc"
                string selectQuery = "SELECT [pqrt], [den], [adresse1], [adresse2]" +
                ", [adresse3], [tele_cop], [compte], " +
                "  [banque], [agence], [autre_frais] FROM part" +
                "  WHERE [nom_cop] = ?";
                
                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open();
                }

                using (OleDbCommand command = new OleDbCommand(selectQuery, _connection))
                {
                    command.Parameters.AddWithValue("?", selectedNomCopro);

                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                    {
                        DataTable dataTable = new DataTable();

                        adapter.Fill(dataTable);

                        if (dataTable.Rows.Count > 0)
                        {
                            DataRow row = dataTable.Rows[0];

                            MAJtxtPorcentagePQRT.Text = row["pqrt"].ToString();
                            MAJtxtPorcentageDEN.Text = row["den"].ToString();
                            //cmd.Parameters.AddWithValue("@datloc", DateTime.Parse(textDateLoyer.Text)); // Make sure to handle dates correctly
                            MAJtxtAdresse1.Text = row["adresse1"].ToString();
                            MAJtxtAdresse2.Text = row["adresse2"].ToString();
                            MAJtxtAdresse3.Text = row["adresse3"].ToString();
                            MAJtxtTelephone.Text = row["tele_cop"].ToString();
                            MAJtxtCompte_Coprop.Text = row["compte"].ToString();
                            MAJtxtBanque_Coprop.Text = row["banque"].ToString();
                            MAJtxtAgence_Coprop.Text = row["agence"].ToString();
                            MAJtxtAutreFrais.Text = row["autre_frais"].ToString();

                        }
                        else
                        {
                            MessageBox.Show("Aucune copropriétaire Séléctionné !");
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading copropriétaire data: " + ex.Message);
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

        private void button4_Click(object sender, EventArgs e)
        {
            if (dataGridViewCoprop.SelectedRows.Count > 0)
            {
                try
                {
                    // Get the selected row
                    DataGridViewRow selectedRow = dataGridViewCoprop.SelectedRows[0];

                    // Get the unique value (e.g., code loc) of the row to delete
                    string codeLoc = selectedRow.Cells["Nom Coproprietaire"].Value.ToString();

                    // Remove the row from DataGridView
                    dataGridViewCoprop.Rows.Remove(selectedRow);

                    // Delete the corresponding row from the database
                    string deleteQuery = "DELETE FROM part WHERE [nom_cop] = @codeCop";

                    using (OleDbCommand cmd = new OleDbCommand(deleteQuery, _connection))
                    {
                        cmd.Parameters.AddWithValue("@codeCop", codeLoc);

                        if (_connection.State != ConnectionState.Open)
                        {
                            _connection.Open();
                        }

                        int rowsAffected = cmd.ExecuteNonQuery();
                        if (rowsAffected > 0)
                        {
                            MessageBox.Show("Le Coproprietaire a été supprimé avec succès.");
                        }
                        else
                        {
                            MessageBox.Show("Erreur lors de la suppression du Coproprietaire.");
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

        private void txtCodeGer_TextChanged(object sender, EventArgs e)
        {
            try
            {
                // Check if the input value is null or empty
                if (string.IsNullOrWhiteSpace(txtCodeGer.Text))
                {
                    // Clear the error message and reset the background color
                    textShowError.Text = string.Empty;
                    txtCodeGer.BackColor = SystemColors.Window; // Default background color
                    _dataTable.DefaultView.RowFilter = string.Empty;
                    DataGridViewListGerance.DataSource = _dataTable.DefaultView;
                    return; // Exit the method
                }

                // Filter the DataView based on the search term
                DataView dv = _dataTable.DefaultView;
                dv.RowFilter = string.Format("CONVERT([Code Gerance], 'System.String') LIKE '%{0}%'", txtCodeGer.Text.Trim());

                if (dv.Count > 0)
                {
                    textShowError.Text = "Code Gerance déjà existe !";
                    txtCodeGer.BackColor = Color.Red;
                }
                else
                {
                    textShowError.Text = string.Empty;
                    txtCodeGer.BackColor = SystemColors.Window;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur lors de la recherche : " + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
