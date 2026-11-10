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
using Rectangle = iTextSharp.text.Rectangle;

namespace PKNK_CNPM.Forms
{
    public partial class FormEncaissement : Form
    {   
        private OleDbConnection _connection;
        private OleDbDataAdapter _dataAdapter;
        private DataTable _dataTable;
        private DataTable _dataTable_Encaissement;
        private DataTable _dataTableValideEncaissement;
        private string QuittancePath = String.Empty;
        private string EtatDetaillePath = String.Empty;

        public FormEncaissement()
        {
            
            InitializeComponent();
            this.Icon = new Icon("AGI.ico");
            InitializeDatabase();
            LoadData();
            FnGetQuittancesPath();
            LoadGerData_dateEncaissement();
            LoadDataQuittanceLocataire();
            LoadGerData();
            LoadGerDataMAJ();
            LoadQuittanceMAJ();
            LoadMonths();
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

                _dataAdapter = new OleDbDataAdapter("SELECT [N] AS [Numero Encaissement], " +
                    "[quittance] AS [Quittance],[code_ger] AS [Code Gerance], [code_loc] AS [Code Locataire]," +
                    "[date_encs] AS [Date Encaissement],[total] AS [Total] FROM Encaissement ORDER BY [N] ASC", _connection);
                OleDbCommandBuilder commandBuilder = new OleDbCommandBuilder(_dataAdapter);

                _dataTable_Encaissement = new DataTable();
                _dataAdapter.Fill(_dataTable_Encaissement);

                DataGridViewListEncaissement.DataSource = _dataTable_Encaissement;
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

        private void FnGetQuittancesPath()
        {
            try
            {
                // Adjust the SQL query to match the exact table and column names
                string selectQuery = "SELECT [Valeur], [Id Param] FROM [Params] WHERE [Id Param] in ('Quittance', 'EtatDetaille')";

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

                        for (int i = 0; i < dataTable.Rows.Count; i++)
                        {
                            DataRow row = dataTable.Rows[i];

                            if (row["Id Param"].ToString() == "Quittance")
                            {
                                QuittancePath = row["Valeur"].ToString();
                            }
                            else
                            {
                                EtatDetaillePath = row["Valeur"].ToString();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                
            }
            finally
            {
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

                txtMoisPaye.DisplayMember = "MoisPaye";
                txtMoisPaye.ValueMember = "MoisPaye";
                txtMoisPaye.DataSource = dataTable;

                MAJtxtPeriodePaye.DisplayMember = "MoisPaye";
                MAJtxtPeriodePaye.ValueMember = "MoisPaye";
                MAJtxtPeriodePaye.DataSource = dataTable;
                
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadDataQuittanceLocataire()
        {
            try
            {
                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open();
                }

                _dataAdapter = new OleDbDataAdapter("SELECT [code loc] AS [Code Locataire], [nom loc] AS [Nom Locataire]," +
                    " [CODE GER] AS [Code Gerance], [total] AS [Total], [CHARGE]  AS [Charge], " +
                    " [TAXE] AS [Taxe], [Champ1] AS [Usage], [Adresse] AS [Adresse] FROM loc " +
                    " WHERE [code loc] IS NOT NULL ORDER BY IIF(IsNumeric([code loc] & ''), Val([code loc]), NULL), [code loc]", _connection);

                OleDbCommandBuilder commandBuilder = new OleDbCommandBuilder(_dataAdapter);

                _dataTable = new DataTable();
                _dataAdapter.Fill(_dataTable);

                // Add checkbox column
                DataGridViewCheckBoxColumn checkBoxColumn = new DataGridViewCheckBoxColumn();
                checkBoxColumn.HeaderText = "";
                checkBoxColumn.Name = "Select";
                checkBoxColumn.Width = 50;
                DataGridViewListLocataires.Columns.Add(checkBoxColumn);

                DataGridViewListLocataires.DataSource = _dataTable;
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

        private void LoadQuittanceMAJ()
        {
            try
            {
                string selectQueryLocataire = "SELECT [quittance] FROM [Encaissement] ORDER BY [quittance]";

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
                    emptyRow["quittance"] = DBNull.Value; // Placeholder for empty row
                    dataTable1.Rows.InsertAt(emptyRow, 0);  // Insert empty row at the top

                    MAJtxtQuitance.DisplayMember = "quittance"; // Display this column in ComboBox
                    MAJtxtQuitance.ValueMember = "quittance"; // Use this column as the value

                    MAJtxtQuitance.DataSource = dataTable1;
                }

            }
            catch (Exception ex)
            {
                // Output the full exception details to help diagnose the issue
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

        private void LoadMonths()
        {
            try
            {
                
                string[] monthNames = DateTimeFormatInfo.CurrentInfo.MonthNames;

                foreach (string month in monthNames)
                {
                    if (!string.IsNullOrEmpty(month))
                    {
                        txtPeriodPay.Items.Add(month + " / " + DateTime.Now.AddMonths(+1).Year);
                    }
                }

                txtPeriodPay.SelectedIndex = DateTime.Now.Month - 1;

            }
            catch
            {

            }
        }

        private void LoadGerDataMAJ()
        {
            try
            {
                string selectQueryLocataire = "SELECT [code loc] FROM [loc] " +
                    "WHERE [code loc] IS NOT NULL ORDER BY IIF(IsNumeric([code loc] & ''), Val([code loc]), NULL), [code loc]";

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

                    // Add an empty row for the default selection
                    DataRow emptyRow = dataTable1.NewRow();
                    emptyRow["code loc"] = DBNull.Value; // Placeholder for empty row
                    dataTable1.Rows.InsertAt(emptyRow, 0);  // Insert empty row at the top

                    // Bind the ComboBox controls to the DataTable
                    MAJtxtCodeLocataire.DisplayMember = "code loc"; // Column to display
                    MAJtxtCodeLocataire.ValueMember = "code loc"; // Column to use as value
                    MAJtxtCodeLocataire.DataSource = dataTable1;

                    // Bind the second ComboBox as well
                    txtCodeLocataire.DisplayMember = "code loc"; // Column to display
                    txtCodeLocataire.ValueMember = "code loc"; // Column to use as value
                    txtCodeLocataire.DataSource = dataTable1;
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

        private void LoadGerData_dateEncaissement()
        {
            try
            {
                string selectQuery = "SELECT FORMAT([date_encs], 'yyyy/MM/dd') AS [date_encs] FROM [Encaissement] GROUP BY [date_encs] ORDER BY [date_encs] DESC";

                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open();
                }

                using (OleDbDataAdapter adapter = new OleDbDataAdapter(selectQuery, _connection))
                {
                    DataTable dataTable = new DataTable();
                    adapter.Fill(dataTable);

                    txtRechercheDateEncaissmentt.DisplayMember = "date_encs"; // Display this column in ComboBox
                    txtRechercheDateEncaissmentt.ValueMember = "date_encs"; // Use this column as the value
                    txtRechercheDateEncaissmentt.DataSource = dataTable;
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

                    txtGerance.DisplayMember = "CODE GER"; // Display this column in ComboBox
                    txtGerance.ValueMember = "CODE GER"; // Use this column as the value
                    txtGerance.DataSource = dataTable;

                    MAJtxtNomGerance.DisplayMember = "CODE GER"; // Display this column in ComboBox
                    MAJtxtNomGerance.ValueMember = "CODE GER"; // Use this column as the value
                    MAJtxtNomGerance.DataSource = dataTable;
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
            if (DataGridViewListEncaissement.SelectedRows.Count > 0)
            {
                try
                {
                    // Get the selected row
                    DataGridViewRow selectedRow = DataGridViewListEncaissement.SelectedRows[0];

                    // Get the unique value (e.g., code loc) of the row to delete
                    string codeQuittance = selectedRow.Cells["quittance"].Value.ToString();

                    // Remove the row from DataGridView
                    DataGridViewListEncaissement.Rows.Remove(selectedRow);

                    // Delete the corresponding row from the database
                    string deleteQuery = "DELETE FROM Encaissement WHERE [quittance] = @Quittance";

                    using (OleDbCommand cmd = new OleDbCommand(deleteQuery, _connection))
                    {
                        cmd.Parameters.AddWithValue("@Quittance", codeQuittance);

                        if (_connection.State != ConnectionState.Open)
                        {
                            _connection.Open();
                        }

                        int rowsAffected = cmd.ExecuteNonQuery();
                        if (rowsAffected > 0)
                        {
                            MessageBox.Show("Encaissement a été supprimé avec succès.");
                        }
                        else
                        {
                            MessageBox.Show("Erreur lors de la suppression d'encaissement.");
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
        private void btnLoad_Click(object sender, EventArgs e)
        {
            LoadData();
        }
        private void btnCanLamSan_Click_1(object sender, EventArgs e)
        {
            string MAX_ID_Encaissement = "1";

            try
            {
                string Query_MAX_ID_Encaissement = "SELECT MAX([N]) AS N FROM Encaissement";

                // Check if the connection is not already open
                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open(); // Open the connection
                }

                // Use OleDbCommand to fetch the data with parameterized query
                using (OleDbCommand command = new OleDbCommand(Query_MAX_ID_Encaissement, _connection))
                {
                    using (OleDbDataAdapter adapter_ID_Encaissment = new OleDbDataAdapter(command))
                    {
                        // Fill a DataTable with the results
                        DataTable dataTable_ID_Encaissment = new DataTable();

                        adapter_ID_Encaissment.Fill(dataTable_ID_Encaissment); // Fill the DataTable

                        if (dataTable_ID_Encaissment.Rows.Count > 0)
                        {
                            // Assuming you want to show the first row's data
                            DataRow row = dataTable_ID_Encaissment.Rows[0];
                            int Res_Int = Convert.ToInt32(row["N"].ToString());
                            Res_Int++;
                            MAX_ID_Encaissement = Res_Int.ToString();
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
                MessageBox.Show("La table Encaissement est Vide ! c'est la première insertion");
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
                bool test1 = RadioBtnMoisPaye.Checked;
                bool test2 = RadioBtnDifdeLoye.Checked;

                string textCodeLocataire_ = String.IsNullOrWhiteSpace(txtCodeLocataire.Text) ? String.Empty : txtCodeLocataire.Text;
                string txtGerance_ = txtGerance.SelectedValue != null ? txtGerance.SelectedValue.ToString() : String.Empty;
                int txtQuittance_ = String.IsNullOrWhiteSpace(txtQuittance.Text) ? 0 : Convert.ToInt32(txtQuittance.Text);
                float txtMontantLoyer_ = 0;
                if (!String.IsNullOrWhiteSpace(txtMontantLoyer.Text))
                {
                    string formattedText = txtMontantLoyer.Text.Replace(',', '.');
                    if (!float.TryParse(formattedText, NumberStyles.Float, CultureInfo.InvariantCulture, out txtMontantLoyer_))
                    {
                        MessageBox.Show("Veuillez entrer un montant valide pour le loyer !");
                        return;
                    }
                }
                string datloc_ = String.IsNullOrWhiteSpace(txtDateEncaissement.Text) ? DateTime.Today.ToString() : DateTime.Parse(txtDateEncaissement.Text).ToString("yyyy/MM/dd");
                string txtMoisPaye_ = txtMoisPaye.SelectedValue!= null ? txtMoisPaye.SelectedValue.ToString() : String.Empty;
                string txtNomLocataire_ = String.IsNullOrWhiteSpace(txtNomLocataire.Text) ? String.Empty : txtNomLocataire.Text;

                string insertQuery = "INSERT INTO Encaissement ([N], [code_loc], [date_encs], [quittance], [code_ger], [total], [intitulé], [Person_Paye], [Nom_Locataire]) " +
                                     "VALUES (@N, @codeLoc, @date_encs, @quittance, @code_ger, @def_loy, @intitule, @Period_Paye, @txtNomLocataire)";

                using (OleDbCommand cmd = new OleDbCommand(insertQuery, _connection))
                {
                    cmd.Parameters.AddWithValue("@N", MAX_ID_Encaissement);
                    cmd.Parameters.AddWithValue("@codeLoc", textCodeLocataire_);
                    cmd.Parameters.AddWithValue("@date_encs", datloc_); ;
                    cmd.Parameters.AddWithValue("@quittance", txtQuittance_);
                    cmd.Parameters.AddWithValue("@code_ger", txtGerance_);
                    cmd.Parameters.AddWithValue("@def_loy", txtMontantLoyer_);
                    cmd.Parameters.AddWithValue("@intitule", txtMoisPaye_);
                    cmd.Parameters.AddWithValue("@Period_Paye", "");
                    cmd.Parameters.AddWithValue("@txtNomLocataire", txtNomLocataire_);
                    // Add more parameters as needed

                    if (_connection.State != ConnectionState.Open)
                    {
                        _connection.Open();
                    }
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Encaissement Bien Ajouté ");

                    Update_PaymentLocataire(textCodeLocataire_);
                    //InitializeComponent();
                    LoadData();
                    Refresh();
                    LoadQuittanceMAJ();
                    LoadGerData_dateEncaissement();
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

        private void Update_PaymentLocataire(string CodeLocataire)
        {
            int MoisPaye = 0;
            int MoiImpay = 0;

            try
            {
                // Define the SQL query with a WHERE clause to filter by the selected "code loc"
                string selectQuery = "SELECT [py], [impy] FROM [loc] WHERE [code loc] = ?";

                // Check if the connection is not already open
                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open(); // Open the connection
                }

                // Use OleDbCommand to fetch the data with parameterized query
                using (OleDbCommand command = new OleDbCommand(selectQuery, _connection))
                {
                    // Add parameter to the query to prevent SQL injection
                    command.Parameters.AddWithValue("?", CodeLocataire);

                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                    {
                        // Fill a DataTable with the results
                        DataTable dataTable = new DataTable();

                        adapter.Fill(dataTable); // Fill the DataTable

                        if (dataTable.Rows.Count > 0)
                        {
                            DataRow row = dataTable.Rows[0];
                            MoisPaye = Convert.ToInt32(row["py"].ToString());
                            MoiImpay = Convert.ToInt32(row["impy"].ToString());

                            MoisPaye++;

                            if (MoiImpay != 0)
                            {
                                MoiImpay--;
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

            /*
            try
            {
                string PeriodeImp = "";

                string updateQuery = "UPDATE loc SET [py] = @MoiPy, [impy] = @MoiImpy, [periode_py] = @PeriodePay, " +
                    "[periode_imp] = @PeriodeImp WHERE [code loc] = @codeLoc"; // Corrected WHERE clause to use SelectedValue

                using (OleDbCommand cmd = new OleDbCommand(updateQuery, _connection))
                {
                    cmd.Parameters.AddWithValue("@MoiImpy", MoisPaye);
                    cmd.Parameters.AddWithValue("@MoiPy", MoiImpay);
                    cmd.Parameters.AddWithValue("@codeLoc", CodeLocataire);
                    cmd.Parameters.AddWithValue("@PeriodePay", PeriodePay);
                    cmd.Parameters.AddWithValue("@PeriodeImp", PeriodeImp);

                    if (_connection.State != ConnectionState.Open)
                    {
                        _connection.Open(); // Open the connection
                    }
                    cmd.ExecuteNonQuery();

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
            */
        }

        private void Refresh()
        {
            try
            {
                txtQuittance.Text = null;
                txtCodeLocataire.SelectedValue = DBNull.Value;
                txtGerance.SelectedValue = DBNull.Value;
                txtMoisPaye.SelectedValue = DBNull.Value;
                txtMontantLoyer.Text = null;
                txtNomLocataire.Text = null;
                txtAdresseLocataire.Text = null;
            }
            catch(Exception ex)
            {

            }
        }

        /*-- UPDATE DATETIME VALUE ----------------------------------------------------------------*/
        private void UpdateAllDateValue()
        {
            try
            {
                string selectedQuittance = MAJtxtQuitance.SelectedValue?.ToString();
                string selectedCodeLocataire = MAJtxtCodeLocataire.SelectedValue?.ToString();

                if (string.IsNullOrEmpty(selectedQuittance))
                {
                    MessageBox.Show("No valid locataire selected.");
                    return;
                }

                string selectQuery = "SELECT [date_encs], [code_ger], [def_loy] FROM [Encaissement] WHERE [quittance] = @quittance";

                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open(); // Open the connection
                }

                using (OleDbCommand command = new OleDbCommand(selectQuery, _connection))
                {
                    command.Parameters.AddWithValue("@Quittance", selectedQuittance);

                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                    {
                        DataTable dataTable = new DataTable();
                        adapter.Fill(dataTable); // Fill the DataTable

                        for (int i = 0; i < dataTable.Rows.Count; i++)
                        {
                            DataRow row = dataTable.Rows[0];
                            string dateEncaissement = row["date_encs"].ToString();

                            // Check if the value is already a DateTime
                            if (row["date_encs"] is DateTime)
                            {
                                DateTime parsedDate = (DateTime)row["date_encs"];
                                string formattedDate = parsedDate.ToString("dd/MM/yyyy");

                                // Update and display
                                UpdateDateEncaissement(selectedQuittance, formattedDate);
                            }
                            else
                            {
                                // Try to parse if it's a string, assuming the format is known (e.g., yyyyMMdd)
                                string[] formats = { "ddMMyyyy", "dd/MM/yyyy", "MM/dd/yyyy" }; // Add your possible formats here

                                if (DateTime.TryParseExact(dateEncaissement, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate))
                                {
                                    string formattedDate = parsedDate.ToString("dd/MM/yyyy");

                                    // Update and display
                                    UpdateDateEncaissement(selectedQuittance, formattedDate);
                                }
                                else
                                {
                                    MessageBox.Show("Invalid date format in the database.");
                                }
                            }

                            MAJtxtNomGerance.SelectedValue = row["code_ger"].ToString();
                            MAJtxtDefLoy.Text = row["def_loy"].ToString();
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
                if (_connection.State == ConnectionState.Open)
                {
                    _connection.Close();
                }
            }
        }
        private void UpdateDateEncaissement(string quittance, string formattedDate)
        {
            string updateQuery = "UPDATE [Encaissement] SET [date_encs] = @formattedDate WHERE [quittance] = @quittance";

            using (OleDbCommand updateCommand = new OleDbCommand(updateQuery, _connection))
            {
                updateCommand.Parameters.AddWithValue("@formattedDate", formattedDate);
                updateCommand.Parameters.AddWithValue("@quittance", quittance);

                int rowsAffected = updateCommand.ExecuteNonQuery();
            }

            MAJtxtDateEncaissement.Text = formattedDate;
        }

        /*-- UPDATE DATETIME VALUE ----------------------------------------------------------------*/

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                bool test1 = RadioBtnMoisPaye.Checked;
                bool test2 = RadioBtnDifdeLoye.Checked;

                // Récupération des valeurs
                string MAJtxtNomGerance_ = MAJtxtNomGerance.SelectedValue?.ToString() ?? string.Empty;
                string MAJtxtCodeLocataire_ = MAJtxtCodeLocataire.SelectedValue?.ToString() ?? string.Empty;
                int MAJtxtQuitance_ = string.IsNullOrWhiteSpace(MAJtxtQuitance.Text) ? 0 : Convert.ToInt32(MAJtxtQuitance.Text);

                float MAJtxtTotal_ = 0;
                if (!string.IsNullOrWhiteSpace(MAJtxtTotal.Text))
                {
                    string formattedText = MAJtxtTotal.Text.Replace(',', '.');
                    if (!float.TryParse(formattedText, NumberStyles.Float, CultureInfo.InvariantCulture, out MAJtxtTotal_))
                    {
                        MessageBox.Show("Veuillez entrer un montant valide pour le loyer !");
                        return;
                    }
                }

                string MAJtxtDateEncaissement_ = string.IsNullOrWhiteSpace(MAJtxtDateEncaissement.Text)
                    ? DateTime.Today.ToString("yyyy-MM-dd")
                    : DateTime.Parse(MAJtxtDateEncaissement.Text).ToString("yyyy-MM-dd");
                string MAJtxtMoiPaye_ = MAJtxtMoiPaye.Text?.Trim() ?? string.Empty;
                string MAJtxtPeriodePaye_ = MAJtxtPeriodePaye.SelectedValue?.ToString() ?? string.Empty;

                string updateQuery = @"UPDATE Encaissement SET 
                               [date_encs] = @date_encs, 
                               [code_loc] = @code_loc, 
                               [code_ger] = @code_ger, 
                               [total] = @Total, 
                               [intitulé] = @MoiPayer, 
                               [Person_Paye] = @MAJtxtPeriodePaye 
                           WHERE [quittance] = @quittance";

                using (OleDbCommand cmd = new OleDbCommand(updateQuery, _connection))
                {
                    cmd.Parameters.Add("@date_encs",MAJtxtDateEncaissement_ );
                    cmd.Parameters.Add("@code_loc",  MAJtxtCodeLocataire_ );
                    cmd.Parameters.Add("@code_ger", MAJtxtNomGerance_);
                    cmd.Parameters.Add("@Total",MAJtxtTotal_ );
                    cmd.Parameters.Add("@MoiPayer", MAJtxtMoiPaye_ );
                    cmd.Parameters.Add("@MAJtxtPeriodePaye", MAJtxtPeriodePaye_ );
                    cmd.Parameters.Add("@quittance", MAJtxtQuitance_ );

                    if (_connection.State != ConnectionState.Open)
                    {
                        _connection.Open();
                    }

                    int rowsAffected = cmd.ExecuteNonQuery();
                    if (rowsAffected > 0)
                    {
                        MessageBox.Show("Encaissement bien modifié.");
                    }
                    else
                    {
                        MessageBox.Show("Aucune ligne mise à jour. Vérifiez que la quittance existe.");
                    }

                    LoadData();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Veuillez vérifier les champs de saisie ! : " + ex.Message);
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
            
        }

        /****************************************************************************************************/

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
        
        private void MAJtxtQuitance_SelectedIndexChanged(object sender, EventArgs e)
        {
            ChargerDonneesLocataireParCode();
        }

        private void ChargerDonneesLocataireParCode()
        {
            try
            {
                string selectedQuittance = MAJtxtQuitance.SelectedValue?.ToString();
                string selectedCodeLocataire = MAJtxtCodeLocataire.SelectedValue?.ToString();

                if (string.IsNullOrEmpty(selectedQuittance))
                {
                    //MessageBox.Show("Numéro de Quittance séléctionné est invalid !");
                    return;
                }

                string selectQuery = " SELECT [date_encs], [code_ger], [code_loc], [def_loy], [intitulé], [total]" +
                                     " FROM [Encaissement] WHERE [quittance] = @quittance";

                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open(); // Open the connection
                }

                // Use OleDbCommand to fetch the data with parameterized query
                using (OleDbCommand command = new OleDbCommand(selectQuery, _connection))
                {
                    // Add parameter to the query to prevent SQL injection
                    command.Parameters.AddWithValue("@Quittance", selectedQuittance);

                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                    {
                        // Fill a DataTable with the results
                        DataTable dataTable = new DataTable();

                        adapter.Fill(dataTable); // Fill the DataTable

                        if (dataTable.Rows.Count > 0)
                        {
                            // Assuming you want to show the first row's data
                            DataRow row = dataTable.Rows[0];

                            DateTime DtEncaissement = DateTime.Now;

                            try
                            {
                                DtEncaissement = Convert.ToDateTime(row["date_encs"].ToString());
                            }
                            catch { }

                            MAJtxtDateEncaissement.Value = DtEncaissement;
                            MAJtxtNomGerance.SelectedValue = row["code_ger"].ToString();
                            MAJtxtDefLoy.Text = row["def_loy"].ToString();
                            MAJtxtCodeLocataire.SelectedValue = row["code_loc"].ToString();
                            MAJtxtMoiPaye.Text = row["intitulé"].ToString();
                            MAJtxtTotal.Text = row["total"].ToString();
                            MAJtxtPeriodePaye.SelectedValue = row["intitulé"].ToString();
                        }
                        else
                        {
                            MessageBox.Show("No data found for the selected Encaissement.");
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

        public void GeneratePdfEncaissement(string filePath,string imagePath, string codeGer, string NomCoprop, 
            string Adresse1, string Adresse2, string Adresse3,
            List<string> ListeCopor, List<string> PartCopor,
            List<string> ListCodeLocataires, List<string> ListNomLocataires,
            List<double> LitsLoyer, List<string> ListPerdioPay, List<string> ListPerdioImpy, List<double> ListMontantPay,
            List<double> ListMontantImp, double commission, double TVA, double Total, double Sold)
        {
            
            // Define the fonts
            iTextSharp.text.Font headerFont = FontFactory.GetFont(FontFactory.COURIER_BOLD, 8);
            iTextSharp.text.Font regularFont = FontFactory.GetFont(FontFactory.COURIER, 8);
            iTextSharp.text.Font boldFont = FontFactory.GetFont(FontFactory.COURIER, 6);

            // Create a new document
            Document document = new Document(PageSize.A4, 30, 30, 10, 10);

            try
            {
                #region Entete ------------------------------------------------------------------------------------- 
                    // Set up the PDF writer
                    PdfWriter.GetInstance(document, new FileStream(filePath, FileMode.Create));
                    document.Open();

                // Load and place the image in the top-right corner
                    iTextSharp.text.Image logo = iTextSharp.text.Image.GetInstance(imagePath);
                    logo.ScaleToFit(80, 50); // Adjust the size of the image
                    logo.SetAbsolutePosition(document.PageSize.Width - logo.ScaledWidth - 10, document.PageSize.Height - logo.ScaledHeight - 75);
                    document.Add(logo);

                    // Header
                    Paragraph header = new Paragraph("AGI - agence de gestion immobilière - SARL au capital de 240.000 dhs\n" +
                                                     "9, rue Oued Zam Casablanca 20.000\n" +
                                                     "Tel : 0522 263924 - 0522 275296 - Fax : 0522 223020 - C.N.S.S : 1540500 - R.C : 76.923 - Patente : 35501274\n\n",
                                                     headerFont);
                    header.Alignment = Element.ALIGN_CENTER;
                    document.Add(header);

                    // Date and recipient details
                    Paragraph dateAndRecipient = new Paragraph("Casablanca, le "+ DateTime.Today.ToShortDateString() +"\n\n" +
                                                               ""+ NomCoprop + "\n" +
                                                               ""+ Adresse1 + "\n" +
                                                               ""+ Adresse2 + "\n" +
                                                               ""+ Adresse3 + "\n" +
                                                               "N° Gérance : "+ codeGer + "\n"
                                                               // + "10 RUE POUPONNIERE\n\n"
                                                               ,regularFont);
                            dateAndRecipient.Alignment = Element.ALIGN_LEFT;
                            document.Add(dateAndRecipient);

                        string MadameMonsieru = "";

                        if (NomCoprop.StartsWith("Madame", StringComparison.OrdinalIgnoreCase))
                        {
                            MadameMonsieru = "Madame";
                        }
                        else if (NomCoprop.StartsWith("Monsieur", StringComparison.OrdinalIgnoreCase))
                        {
                            MadameMonsieru = "Monsieur";
                        }
                        else
                        {
                            MadameMonsieru = "Madame Monsieur";
                        }

                // Opening statement
                Paragraph opening = new Paragraph(""+ MadameMonsieru + ",\n" +
                                                      "Nous vous adressons ci-dessous relevé des loyers du Trimestre "+DateTime.Now.Year+"\n\n",
                                                      regularFont);
                    document.Add(opening);

                    // Create a table with 6 columns for "Code Loc", "Nom Locataire", "Loyer", etc.
                    PdfPTable table = new PdfPTable(7);
                    table.WidthPercentage = 100;
                    table.SetWidths(new float[] { 2, 3, 2, 2, 2, 2, 2 });

                    // Add table headers
                    PdfPCell[] headers = {
                        new PdfPCell(new Phrase("Code Loc", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER },
                        new PdfPCell(new Phrase("Nom Locataire", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER },
                        new PdfPCell(new Phrase("Loyer", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER },
                        new PdfPCell(new Phrase("Période payé", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER },
                        new PdfPCell(new Phrase("Montant payé", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER },
                        new PdfPCell(new Phrase("Période impy", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER },
                        new PdfPCell(new Phrase("Montant impy", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER }
                    };

                    foreach (var headerCell in headers)
                    {
                        headerCell.Padding = 5;
                        headerCell.BackgroundColor = BaseColor.LIGHT_GRAY;
                        table.AddCell(headerCell);
                    }

                    for(int i = 0; i < ListCodeLocataires.Count; i++)
                    {
                        table.AddCell(new PdfPCell(new Phrase(ListCodeLocataires[i], boldFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                        table.AddCell(new PdfPCell(new Phrase(ListNomLocataires[i], boldFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                        table.AddCell(new PdfPCell(new Phrase(LitsLoyer[i].ToString("00.00"), boldFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                        table.AddCell(new PdfPCell(new Phrase(ListPerdioPay[i], boldFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                        table.AddCell(new PdfPCell(new Phrase(ListMontantPay[i].ToString("00.00"), boldFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                        table.AddCell(new PdfPCell(new Phrase(ListPerdioImpy[i], boldFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                        table.AddCell(new PdfPCell(new Phrase(ListMontantImp[i].ToString("00.00"), boldFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                    }


                    PdfPCell totalLabelCell = new PdfPCell(new Phrase("Total : ", boldFont)) { Colspan = 4, HorizontalAlignment = Element.ALIGN_RIGHT };
                    table.AddCell(totalLabelCell);
                    table.AddCell(new PdfPCell(new Phrase(ListMontantPay.Sum().ToString("00.00"))) { HorizontalAlignment = Element.ALIGN_RIGHT });
                    table.AddCell(new PdfPCell(new Phrase(ListMontantImp.Sum().ToString("00.00"))) { HorizontalAlignment = Element.ALIGN_RIGHT });

                    document.Add(table);
                    #endregion

                #region table Montant ------------------------------------------------------------------
                // Code to add "Dépense" table after the main table
                PdfPTable expenseTable = new PdfPTable(2);
                expenseTable.WidthPercentage = 50;
                expenseTable.SetWidths(new float[] { 3, 1 });
                expenseTable.SpacingBefore = 20; // Adds some space between main table and expense table

                // Header for the "Dépense" table
                PdfPCell expenseHeader1 = new PdfPCell(new Phrase("Dépense", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER };
                PdfPCell expenseHeader2 = new PdfPCell(new Phrase("Montant", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER };
                expenseHeader1.BackgroundColor = BaseColor.LIGHT_GRAY;
                expenseHeader2.BackgroundColor = BaseColor.LIGHT_GRAY;

                expenseTable.AddCell(expenseHeader1);
                expenseTable.AddCell(expenseHeader2);

                // Add data rows for the "Dépense" table
                expenseTable.AddCell(new PdfPCell(new Phrase("Commission", regularFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                expenseTable.AddCell(new PdfPCell(new Phrase("490,00", regularFont)) { HorizontalAlignment = Element.ALIGN_RIGHT });

                expenseTable.AddCell(new PdfPCell(new Phrase("TVA", regularFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                expenseTable.AddCell(new PdfPCell(new Phrase("98,00", regularFont)) { HorizontalAlignment = Element.ALIGN_RIGHT });

                // Add Total row
                PdfPCell totalLabelCellExpense = new PdfPCell(new Phrase("Total :", boldFont)) { HorizontalAlignment = Element.ALIGN_RIGHT };
                totalLabelCellExpense.Colspan = 1; // Span over one cell
                expenseTable.AddCell(totalLabelCellExpense);
                expenseTable.AddCell(new PdfPCell(new Phrase("588,00", boldFont)) { HorizontalAlignment = Element.ALIGN_RIGHT });

                // Add the "Dépense" table to the document
                document.Add(expenseTable);
                    #endregion

                #region Solde ---------------------------------------------------------------------------------------
                // Code to add "Solde" table below the "Dépense" table
                PdfPTable soldeTable = new PdfPTable(2);
                soldeTable.WidthPercentage = 30;
                soldeTable.HorizontalAlignment = Element.ALIGN_RIGHT; // Aligns the table to the right
                soldeTable.SpacingBefore = 10; // Adds space between "Dépense" table and "Solde"

                // Add "Solde" row
                PdfPCell soldeLabelCell = new PdfPCell(new Phrase("Solde :", boldFont))
                {
                    Border = iTextSharp.text.Rectangle.NO_BORDER, // No border to match the styling
                    HorizontalAlignment = Element.ALIGN_RIGHT
                };
                PdfPCell soldeAmountCell = new PdfPCell(new Phrase("4 312,00", boldFont))
                {
                    Border = iTextSharp.text.Rectangle.NO_BORDER, // No border to match the styling
                    HorizontalAlignment = Element.ALIGN_RIGHT
                };

                soldeTable.AddCell(soldeLabelCell);
                soldeTable.AddCell(soldeAmountCell);

                // Add the "Solde" table to the document
                document.Add(soldeTable);
                    #endregion

                #region table Copropriétaire -----------------------------------------------------------------------
                    // Create the "Copropriétaire" table with three columns
                    PdfPTable coproprietaireTable = new PdfPTable(3);
                    coproprietaireTable.WidthPercentage = 60;
                    coproprietaireTable.SpacingBefore = 10; // Adds space before this table

                    // Set widths for the columns: Copropriétaire (50%), Taux (25%), Montant (25%)
                    coproprietaireTable.SetWidths(new float[] { 2f, 1f, 1f });

                    // Add headers
                    PdfPCell coproprietaireHeaderCell = new PdfPCell(new Phrase("Copropriétaire", boldFont)) { HorizontalAlignment = Element.ALIGN_LEFT };
                    PdfPCell tauxHeaderCell = new PdfPCell(new Phrase("Taux", boldFont)) { HorizontalAlignment = Element.ALIGN_LEFT };
                    PdfPCell montantHeaderCell = new PdfPCell(new Phrase("Montant", boldFont)) { HorizontalAlignment = Element.ALIGN_RIGHT };

                    coproprietaireTable.AddCell(coproprietaireHeaderCell);
                    coproprietaireTable.AddCell(tauxHeaderCell);
                    coproprietaireTable.AddCell(montantHeaderCell);

                    // Add data rows
                    coproprietaireTable.AddCell(new PdfPCell(new Phrase("Madame ZIANE Lamia", regularFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                    coproprietaireTable.AddCell(new PdfPCell(new Phrase("35/64", regularFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                    coproprietaireTable.AddCell(new PdfPCell(new Phrase("2 358,13", regularFont)) { HorizontalAlignment = Element.ALIGN_RIGHT });

                    coproprietaireTable.AddCell(new PdfPCell(new Phrase("Madame JAWAANI Fatima", regularFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                    coproprietaireTable.AddCell(new PdfPCell(new Phrase("22/64", regularFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                    coproprietaireTable.AddCell(new PdfPCell(new Phrase("1 482,25", regularFont)) { HorizontalAlignment = Element.ALIGN_RIGHT });

                    coproprietaireTable.AddCell(new PdfPCell(new Phrase("Madame ZIANE Latifa", regularFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                    coproprietaireTable.AddCell(new PdfPCell(new Phrase("7/64", regularFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                    coproprietaireTable.AddCell(new PdfPCell(new Phrase("471,63", regularFont)) { HorizontalAlignment = Element.ALIGN_RIGHT });

                    // Add the "Copropriétaire" table to the document
                    document.Add(coproprietaireTable);
                    #endregion

                #region fouter -----------------------------------------------------------------------------------------
                Paragraph closingMessage = new Paragraph("Veuillez agréer, "+ MadameMonsieru + ", l'expression de nos meilleures salutations.\n\n", regularFont);
                closingMessage.Alignment = Element.ALIGN_LEFT;
                document.Add(closingMessage);

                // Add signature line
                Paragraph signature = new Paragraph("Le gérant\n\n", regularFont);
                signature.Alignment = Element.ALIGN_CENTER;
                document.Add(signature);

                // Add disclaimer
                /*
                Paragraph disclaimer = new Paragraph("<< Sauf erreur ou omission >>", regularFont);
                disclaimer.Alignment = Element.ALIGN_CENTER;
                document.Add(disclaimer);
                */
                // Close the document
                document.Close();
                #endregion

            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
            finally
            {
                document.Close();
            }
        }

        public void GenerateDateEncaissement_EtatJornalie(string filePath, List<string> ListCodeLocataires, List<string> ListNomLocataires,
        List<string> ListeQuittance, List<string> ListPerdioPay, List<string> ListCodeGer, List<double> ListMontantPay,
        List<double> ListMontantImp, double commission, double TVA, double Total, double Sold)
            {

            // Define the fonts
            iTextSharp.text.Font headerFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8);
            iTextSharp.text.Font regularFont = FontFactory.GetFont(FontFactory.HELVETICA, 8);
            iTextSharp.text.Font boldFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8);

            // Create a new document
            Document document = new Document(PageSize.A4);

            try
            {
                #region Entete ------------------------------------------------------------------------------------- 
                // Set up the PDF writer
                PdfWriter.GetInstance(document, new FileStream(filePath, FileMode.Create));
                document.Open();

                // Header
                Paragraph header = new Paragraph("AGI - agence de gestion immobilière - SARL au capital de 240.000 dhs\n" +
                                                 "9, rue Oued Zam Casablanca 20.000\n" +
                                                 "Tel : 0522 263924 - 0522 275296 - Fax : 0522 223020 - C.N.S.S : 1540500 - R.C : 76.923 - Patente : 35501274\n\n",
                                                 headerFont);
                header.Alignment = Element.ALIGN_CENTER;
                document.Add(header);

                // Date and recipient details
                Paragraph dateAndRecipient = new Paragraph("", regularFont);
                dateAndRecipient.Alignment = Element.ALIGN_LEFT;
                document.Add(dateAndRecipient);

                // Opening statement
                Paragraph opening = new Paragraph("", regularFont);
                document.Add(opening);

                // Create a table with 6 columns for "Code Loc", "Nom Locataire", "Loyer", etc.
                PdfPTable table = new PdfPTable(6);
                table.WidthPercentage = 100;
                table.SetWidths(new float[] { 2, 3, 2, 2, 2, 2 });

                for(int j = 0; j < ListCodeGer.Count; j++)
                {
                    PdfPCell[] headers = {
                        new PdfPCell(new Phrase("Code Gerance", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER },
                        new PdfPCell(new Phrase("Code Locataire", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER },
                        new PdfPCell(new Phrase("Nom Locataire", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER },
                        new PdfPCell(new Phrase("Quittance", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER },
                        new PdfPCell(new Phrase("Mois Payé", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER },
                        //new PdfPCell(new Phrase("Période impy", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER },
                        new PdfPCell(new Phrase("Montant impy", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER }
                    };
                    foreach (var headerCell in headers)
                    {
                        headerCell.Padding = 5;
                        headerCell.BackgroundColor = BaseColor.LIGHT_GRAY;
                        table.AddCell(headerCell);
                    }

                    for (int i = 0; i < ListCodeLocataires.Count; i++)
                    {
                        table.AddCell(new PdfPCell(new Phrase(ListCodeGer[j])) { HorizontalAlignment = Element.ALIGN_LEFT });
                        table.AddCell(new PdfPCell(new Phrase(ListCodeLocataires[i])) { HorizontalAlignment = Element.ALIGN_LEFT });
                        table.AddCell(new PdfPCell(new Phrase(ListNomLocataires[i].ToString())) { HorizontalAlignment = Element.ALIGN_LEFT });
                        table.AddCell(new PdfPCell(new Phrase(ListeQuittance[i])) { HorizontalAlignment = Element.ALIGN_LEFT });
                        table.AddCell(new PdfPCell(new Phrase(ListPerdioPay[i].ToString())) { HorizontalAlignment = Element.ALIGN_LEFT });
                        //table.AddCell(new PdfPCell(new Phrase(ListPerdioImpy[i])) { HorizontalAlignment = Element.ALIGN_LEFT });
                        table.AddCell(new PdfPCell(new Phrase(ListMontantPay[i].ToString("00.00"))) { HorizontalAlignment = Element.ALIGN_LEFT });
                    }

                }

                PdfPCell totalLabelCell = new PdfPCell(new Phrase("Total Encaiseement: ", boldFont)) { Colspan = 4, HorizontalAlignment = Element.ALIGN_RIGHT };
                table.AddCell(totalLabelCell);
                table.AddCell(new PdfPCell(new Phrase()) { HorizontalAlignment = Element.ALIGN_RIGHT });
                table.AddCell(new PdfPCell(new Phrase(ListMontantImp.Sum().ToString("00.00"))) { HorizontalAlignment = Element.ALIGN_RIGHT });

                document.Add(table);
                #endregion

                #region table Montant ------------------------------------------------------------------
                // Code to add "Dépense" table after the main table
                PdfPTable expenseTable = new PdfPTable(2);
                expenseTable.WidthPercentage = 50;
                expenseTable.SetWidths(new float[] { 3, 1 });
                expenseTable.SpacingBefore = 20; // Adds some space between main table and expense table

                // Header for the "Dépense" table
                PdfPCell expenseHeader1 = new PdfPCell(new Phrase("Dépense", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER };
                PdfPCell expenseHeader2 = new PdfPCell(new Phrase("Montant", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER };
                expenseHeader1.BackgroundColor = BaseColor.LIGHT_GRAY;
                expenseHeader2.BackgroundColor = BaseColor.LIGHT_GRAY;

                expenseTable.AddCell(expenseHeader1);
                expenseTable.AddCell(expenseHeader2);

                // Add data rows for the "Dépense" table
                expenseTable.AddCell(new PdfPCell(new Phrase("Commission", regularFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                expenseTable.AddCell(new PdfPCell(new Phrase("490,00", regularFont)) { HorizontalAlignment = Element.ALIGN_RIGHT });

                expenseTable.AddCell(new PdfPCell(new Phrase("TVA", regularFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                expenseTable.AddCell(new PdfPCell(new Phrase("98,00", regularFont)) { HorizontalAlignment = Element.ALIGN_RIGHT });

                // Add Total row
                PdfPCell totalLabelCellExpense = new PdfPCell(new Phrase("Total :", boldFont)) { HorizontalAlignment = Element.ALIGN_RIGHT };
                totalLabelCellExpense.Colspan = 1; // Span over one cell
                expenseTable.AddCell(totalLabelCellExpense);
                expenseTable.AddCell(new PdfPCell(new Phrase("588,00", boldFont)) { HorizontalAlignment = Element.ALIGN_RIGHT });

                // Add the "Dépense" table to the document
                document.Add(expenseTable);
                #endregion

                #region Solde ---------------------------------------------------------------------------------------
                // Code to add "Solde" table below the "Dépense" table
                PdfPTable soldeTable = new PdfPTable(2);
                soldeTable.WidthPercentage = 30;
                soldeTable.HorizontalAlignment = Element.ALIGN_RIGHT; // Aligns the table to the right
                soldeTable.SpacingBefore = 10; // Adds space between "Dépense" table and "Solde"

                // Add "Solde" row
                PdfPCell soldeLabelCell = new PdfPCell(new Phrase("Solde :", boldFont))
                {
                    Border = iTextSharp.text.Rectangle.NO_BORDER, // No border to match the styling
                    HorizontalAlignment = Element.ALIGN_RIGHT
                };
                PdfPCell soldeAmountCell = new PdfPCell(new Phrase("4 312,00", boldFont))
                {
                    Border = iTextSharp.text.Rectangle.NO_BORDER, // No border to match the styling
                    HorizontalAlignment = Element.ALIGN_RIGHT
                };

                soldeTable.AddCell(soldeLabelCell);
                soldeTable.AddCell(soldeAmountCell);

                // Add the "Solde" table to the document
                document.Add(soldeTable);
                #endregion

                #region table Copropriétaire -----------------------------------------------------------------------
                // Create the "Copropriétaire" table with three columns
                PdfPTable coproprietaireTable = new PdfPTable(3);
                coproprietaireTable.WidthPercentage = 60;
                coproprietaireTable.SpacingBefore = 10; // Adds space before this table

                // Set widths for the columns: Copropriétaire (50%), Taux (25%), Montant (25%)
                coproprietaireTable.SetWidths(new float[] { 2f, 1f, 1f });

                // Add headers
                PdfPCell coproprietaireHeaderCell = new PdfPCell(new Phrase("Copropriétaire", boldFont)) { HorizontalAlignment = Element.ALIGN_LEFT };
                PdfPCell tauxHeaderCell = new PdfPCell(new Phrase("Taux", boldFont)) { HorizontalAlignment = Element.ALIGN_LEFT };
                PdfPCell montantHeaderCell = new PdfPCell(new Phrase("Montant", boldFont)) { HorizontalAlignment = Element.ALIGN_RIGHT };

                coproprietaireTable.AddCell(coproprietaireHeaderCell);
                coproprietaireTable.AddCell(tauxHeaderCell);
                coproprietaireTable.AddCell(montantHeaderCell);

                // Add data rows
                coproprietaireTable.AddCell(new PdfPCell(new Phrase("Madame ZIANE Lamia", regularFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                coproprietaireTable.AddCell(new PdfPCell(new Phrase("35/64", regularFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                coproprietaireTable.AddCell(new PdfPCell(new Phrase("2 358,13", regularFont)) { HorizontalAlignment = Element.ALIGN_RIGHT });

                coproprietaireTable.AddCell(new PdfPCell(new Phrase("Madame JAWAANI Fatima", regularFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                coproprietaireTable.AddCell(new PdfPCell(new Phrase("22/64", regularFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                coproprietaireTable.AddCell(new PdfPCell(new Phrase("1 482,25", regularFont)) { HorizontalAlignment = Element.ALIGN_RIGHT });

                coproprietaireTable.AddCell(new PdfPCell(new Phrase("Madame ZIANE Latifa", regularFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                coproprietaireTable.AddCell(new PdfPCell(new Phrase("7/64", regularFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                coproprietaireTable.AddCell(new PdfPCell(new Phrase("471,63", regularFont)) { HorizontalAlignment = Element.ALIGN_RIGHT });

                // Add the "Copropriétaire" table to the document
                document.Add(coproprietaireTable);
                #endregion

                #region fouter -----------------------------------------------------------------------------------------
                Paragraph closingMessage = new Paragraph("Veuillez agréer, Madame, l'expression de nos meilleures salutations.\n\n", regularFont);
                closingMessage.Alignment = Element.ALIGN_LEFT;
                document.Add(closingMessage);

                // Add signature line
                Paragraph signature = new Paragraph("Le gérant\n\n", regularFont);
                signature.Alignment = Element.ALIGN_CENTER;
                document.Add(signature);

                // Add disclaimer
                Paragraph disclaimer = new Paragraph("<< Sauf erreur ou omission >>", regularFont);
                disclaimer.Alignment = Element.ALIGN_CENTER;
                document.Add(disclaimer);

                // Close the document
                document.Close();
                #endregion

            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
            finally
            {
                document.Close();
            }
        }

        public void GenerateQuittances(string filePath, string imagePath,string NumQuitt, string CodeLoc, string NomLoc, 
            string Adresse, double Loyer, double charge,double Tsc, double Totale, string MoisDe, string Usage)
        {
            // Define the fonts
            iTextSharp.text.Font regularFont = FontFactory.GetFont(FontFactory.COURIER_BOLDOBLIQUE, 5);
            iTextSharp.text.Font boldFont = FontFactory.GetFont(FontFactory.COURIER_OBLIQUE, 5);

            Document document = new Document(PageSize.A5, 2, 2, 2, 2);

            try
            {
                #region Header ------------------------------------------------------------------------------------- 
                // Set up the PDF writer
                PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(filePath, FileMode.Create));
                document.Open();

                // Draw vertical line to separate the main content from the 17% right section
                float separationPosition = document.PageSize.Width * 0.83f;  // Calculate 83% of the page width
                PdfContentByte canvas = writer.DirectContent;
                canvas.MoveTo(separationPosition, document.PageSize.Height); // Start at the top of the page
                canvas.LineTo(separationPosition, 700);                        // Draw to the bottom of the page
                canvas.Stroke();

                // Load and place the image in the top-right corner
                iTextSharp.text.Image logo = iTextSharp.text.Image.GetInstance(imagePath);
                logo.ScaleToFit(80, 50); // Adjust the size of the image
                logo.SetAbsolutePosition(document.PageSize.Width - logo.ScaledWidth - 74, document.PageSize.Height - logo.ScaledHeight - 10);
                document.Add(logo);

                // Create a table with 2 columns for the headers side by side
                PdfPTable headerTable = new PdfPTable(3)
                {
                    WidthPercentage = 100, // Full width of the page
                    SpacingBefore = 0f, // No space before the header table
                    SpacingAfter = 0f, // No space after the header table
                    HorizontalAlignment = Element.ALIGN_LEFT // Align the content to the left
                };

                // First column for the first header (header)
                PdfPCell headerCell1 = new PdfPCell(new Phrase(
                    "Agent de Gestion Immobilière\n" +
                    "S.A.R.L. Au capital de 240.000 Dhs\n" +
                    "Siège social 9 Rue Oued Zam 20 000 Casablanca\n" +
                    "I.C.E 000168441000035\n" +
                    "Tél : 05.22.26.39.24 - Fax : 05.22.22.30.20\n" +
                    "E-mail : agimge@gmail.com\n" +
                    "Site web : agipro.net",
                    FontFactory.GetFont(FontFactory.COURIER_BOLD, 6)
                ))
                {
                    Border = Rectangle.NO_BORDER, // No border for this cell
                    HorizontalAlignment = Element.ALIGN_LEFT
                };


                iTextSharp.text.Font greenFont = FontFactory.GetFont(FontFactory.COURIER_BOLD, 8);
                greenFont.Color = BaseColor.DARK_GRAY;

                iTextSharp.text.Font blueFont = FontFactory.GetFont(FontFactory.COURIER_BOLD, 5);
                blueFont.Color = BaseColor.BLUE;

                Chunk titlePart1 = new Chunk("توصيل الكراء\n", greenFont);
                Chunk titlePart2 = new Chunk("Quittance de Loyer\n\n", greenFont);
                Chunk detailsPart1 = new Chunk("N° : "+ NumQuitt + "n\"" , blueFont);
                Chunk detailsPart2 = new Chunk("Nom : "+ NomLoc + "\n", blueFont);
                Chunk detailsPart3 = new Chunk("Code : "+ CodeLoc + "\n", blueFont);
                Chunk detailsPart4 = new Chunk("Adresse : "+ Adresse + "", blueFont);

                Phrase phrase = new Phrase();
                phrase.Add(titlePart1);
                phrase.Add(titlePart2);
                phrase.Add(detailsPart1);
                phrase.Add(detailsPart2);
                phrase.Add(detailsPart3);
                phrase.Add(detailsPart4);

                PdfPCell headerCell2 = new PdfPCell(phrase)
                {
                    Border = Rectangle.NO_BORDER,
                    HorizontalAlignment = Element.ALIGN_CENTER
                };


                Chunk titlePart1_InforADM = new Chunk(MoisDe.ToString() + "\n", greenFont);
                Chunk titlePart2_InforADM = new Chunk("C N°"+ NumQuitt + " \n\n", greenFont);
                Chunk detailsPart1_InforADM = new Chunk("N° : " + NumQuitt + "n\"", blueFont);
                Chunk detailsPart2_InforADM = new Chunk("Nom : " + NomLoc + "\n", blueFont);
                Chunk detailsPart3_InforADM = new Chunk("Code : " + CodeLoc + "\n", blueFont);
                Chunk detailsPart4_InforADM = new Chunk("Total : " + Totale + "", blueFont);

                Phrase phrase_InforADM = new Phrase();
                phrase_InforADM.Add(titlePart1_InforADM);
                phrase_InforADM.Add(titlePart2_InforADM);
                phrase_InforADM.Add(detailsPart1_InforADM);
                phrase_InforADM.Add(detailsPart3_InforADM);
                phrase_InforADM.Add(detailsPart2_InforADM);
                phrase_InforADM.Add(detailsPart4_InforADM);

                PdfPCell headerCell3 = new PdfPCell(phrase_InforADM)
                {
                    Border = Rectangle.NO_BORDER, // No border for this cell
                    HorizontalAlignment = Element.ALIGN_RIGHT // Align to the left (or right if needed)
                };

                // Add the cells to the table
                headerTable.AddCell(headerCell1);
                headerTable.AddCell(headerCell2);
                headerTable.AddCell(headerCell3);

                // Add the header table to the document
                document.Add(headerTable);

                // Add space after header table
                //document.Add(new Paragraph("\n"));

                #endregion

                // Main content table
                PdfPTable mainContentTable = new PdfPTable(8)
                {
                    WidthPercentage = 83, // Adjust to the page width
                    SpacingBefore = 10f, // Adds space before the table
                    HorizontalAlignment = Element.ALIGN_LEFT // Aligns the table to the left
                };

                mainContentTable.SpacingAfter = 0f; // No extra space after the table if needed
                mainContentTable.DefaultCell.PaddingLeft = 0f;

                // Define headers for main content table
                PdfPCell[] headers = {
                    new PdfPCell(new Phrase("Loyer", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER },
                    new PdfPCell(new Phrase("Charge", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER },
                    new PdfPCell(new Phrase("T.S.C", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER },
                    new PdfPCell(new Phrase("Totale", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER },
                    new PdfPCell(new Phrase("Mois de", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER },
                    new PdfPCell(new Phrase("Signature", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER },
                    new PdfPCell(new Phrase("Payé par", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER },
                    new PdfPCell(new Phrase("Usage", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER }
                };

                foreach (var headerCell in headers)
                {
                    headerCell.Padding = 5;
                    headerCell.BackgroundColor = BaseColor.LIGHT_GRAY;
                    mainContentTable.AddCell(headerCell);
                }

                mainContentTable.AddCell(new PdfPCell(new Phrase(Loyer.ToString("00.00"), boldFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                mainContentTable.AddCell(new PdfPCell(new Phrase(charge.ToString("00.00"), boldFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                mainContentTable.AddCell(new PdfPCell(new Phrase(Tsc.ToString("00.00"), boldFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                mainContentTable.AddCell(new PdfPCell(new Phrase(Totale.ToString("00.00"), boldFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                mainContentTable.AddCell(new PdfPCell(new Phrase(MoisDe.ToString(), boldFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                mainContentTable.AddCell(new PdfPCell(new Phrase("Le gérant", boldFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                mainContentTable.AddCell(new PdfPCell(new Phrase("Chèque N° :", boldFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                mainContentTable.AddCell(new PdfPCell(new Phrase(Usage, boldFont)) { HorizontalAlignment = Element.ALIGN_LEFT });

                // Add main content table to the document
                document.Add(mainContentTable);
                /*
                Paragraph rightSectionContent = new Paragraph("Additional Information", boldFont)
                {
                    Alignment = Element.ALIGN_LEFT, 
                    SpacingBefore = 0f,
                    IndentationLeft = separationPosition + 10 
                };
                
                document.Add(rightSectionContent);
                */
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
            finally
            {
                document.Close();
            }
        }

        private void fnValiderEncaissementParCoprop(string CodeGerance, string NomCoprop, string Adresse1, string Adresse2, string Adresse3)
        {
            List<string> ListCodeLocataires = new List<string>();
            List<string> ListNomLocataires = new List<string>();
            List<double> LitsLoyer = new List<double>();
            List<string> ListPerdioPay = new List<string>();
            List<string> ListCodeGer = new List<string>();
            List<string> ListPerdioImpy = new List<string>();
            List<double> ListMontantPay = new List<double>();
            List<double> ListMontantImp = new List<double>();
            List<string> ListeQuittance = new List<string>();
            List<string> ListeCopor = new List<string>();
            List<string> PartCopor = new List<string>();
            double commission = 0;
            double TVA = 0;
            double Total = 0;
            double Sold = 0;

            try
            {
                string selectQuery_loc_encs = "SELECT [N°], [code_loc], [nom_loc], [code_ger], [loyer], [periode_py], " +
                "[montant_py], [periode_impy], [montant_im] FROM [loc_encs] WHERE [code_ger] = '65'";

                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open();
                }

                using (OleDbCommand command_loc_encs = new OleDbCommand(selectQuery_loc_encs, _connection))
                {   
                    // Add the parameter in the order of placeholders
                    command_loc_encs.Parameters.Add(new OleDbParameter("?", CodeGerance));

                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(command_loc_encs))
                    {
                        DataTable dataTable = new DataTable();
                        adapter.Fill(dataTable);

                        if (dataTable.Rows.Count == 0)
                        {
                            MessageBox.Show("Merci de saisir l'encaissement pour ces copropriétaires avant de lancer la validation des encaissements !");
                            return;
                        }

                        foreach (DataRow row_ve in dataTable.Rows)
                        {
                            if (!ListCodeGer.Contains(row_ve["code_ger"].ToString()))
                            {
                                ListCodeGer.Add(row_ve["code_ger"].ToString());
                            }
                            ListeQuittance.Add(row_ve["N°"].ToString());
                            ListCodeLocataires.Add(row_ve["code_loc"].ToString());
                            ListNomLocataires.Add(row_ve["nom_loc"].ToString());
                            LitsLoyer.Add(Convert.ToDouble(row_ve["loyer"].ToString()));
                            ListPerdioPay.Add(row_ve["periode_py"].ToString());
                            ListPerdioImpy.Add(row_ve["periode_impy"].ToString());
                            ListMontantPay.Add(Convert.ToDouble(row_ve["montant_py"].ToString()));
                            ListMontantImp.Add(Convert.ToDouble(row_ve["montant_im"].ToString()));
                        }
                    }

                    GeneratePdfEncaissement(
                        $@"C:\Encaissement_{CodeGerance}_{NomCoprop}.pdf",
                        @"C:\Users\Pc\Downloads\detail_clinic_winform-main\Logo_AGI_page-0001-removebg-preview.png",
                         Adresse1,  Adresse2,  Adresse3,CodeGerance, NomCoprop,
                         ListeCopor, PartCopor,
                        ListCodeLocataires,
                        ListNomLocataires,
                        LitsLoyer,
                        ListPerdioPay,
                        ListPerdioImpy,
                        ListMontantPay,
                        ListMontantImp,
                        commission,
                        TVA,
                        Total,
                        Sold);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading Coproprietaire data: " + ex.Message);
            }
            finally
            {
                if (_connection.State == ConnectionState.Open)
                {
                    _connection.Close();
                }
            }
        }

        private void txtCodeLocataire_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                string selectedLocataire = txtCodeLocataire.SelectedValue?.ToString();

                if (string.IsNullOrEmpty(selectedLocataire))
                {
                    return;
                }

                string selectQuery = " SELECT [adresse], [CODE GER], [total] ,[nom loc], [datloc] FROM [loc] WHERE [code loc] = @CodeLocataire";

                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open(); // Open the connection
                }

                // Use OleDbCommand to fetch the data with parameterized query
                using (OleDbCommand command = new OleDbCommand(selectQuery, _connection))
                {
                    // Add parameter to the query to prevent SQL injection
                    command.Parameters.AddWithValue("@CodeLocataire", selectedLocataire);

                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                    {
                        // Fill a DataTable with the results
                        DataTable dataTable = new DataTable();

                        adapter.Fill(dataTable); // Fill the DataTable

                        if (dataTable.Rows.Count > 0)
                        {
                            // Assuming you want to show the first row's data
                            DataRow row = dataTable.Rows[0];

                            string test = Convert.ToDouble(row["total"]).ToString();

                            txtNomLocataire.Text =  row["nom loc"].ToString();
                            txtAdresseLocataire.Text = row["adresse"].ToString();
                            txtGerance.SelectedValue = row["CODE GER"].ToString();
                            txtMontantLoyer.Text = row["total"].ToString();
                            dateEmbauch.Value = Convert.ToDateTime(row["datloc"].ToString());
                        }
                        else
                        {
                            MessageBox.Show("Merci de vérifier si le nom et l'adresse de ce locataire est bien enregistrer au niveau du base de donnée !");
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

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            DataView dv = _dataTable.DefaultView;
            dv.RowFilter = string.Format("[Nom Locataire] LIKE '%{0}%'", txtRecherch1.Text.Trim());
            DataGridViewListLocataires.DataSource = dv;
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
                    dv.RowFilter = string.Format("", "");
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

        private void btnGenereQuittance_Click(object sender, EventArgs e)
        {   
            try 
            {   
                var selectedRows = new List<DataRow>();

                foreach (DataGridViewRow row in DataGridViewListLocataires.Rows)
                {
                    if (Convert.ToBoolean(row.Cells["Select"].Value) == true)
                    {
                        selectedRows.Add(((DataRowView)row.DataBoundItem).Row);
                    }
                }
                int i = 0;
                foreach (var row in selectedRows)
                {   
                    string Usage = "";

                    if(row["Usage"].ToString() == "H")
                    {
                        Usage = "Habitation";
                    }
                    else
                    {
                        Usage = "Comercial";
                    }

                    GenerateQuittances(QuittancePath +@"\Quittance_"+ row["Code Locataire"] + "_"+ row["Nom Locataire"] + ".pdf",
                    @"C:\Users\Pc\Downloads\detail_clinic_winform-main\Logo_AGI_page-0001-removebg-preview.png",
                    i.ToString(), row["Code Locataire"].ToString(), row["Nom Locataire"].ToString(),
                    row["Adresse"].ToString(), Convert.ToDouble(row["Total"].ToString()),
                    Convert.ToDouble(row["Charge"].ToString()), Convert.ToDouble(row["Taxe"].ToString()),
                    Convert.ToDouble(row["Total"].ToString()), txtPeriodPay.Text, Usage);
                    i++;
                    //Console.WriteLine($"Selected Row - Code Locataire: {row["Code Locataire"]}, Nom Locataire: {row["Nom Locataire"]}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("merci de selectionner au mois un locataire ! : Erreur" + ex.Message);
            }
        }

        private void btnValiderencaissement_Click(object sender, EventArgs e)
        {
            try
            {
                DataTable DTLocataire = new DataTable();

                string DateEncaissment = txtRechercheDateEncaissmentt.SelectedValue?.ToString();

                if (string.IsNullOrEmpty(DateEncaissment))
                {
                    MessageBox.Show("Veillez séléctionner une Date !");
                    return;
                }

                string selectQuery = @"SELECT [quittance], [date_encs], [code_ger], [code_loc], [def_loy], [intitulé], [total],
                                     [Nom_Locataire] FROM [Encaissement] 
                                     WHERE FORMAT([date_encs], 'yyyy/MM/dd') = FORMAT(@DateEncaissement, 'yyyy/MM/dd')";

                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open();
                }

                DateTime DateEncaissement = Convert.ToDateTime(DateEncaissment);

                using (OleDbCommand command = new OleDbCommand(selectQuery, _connection))
                {   
                    command.Parameters.AddWithValue("@DateEncaissement", DateEncaissement);

                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                    {
                        adapter.Fill(DTLocataire); 
                    }
                }

                GeneratePdfEtatJournalier(DTLocataire, DateEncaissement);
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

        private string getLast2Yers(int year)
        {
            string lastTwoDigitsString = "";
            try
            {
                string yearString = year.ToString();
                lastTwoDigitsString = yearString.Substring(yearString.Length - 2);
            }
            catch (Exception ex)
            {

            }
            return lastTwoDigitsString;
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

        public void GeneratePdfEtatJournalier(DataTable DTLocataire, DateTime Dt_Encaissement)
        {   
            string filePath = EtatDetaillePath + $@"\EtatDetaille_"+Dt_Encaissement.ToString("yyyy-MM-dd")+".pdf";
            string imagePath = @"Logo_AGI_page-0001-removebg-preview.png";


            var CodeGerData = DTLocataire.AsEnumerable()
                             .GroupBy(row => row.Field<int>("code_ger")) // Utilisation de int
                             .Select(g => new
                             {
                                 CodeGer = g.Key.ToString() // Conversion en string si nécessaire
                             });

            iTextSharp.text.Font headerFont = FontFactory.GetFont(FontFactory.COURIER_BOLD, 10);
            iTextSharp.text.Font regularFont = FontFactory.GetFont(FontFactory.COURIER, 10);
            iTextSharp.text.Font boldFont = FontFactory.GetFont(FontFactory.COURIER, 8);

            // Create a new document
            Document document = new Document(PageSize.A4, 30, 30, 10, 10);

            try
            {
                #region Entete ------------------------------------------------------------------------------------- 
                // Set up the PDF writer
                PdfWriter.GetInstance(document, new FileStream(filePath, FileMode.Create));
                document.Open();

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

                Paragraph closingMessage = new Paragraph("Date Encaissement : " +Dt_Encaissement.ToString("yyyy/MM/dd"), headerFont);
                closingMessage.Alignment = Element.ALIGN_LEFT;
                closingMessage.SpacingAfter = 15;
                document.Add(closingMessage);

                // Create a table with 6 columns for "Code Loc", "Nom Locataire", "Loyer", etc.
                PdfPTable table = new PdfPTable(5);
                table.WidthPercentage = 100;
                table.SetWidths(new float[] { 2, 3, 2, 2, 2 });

                // Add table headers
                PdfPCell[] headers = {
                        new PdfPCell(new Phrase("Code Locataire", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER },
                        new PdfPCell(new Phrase("Nom Locataire", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER },
                        new PdfPCell(new Phrase("Quittance", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER },
                        new PdfPCell(new Phrase("Mois payé", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER },
                        new PdfPCell(new Phrase("Montant payé", boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER },
                    };

                foreach (var headerCell in headers)
                {
                    headerCell.Padding = 5;
                    headerCell.BackgroundColor = BaseColor.LIGHT_GRAY;
                    table.AddCell(headerCell);
                }

                document.Add(table);

                double Valtotal = 0;

                foreach (var GerCode in CodeGerData)
                {
                    PdfPTable table_Etat = new PdfPTable(5);
                    table_Etat.WidthPercentage = 100;
                    table_Etat.SetWidths(new float[] { 2, 3, 2, 2, 2 });

                    Valtotal = 0;

                    Paragraph closingCodeGer = new Paragraph("Code Gerance : " + GerCode.CodeGer.ToString() , headerFont);
                    closingCodeGer.Alignment = Element.ALIGN_LEFT;
                    closingCodeGer.SpacingAfter = 6;
                    document.Add(closingCodeGer);

                    for (int i = 0; i < DTLocataire.Rows.Count; i++)
                    {
                        if (DTLocataire.Rows[i]["code_ger"].ToString() == GerCode.CodeGer)
                        {
                            //[quittance], [date_encs], [code_ger], [code_loc], [def_loy], [intitulé], [total], [Nom_Locataire]

                            table_Etat.AddCell(new PdfPCell(new Phrase(DTLocataire.Rows[i]["code_loc"].ToString(), boldFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                            table_Etat.AddCell(new PdfPCell(new Phrase(DTLocataire.Rows[i]["Nom_Locataire"].ToString(), boldFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                            table_Etat.AddCell(new PdfPCell(new Phrase(DTLocataire.Rows[i]["quittance"].ToString(), boldFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                            table_Etat.AddCell(new PdfPCell(new Phrase(DTLocataire.Rows[i]["intitulé"].ToString()+"/"+ getLast2Yers(Dt_Encaissement.Year), boldFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                            table_Etat.AddCell(new PdfPCell(new Phrase(Convert.ToDouble(DTLocataire.Rows[i]["total"].ToString()).ToString("00.00"), boldFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                            Valtotal = Valtotal + Convert.ToDouble(DTLocataire.Rows[i]["total"].ToString());
                        }
                    }

                    document.Add(table_Etat);

                    #region Solde ---------------------------------------------------------------------------------------
                    // Code to add "Solde" table below the "Dépense" table
                    PdfPTable soldeTable = new PdfPTable(1);
                    soldeTable.WidthPercentage = 50;
                    soldeTable.HorizontalAlignment = Element.ALIGN_RIGHT; // Aligns the table to the right

                    PdfPCell soldeLabelCell = new PdfPCell(new Phrase("Total : "+ Valtotal.ToString("00.00"), FontFactory.GetFont(FontFactory.COURIER_BOLD, 10)))
                    {
                        Border = iTextSharp.text.Rectangle.NO_BORDER,
                        HorizontalAlignment = Element.ALIGN_RIGHT
                    };

                    soldeTable.AddCell(soldeLabelCell);

                    document.Add(soldeTable);
                    #endregion

                }

                double TotalPay = CalculateTotalSum(DTLocataire, "total");
                // Code to add "Solde" table below the "Dépense" table
                PdfPTable soldeTableTotale = new PdfPTable(1);
                soldeTableTotale.WidthPercentage = 50;
                soldeTableTotale.SpacingBefore = 20;
                soldeTableTotale.HorizontalAlignment = Element.ALIGN_RIGHT; // Aligns the table to the right

                PdfPCell soldeLabelCellTotale = new PdfPCell(new Phrase("Montant Total : "+ TotalPay.ToString("00.00"), FontFactory.GetFont(FontFactory.COURIER_BOLD, 11)))
                {
                    Border = iTextSharp.text.Rectangle.NO_BORDER,
                    HorizontalAlignment = Element.ALIGN_RIGHT
                };

                soldeTableTotale.AddCell(soldeLabelCellTotale);

                // Add the "Solde" table to the document
                document.Add(soldeTableTotale);

                ShowFile(filePath);

                #endregion
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
            finally
            {
                document.Close();
            }
        }

        private double CalculateTotalSum(DataTable DTLocataire, string ColumnName)
        {
            double totalSum = 0;

            try
            {

                foreach (DataRow row in DTLocataire.Rows)
                {
                    string listMontantPay = row[ColumnName].ToString();
                    if (string.IsNullOrWhiteSpace(listMontantPay))
                    {
                        continue;
                    }
                    string[] montantArray = listMontantPay.Split(',');

                    totalSum += montantArray
                        .Select(montant => double.TryParse(montant.Trim(), out double value) ? value : 0)
                        .Sum();
                }
            }
            catch (Exception ex)
            {

            }
            return totalSum;
        }


        private void txtQuittance_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(txtQuittance.Text))
                {
                    textShowError.Text = string.Empty;
                    txtQuittance.BackColor = SystemColors.Window;
                    _dataTable_Encaissement.DefaultView.RowFilter = string.Empty;
                    DataGridViewListEncaissement.DataSource = _dataTable_Encaissement.DefaultView;
                    return;
                }

                DataView dv = _dataTable_Encaissement.DefaultView;

                // Convert the numeric column to string for filtering
                dv.RowFilter = string.Format("CONVERT([Quittance], 'System.String') LIKE '%{0}%'", txtQuittance.Text.Trim());

                if (dv.Count > 0)
                {
                    textShowError.Text = "Code Quittance déjà existe !";
                    txtQuittance.BackColor = Color.Red;
                }
                else
                {
                    textShowError.Text = string.Empty;
                    txtQuittance.BackColor = SystemColors.Window;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur lors de la recherche : " + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void txtRecherch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                DataView dv = _dataTable_Encaissement.DefaultView;
                dv.RowFilter = string.Format("CONVERT([Quittance], 'System.String') LIKE '%{0}%'", txtRecherch.Text.Trim());
                DataGridViewListEncaissement.DataSource = dv;
            }
            catch(Exception ex)
            {
                MessageBox.Show("Erreur de recherche : " + ex.Message);
            }
        }

        private void Supprimer_Click(object sender, EventArgs e)
        {
            if (DataGridViewListEncaissement.SelectedRows.Count > 0)
            {
                try
                {
                    DataGridViewRow selectedRow = DataGridViewListEncaissement.SelectedRows[0];

                    string NumEncaiss = selectedRow.Cells["Numero Encaissement"].Value.ToString();

                    DataGridViewListEncaissement.Rows.Remove(selectedRow);

                    string deleteQuery = "DELETE FROM Encaissement WHERE [N] = @NumEncaiss";

                    using (OleDbCommand cmd = new OleDbCommand(deleteQuery, _connection))
                    {
                        cmd.Parameters.AddWithValue("@NumEncaiss", NumEncaiss);

                        if (_connection.State != ConnectionState.Open)
                        {
                            _connection.Open();
                        }

                        int rowsAffected = cmd.ExecuteNonQuery();
                        if (rowsAffected > 0)
                        {
                            MessageBox.Show("     Encaissement a été supprimé avec succès.     ");
                        }
                        else
                        {
                            MessageBox.Show("     Erreur lors de la suppression d'encaissement.      ");
                        }
                    }

                    LoadData();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("        Erreur du supprission dans la ligne : " + ex.Message);
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

        private void MAJtxtQuitance_TextChanged(object sender, EventArgs e)
        {

        }

        private void MAJtxtQuitance_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) // Si l'utilisateur appuie sur Entrée
            {
                ChargerDonneesLocataireParCode(); // Appelle votre méthode pour charger les données
            }
        }

        private void MAJtxtQuitance_Leave(object sender, EventArgs e)
        {
            // Appelé lorsque le contrôle perd le focus
            ChargerDonneesLocataireParCode();
        }
    }
}
