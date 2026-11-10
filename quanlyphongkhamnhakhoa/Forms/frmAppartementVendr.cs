using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms; 

namespace PKNK_CNPM.Forms
{
    public partial class FormLocaleVendr : Form
    {
        private OleDbConnection _connection;
        private OleDbDataAdapter _dataAdapter;
        private DataTable _dataTable;
        private string AffairesPath = "";

        public FormLocaleVendr()
        {
            InitializeComponent();
            this.Icon = new Icon("AGI.ico");
            InitializeDatabase();
            LoadData();

            LoadLocaux();
            LoadVille();
            LoadArrondissement();
            LoadQuarttier();

            LoadEtats();
            LoadTypeSol();
            LoadSalleDeBains();
            LoadChambres();
            LoadCarateristiques();
            LoadInterieur();
            LoadOptionSupp();
            FnGetAffairesPhotosPath();
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

                _dataAdapter = new OleDbDataAdapter("SELECT [Id_Appartement] AS [Code Locale], [Etat], [Arrondissement], " +
                               "[Quartier], [Prix] FROM Appartement  WHERE [Type_Appartement] = 'V' ORDER BY [Id_Appartement]", _connection);

                OleDbCommandBuilder commandBuilder = new OleDbCommandBuilder(_dataAdapter);

                _dataTable = new DataTable();
                _dataAdapter.Fill(_dataTable);

                if (_dataTable.Columns.Contains("Prix"))
                {
                    // Add a new column for the formatted Prix
                    _dataTable.Columns.Add("Formatted Prix", typeof(string));

                    foreach (DataRow row in _dataTable.Rows)
                    {
                        if (double.TryParse(row["Prix"]?.ToString(), out double PrixValue))
                        {
                            row["Formatted Prix"] = $"{PrixValue:0.00} DH"; // Format with 2 decimal places
                        }
                        else
                        {
                            row["Formatted Prix"] = "Invalid Value"; // Handle non-numeric cases
                        }
                    }

                    // Optional: Remove the original 'Prix' column if not needed
                    _dataTable.Columns.Remove("Prix");
                    _dataTable.Columns["Formatted Prix"].ColumnName = "Prix";
                }

                DataGridViewListeLocaux.DataSource = _dataTable;
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

        private void LoadEtats()
        {
            try
            {
                // Adjust the SQL query to match the exact table and column names
                string selectQueryArrondissement = "SELECT [Details] FROM [DetailsLocale] WHERE Table = 'Etat'";

                // Open the connection
                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open(); // Open the connection
                }

                // Fetch the data for ger
                using (OleDbDataAdapter adapterArrondissement = new OleDbDataAdapter(selectQueryArrondissement, _connection))
                {
                    DataTable dataTableArrondissement = new DataTable();
                    adapterArrondissement.Fill(dataTableArrondissement);

                    DataRow emptyRow = dataTableArrondissement.NewRow();
                    emptyRow["Details"] = DBNull.Value; // or string.Empty if you prefer empty strings
                    emptyRow["Details"] = string.Empty; // or set a placeholder text like "<Select>"
                    dataTableArrondissement.Rows.InsertAt(emptyRow, 0);

                    ListEtatLocale.DisplayMember = "Details"; // Display this column in ComboBox
                    ListEtatLocale.ValueMember = "Details"; // Use this column as the value
                    ListEtatLocale.DataSource = dataTableArrondissement;


                    MAJListeEtat.DisplayMember = "Details"; // Display this column in ComboBox
                    MAJListeEtat.ValueMember = "Details"; // Use this column as the value
                    MAJListeEtat.DataSource = dataTableArrondissement;
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
        private void FnGetAffairesPhotosPath()
        {
            try
            {
                // Adjust the SQL query to match the exact table and column names
                string selectQuery = "SELECT [Valeur] FROM [Params] WHERE [Id Param] = 'Affaires'";

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

                        if (dataTable.Rows.Count > 0)
                        {
                            DataRow row = dataTable.Rows[0];

                            AffairesPath = row["Valeur"].ToString();
                        }
                        else
                        {
                            MessageBox.Show("Aucune Path Séléctionné merci d'ajouter un path dans l'écran Paramtres avancé !");
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

        private void LoadLocaux()
        {
            try
            {
                // Adjust the SQL query to match the exact table and column names
                string selectQueryArrondissement = "SELECT [Id_Appartement] FROM [Appartement] WHERE [Type_Appartement] = 'V' ORDER BY [Id_Appartement]";

                // Open the connection
                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open(); // Open the connection
                }

                // Fetch the data for ger
                using (OleDbDataAdapter adapterArrondissement = new OleDbDataAdapter(selectQueryArrondissement, _connection))
                {
                    DataTable dataTableArrondissement = new DataTable();
                    adapterArrondissement.Fill(dataTableArrondissement);

                    DataRow emptyRow = dataTableArrondissement.NewRow();
                    emptyRow["Id_Appartement"] = DBNull.Value; // or string.Empty if you prefer empty strings
                    dataTableArrondissement.Rows.InsertAt(emptyRow, 0);

                    MAJListeCodeLocale.DisplayMember = "Id_Appartement"; // Display this column in ComboBox
                    MAJListeCodeLocale.ValueMember = "Id_Appartement"; // Use this column as the value
                    MAJListeCodeLocale.DataSource = dataTableArrondissement;

                    ListeCodeLocaleDetails.DisplayMember = "Id_Appartement"; // Display this column in ComboBox
                    ListeCodeLocaleDetails.ValueMember = "Id_Appartement"; // Use this column as the value
                    ListeCodeLocaleDetails.DataSource = dataTableArrondissement;

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
        private void LoadVille()
        {
            try
            {
                // Adjust the SQL query to match the exact table and column names
                string selectQueryArrondissement = "SELECT [Ville] FROM [Quartiers] GROUP BY [Ville]";

                // Open the connection
                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open(); // Open the connection
                }

                // Fetch the data for ger
                using (OleDbDataAdapter adapterArrondissement = new OleDbDataAdapter(selectQueryArrondissement, _connection))
                {
                    DataTable dataTableArrondissement = new DataTable();
                    adapterArrondissement.Fill(dataTableArrondissement);

                    DataRow emptyRow = dataTableArrondissement.NewRow();
                    emptyRow["Ville"] = DBNull.Value; // or string.Empty if you prefer empty strings
                    emptyRow["Ville"] = string.Empty; // or set a placeholder text like "<Select>"
                    dataTableArrondissement.Rows.InsertAt(emptyRow, 0);

                    ListeVille.DisplayMember = "Ville"; // Display this column in ComboBox
                    ListeVille.ValueMember = "Ville"; // Use this column as the value
                    ListeVille.DataSource = dataTableArrondissement;


                    MAJListeVille.DisplayMember = "Ville"; // Display this column in ComboBox
                    MAJListeVille.ValueMember = "Ville"; // Use this column as the value
                    MAJListeVille.DataSource = dataTableArrondissement;
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
        private void LoadArrondissement()
        {
            try
            {
                string selectedVille = ListeVille.SelectedValue?.ToString();

                string selectQueryArrondissement = "SELECT [Arrondissement] FROM [Quartiers] GROUP BY [Arrondissement]";

                if (!string.IsNullOrEmpty(selectedVille))
                {
                    selectQueryArrondissement = "SELECT [Arrondissement] FROM [Quartiers] WHERE Ville = '" + selectedVille + "' GROUP BY [Arrondissement]";
                }

                // Adjust the SQL query to match the exact table and column names

                // Open the connection
                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open(); // Open the connection
                }

                // Fetch the data for ger
                using (OleDbDataAdapter adapterArrondissement = new OleDbDataAdapter(selectQueryArrondissement, _connection))
                {
                    DataTable dataTableArrondissement = new DataTable();
                    adapterArrondissement.Fill(dataTableArrondissement);

                    DataRow emptyRow = dataTableArrondissement.NewRow();
                    emptyRow["Arrondissement"] = DBNull.Value; // or string.Empty if you prefer empty strings
                    emptyRow["Arrondissement"] = string.Empty; // or set a placeholder text like "<Select>"
                    dataTableArrondissement.Rows.InsertAt(emptyRow, 0);

                    ListArrondissement.DisplayMember = "Arrondissement"; // Display this column in ComboBox
                    ListArrondissement.ValueMember = "Arrondissement"; // Use this column as the value
                    ListArrondissement.DataSource = dataTableArrondissement;

                    MAJListArrondissement.DisplayMember = "Arrondissement"; // Display this column in ComboBox
                    MAJListArrondissement.ValueMember = "Arrondissement"; // Use this column as the value
                    MAJListArrondissement.DataSource = dataTableArrondissement;
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
        private void LoadQuarttier()
        {
            try
            {
                string selectedVille = ListArrondissement.SelectedValue?.ToString();

                string selectQuery = "SELECT [Quartier] FROM [Quartiers]";

                if (!string.IsNullOrEmpty(selectedVille))
                {
                    selectQuery = "SELECT [Quartier] FROM [Quartiers] WHERE Arrondissement = '" + selectedVille + "'";
                }

                // Adjust the SQL query to match the exact table and column names

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
                    emptyRow["Quartier"] = DBNull.Value; // or string.Empty if you prefer empty strings
                    emptyRow["Quartier"] = string.Empty; // or set a placeholder text like "<Select>"
                    dataTable.Rows.InsertAt(emptyRow, 0);

                    ListQuertier.DisplayMember = "Quartier"; // Display this column in ComboBox
                    ListQuertier.ValueMember = "Quartier"; // Use this column as the value
                    ListQuertier.DataSource = dataTable;

                    MAJListQuertier.DisplayMember = "Quartier"; // Display this column in ComboBox
                    MAJListQuertier.ValueMember = "Quartier"; // Use this column as the value
                    MAJListQuertier.DataSource = dataTable;
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
        private void LoadTypeSol()
        {
            try
            {
                // Adjust the SQL query to match the exact table and column names
                string selectQueryArrondissement = "SELECT [Details] FROM [DetailsLocale] WHERE Table = 'TypeSol'";

                // Open the connection
                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open(); // Open the connection
                }

                // Fetch the data for ger
                using (OleDbDataAdapter adapterArrondissement = new OleDbDataAdapter(selectQueryArrondissement, _connection))
                {
                    DataTable dataTableArrondissement = new DataTable();
                    adapterArrondissement.Fill(dataTableArrondissement);

                    DataRow emptyRow = dataTableArrondissement.NewRow();
                    emptyRow["Details"] = DBNull.Value; // or string.Empty if you prefer empty strings
                    emptyRow["Details"] = string.Empty; // or set a placeholder text like "<Select>"
                    dataTableArrondissement.Rows.InsertAt(emptyRow, 0);

                    ListTypeSol.DisplayMember = "Details"; // Display this column in ComboBox
                    ListTypeSol.ValueMember = "Details"; // Use this column as the value
                    ListTypeSol.DataSource = dataTableArrondissement;


                    MAJListeTypeSol.DisplayMember = "Details"; // Display this column in ComboBox
                    MAJListeTypeSol.ValueMember = "Details"; // Use this column as the value
                    MAJListeTypeSol.DataSource = dataTableArrondissement;
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
        private void LoadSalleDeBains()
        {
            try
            {
                // Adjust the SQL query to match the exact table and column names
                string selectQueryArrondissement = "SELECT [Details] FROM [DetailsLocale] WHERE Table = 'SalleBains'";

                // Open the connection
                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open(); // Open the connection
                }

                // Fetch the data for ger
                using (OleDbDataAdapter adapterArrondissement = new OleDbDataAdapter(selectQueryArrondissement, _connection))
                {
                    DataTable dataTableArrondissement = new DataTable();
                    adapterArrondissement.Fill(dataTableArrondissement);

                    DataRow emptyRow = dataTableArrondissement.NewRow();
                    emptyRow["Details"] = DBNull.Value; // or string.Empty if you prefer empty strings
                    emptyRow["Details"] = string.Empty; // or set a placeholder text like "<Select>"
                    dataTableArrondissement.Rows.InsertAt(emptyRow, 0);

                    ListSalleBains.DisplayMember = "Details"; // Display this column in ComboBox
                    ListSalleBains.ValueMember = "Details"; // Use this column as the value
                    ListSalleBains.DataSource = dataTableArrondissement;

                    MAJListeSalleBains.DisplayMember = "Details"; // Display this column in ComboBox
                    MAJListeSalleBains.ValueMember = "Details"; // Use this column as the value
                    MAJListeSalleBains.DataSource = dataTableArrondissement;
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

        private void LoadChambres()
        {
            try
            {
                // Adjust the SQL query to match the exact table and column names
                string selectQueryArrondissement = "SELECT [Details] FROM [DetailsLocale] WHERE Table = 'Chambres'";

                // Open the connection
                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open(); // Open the connection
                }

                // Fetch the data for ger
                using (OleDbDataAdapter adapterArrondissement = new OleDbDataAdapter(selectQueryArrondissement, _connection))
                {
                    DataTable dataTableArrondissement = new DataTable();
                    adapterArrondissement.Fill(dataTableArrondissement);

                    DataRow emptyRow = dataTableArrondissement.NewRow();
                    emptyRow["Details"] = DBNull.Value; // or string.Empty if you prefer empty strings
                    emptyRow["Details"] = string.Empty; // or set a placeholder text like "<Select>"
                    dataTableArrondissement.Rows.InsertAt(emptyRow, 0);

                    ListChambres.DisplayMember = "Details"; // Display this column in ComboBox
                    ListChambres.ValueMember = "Details"; // Use this column as the value
                    ListChambres.DataSource = dataTableArrondissement;

                    MAJListeChambres.DisplayMember = "Details"; // Display this column in ComboBox
                    MAJListeChambres.ValueMember = "Details"; // Use this column as the value
                    MAJListeChambres.DataSource = dataTableArrondissement;
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
        private void LoadCarateristiques()
        {
            try
            {
                // Adjust the SQL query to match the exact table and column names
                string selectQueryArrondissement = "SELECT [Details] FROM [DetailsLocale] WHERE Table = 'Caracteriqtiques'";

                // Open the connection
                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open(); // Open the connection
                }

                // Fetch the data for ger
                using (OleDbDataAdapter adapterArrondissement = new OleDbDataAdapter(selectQueryArrondissement, _connection))
                {
                    DataTable dataTableArrondissement = new DataTable();
                    adapterArrondissement.Fill(dataTableArrondissement);

                    DataRow emptyRow = dataTableArrondissement.NewRow();
                    emptyRow["Details"] = DBNull.Value; // or string.Empty if you prefer empty strings
                    emptyRow["Details"] = string.Empty; // or set a placeholder text like "<Select>"
                    dataTableArrondissement.Rows.InsertAt(emptyRow, 0);

                    ListCaracteristiques.DisplayMember = "Details"; // Display this column in ComboBox
                    ListCaracteristiques.ValueMember = "Details"; // Use this column as the value
                    ListCaracteristiques.DataSource = dataTableArrondissement;

                    MAJListeCaracteriqtiques.DisplayMember = "Details"; // Display this column in ComboBox
                    MAJListeCaracteriqtiques.ValueMember = "Details"; // Use this column as the value
                    MAJListeCaracteriqtiques.DataSource = dataTableArrondissement;
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
        private void LoadInterieur()
        {
            try
            {
                // Adjust the SQL query to match the exact table and column names
                string selectQueryArrondissement = "SELECT [Details] FROM [DetailsLocale] WHERE Table = 'Interieur'";

                // Open the connection
                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open(); // Open the connection
                }

                // Fetch the data for ger
                using (OleDbDataAdapter adapterArrondissement = new OleDbDataAdapter(selectQueryArrondissement, _connection))
                {
                    DataTable dataTableArrondissement = new DataTable();
                    adapterArrondissement.Fill(dataTableArrondissement);

                    DataRow emptyRow = dataTableArrondissement.NewRow();
                    emptyRow["Details"] = DBNull.Value; // or string.Empty if you prefer empty strings
                    emptyRow["Details"] = string.Empty; // or set a placeholder text like "<Select>"
                    dataTableArrondissement.Rows.InsertAt(emptyRow, 0);

                    ListInterieur.DisplayMember = "Details"; // Display this column in ComboBox
                    ListInterieur.ValueMember = "Details"; // Use this column as the value
                    ListInterieur.DataSource = dataTableArrondissement;

                    MAJListeInterieur.DisplayMember = "Details"; // Display this column in ComboBox
                    MAJListeInterieur.ValueMember = "Details"; // Use this column as the value
                    MAJListeInterieur.DataSource = dataTableArrondissement;
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
        private void LoadOptionSupp()
        {
            try
            {
                // Adjust the SQL query to match the exact table and column names
                string selectQueryArrondissement = "SELECT [Details] FROM [DetailsLocale] WHERE Table = 'OptionSupp'";

                // Open the connection
                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open(); // Open the connection
                }

                // Fetch the data for ger
                using (OleDbDataAdapter adapterArrondissement = new OleDbDataAdapter(selectQueryArrondissement, _connection))
                {
                    DataTable dataTableArrondissement = new DataTable();
                    adapterArrondissement.Fill(dataTableArrondissement);

                    DataRow emptyRow = dataTableArrondissement.NewRow();
                    emptyRow["Details"] = DBNull.Value; // or string.Empty if you prefer empty strings
                    emptyRow["Details"] = string.Empty; // or set a placeholder text like "<Select>"
                    dataTableArrondissement.Rows.InsertAt(emptyRow, 0);

                    ListOptionSupp.DisplayMember = "Details"; // Display this column in ComboBox
                    ListOptionSupp.ValueMember = "Details"; // Use this column as the value
                    ListOptionSupp.DataSource = dataTableArrondissement;

                    MAJListeOptionSupp.DisplayMember = "Details"; // Display this column in ComboBox
                    MAJListeOptionSupp.ValueMember = "Details"; // Use this column as the value
                    MAJListeOptionSupp.DataSource = dataTableArrondissement;
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


        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (DataGridViewListeLocaux.SelectedRows.Count > 0)
            {
                try
                {
                    DataGridViewRow selectedRow = DataGridViewListeLocaux.SelectedRows[0];

                    string CodeLocale = selectedRow.Cells["Code Locale"].Value.ToString();

                    DataGridViewListeLocaux.Rows.Remove(selectedRow);

                    string deleteQuery = "DELETE FROM Appartement WHERE [Id_Appartement] = @CodeLocale";

                    using (OleDbCommand cmd = new OleDbCommand(deleteQuery, _connection))
                    {
                        cmd.Parameters.AddWithValue("@CodeLocale", CodeLocale);

                        if (_connection.State != ConnectionState.Open)
                        {
                            _connection.Open();
                        }

                        int rowsAffected = cmd.ExecuteNonQuery();
                        if (rowsAffected > 0)
                        {
                            MessageBox.Show("Locale a été supprimé avec succès.");
                        }
                        else
                        {
                            MessageBox.Show("Erreur lors de la suppression du Locale.");
                        }
                    }

                    // Refresh the DataGridView after deletion
                    LoadData();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erreur lors de la suppression du Locale.");
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
                MessageBox.Show("Veillez séléctionner une ligne");
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

            if (DataGridViewListeLocaux.SelectedRows.Count > 0)
            {
                DataGridViewRow selectedRow = DataGridViewListeLocaux.SelectedRows[0];
                selectedRow.Cells["code_ger"].Value = ListEtatLocale.Text;
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


        /****************************************************************************************************/

        private void FormLocaleVendr_Load(object sender, EventArgs e)
        {
        }
      

        private void btnKhamLamSan_Click(object sender, EventArgs e)
        {

        }

        private void btnCanLamSan_Click(object sender, EventArgs e)
        {

        }

        private void MAJCodeLocatire_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void txtRecherch_TextChanged(object sender, EventArgs e)
        {

        }

        private void ListeVille_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadArrondissement();
        }

        private void ListArrondissement_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadQuarttier();
        }

        private void MAJListeCodeLocale_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                // Get the selected value from the ComboBox, assuming the ValueMember is set to "code loc"
                string selectedCodeLoc = MAJListeCodeLocale.SelectedValue?.ToString();

                if (string.IsNullOrEmpty(selectedCodeLoc))
                {
                    return;
                }

                string selectQuery = "SELECT [Arrondissement] ,[Etat] , [Quartier] ,[Ville] ," +
                "[AdresseComplete] , [TypeSol] , [EtageBien] , [Surface] ," +
                "[Prix] , [SallesBains] , [Chambres] , [CaractqGene] ," +
                "[Interieur] , [OptionSupp] , [Description] FROM Appartement " +
                " WHERE [Id_Appartement] = @MAJtxtCodeLocale";

                // Check if the connection is not already open
                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open(); // Open the connection
                }

                // Use OleDbCommand to fetch the data with parameterized query
                using (OleDbCommand command = new OleDbCommand(selectQuery, _connection))
                {
                    // Add parameter to the query to prevent SQL injection
                    command.Parameters.AddWithValue("Id_Appartement", selectedCodeLoc);

                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                    {
                        // Fill a DataTable with the results
                        DataTable dataTable = new DataTable();

                        adapter.Fill(dataTable); // Fill the DataTable

                        if (dataTable.Rows.Count > 0)
                        {
                            // Assuming you want to show the first row's data
                            DataRow row = dataTable.Rows[0];

                            MAJListeEtat.SelectedValue = row["Etat"].ToString();
                            MAJListeVille.SelectedValue = row["Ville"].ToString();
                            MAJListArrondissement.SelectedValue = row["Arrondissement"].ToString();
                            MAJListQuertier.SelectedValue = row["Quartier"].ToString();
                            MAJtxtAdresseComplet.Text = row["AdresseComplete"].ToString();
                            MAJListeTypeSol.SelectedValue = row["TypeSol"].ToString();
                            MAJListeEtageBien.SelectedValue = row["EtageBien"].ToString();

                            MAJtxtSurface.Text = row["Surface"].ToString();
                            MAJtxtPrix.Text = row["Prix"].ToString();
                            MAJListeSalleBains.SelectedValue = row["SallesBains"].ToString();
                            MAJListeChambres.SelectedValue = row["Chambres"].ToString();
                            MAJListeCaracteriqtiques.SelectedValue = row["CaractqGene"].ToString();
                            MAJListeInterieur.SelectedValue = row["Interieur"].ToString();
                            MAJListeOptionSupp.SelectedValue = row["OptionSupp"].ToString();
                            MAJtxtDescription.Text = row["Description"].ToString();
                            //MAJLoyer.Text = row["CODE GER"].ToString();
                        }
                        else
                        {
                            MessageBox.Show("No data found for the Locale Bien.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading Locale data: " + ex.Message);
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

        private void PictureBox_Click(object sender, EventArgs e)
        {
            var pictureBox = sender as PictureBox;

            if (pictureBox != null)
            {
                string imagePath = pictureBox.Tag.ToString();

                try
                {
                    using (var fs = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
                    using (var ms = new MemoryStream())
                    {
                        fs.CopyTo(ms);
                        ms.Seek(0, SeekOrigin.Begin);

                        // Validate and load the image
                        using (var img = Image.FromStream(ms))
                        {
                            // Set the image as the background of the panel
                            mainPanel.BackgroundImage = (Image)img.Clone();
                            mainPanel.BackgroundImageLayout = ImageLayout.Stretch; // Make the image fill 100% of the panel
                        }
                    }
                }
                catch (OutOfMemoryException)
                {
                    MessageBox.Show($"The file is not a valid image or is too large: {imagePath}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"An error occurred while displaying {imagePath}: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnCanLamSan_Click_2(object sender, EventArgs e)
        {
            string MAX_ID_Credit = "1";

            try
            {
                string Query_MAX_ID_Credit = "SELECT MAX([N°]) AS N° FROM Appartement";

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
                            // MessageBox.Show("Impossible de récupérer le MAX N° Locale !!");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading Locale data: " + ex.Message);
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
                string txtCodeLocale_ = String.IsNullOrWhiteSpace(txtCodeLocale.Text) ? String.Empty : txtCodeLocale.Text;
                string ListEtatLocale_ = ListEtatLocale.SelectedValue != null ? ListEtatLocale.SelectedValue.ToString() : String.Empty;
                string ListArrondissement_ = ListArrondissement.SelectedValue != null ? ListArrondissement.SelectedValue.ToString() : String.Empty;
                string ListQuertier_ = ListQuertier.SelectedValue != null ? ListQuertier.SelectedValue.ToString() : String.Empty;
                string ListTypeSol_ = ListTypeSol.SelectedValue != null ? ListTypeSol.SelectedValue.ToString() : String.Empty;
                string ListEtage_ = ListEtage.SelectedValue != null ? ListEtage.SelectedValue.ToString() : String.Empty;
                string ListSalleBains_ = ListSalleBains.SelectedValue != null ? ListSalleBains.SelectedValue.ToString() : String.Empty;
                string ListChambres_ = ListChambres.SelectedValue != null ? ListChambres.SelectedValue.ToString() : String.Empty;
                string ListCaracteristiques_ = ListCaracteristiques.SelectedValue != null ? ListCaracteristiques.SelectedValue.ToString() : String.Empty;
                string ListInterieur_ = ListInterieur.SelectedValue != null ? ListInterieur.SelectedValue.ToString() : String.Empty;
                string ListOptionSupp_ = ListOptionSupp.SelectedValue != null ? ListOptionSupp.SelectedValue.ToString() : String.Empty;
                string txtAdresse_ = String.IsNullOrWhiteSpace(txtAdresse.Text) ? String.Empty : txtAdresse.Text;
                string txtDescription_ = String.IsNullOrWhiteSpace(txtDescription.Text) ? String.Empty : txtDescription.Text;
                int txtSurface_ = String.IsNullOrWhiteSpace(txtSurface.Text) ? 0 : Convert.ToInt32(txtSurface.Text);
                int txtPrix_ = String.IsNullOrWhiteSpace(txtPrix.Text) ? 0 : Convert.ToInt32(txtPrix.Text);



                string insertQuery = "INSERT INTO Appartement ([N°], [Id_Appartement], [Etat], [Arrondissement], [Quartier]," +
                "[AdresseComplete], [TypeSol], [EtageBien], [Surface], [Prix], [SallesBains], [Chambres], [CaractqGene]," +
                "[Interieur], [OptionSupp], [Description]) " +
                " VALUES (@N, @txtCodeLocale, @ListEtatLocale, @ListArrondissement, @ListQuertier," +
                "@txtAdresse, @ListTypeSol, @ListEtage, @txtSurface, @txtPrix, @ListSalleBains, @ListChambres, @ListCaracteristiques," +
                "@ListInterieur, @ListOptionSupp, @txtDescription)";

                using (OleDbCommand cmd = new OleDbCommand(insertQuery, _connection))
                {
                    cmd.Parameters.AddWithValue("@N", MAX_ID_Credit);
                    cmd.Parameters.AddWithValue("@txtCodeLocale", txtCodeLocale_);
                    cmd.Parameters.AddWithValue("@ListEtatLocale", ListEtatLocale_);
                    cmd.Parameters.AddWithValue("@ListArrondissement", ListArrondissement_);
                    cmd.Parameters.AddWithValue("@ListQuertier", ListQuertier_);
                    cmd.Parameters.AddWithValue("@txtAdresse", txtAdresse_);
                    cmd.Parameters.AddWithValue("@ListTypeSol", ListTypeSol_);
                    cmd.Parameters.AddWithValue("@ListEtage", ListEtage_);
                    cmd.Parameters.AddWithValue("@txtSurface", txtSurface_);
                    cmd.Parameters.AddWithValue("@txtPrix", txtPrix_);
                    cmd.Parameters.AddWithValue("@ListSalleBains", ListSalleBains_);
                    cmd.Parameters.AddWithValue("@ListChambres", ListChambres_);
                    cmd.Parameters.AddWithValue("@ListCaracteristiques", ListCaracteristiques_);
                    cmd.Parameters.AddWithValue("@ListInterieur", ListInterieur_);
                    cmd.Parameters.AddWithValue("@ListOptionSupp", ListOptionSupp_);
                    cmd.Parameters.AddWithValue("@txtDescription", txtDescription_);

                    if (_connection.State != ConnectionState.Open)
                    {
                        _connection.Open(); // Open the connection
                    }
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Locale Bien Ajouté ");

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

        private void SelectImages_Click_1(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtCodeLocale.Text))
            {
                return;
            }

            OpenFileDialog ofd = new OpenFileDialog
            {
                Title = "Selectionner des Photos",
                Multiselect = true,
                Filter = "JPG|*.jpg|JPEG|*.jpeg|GIF|*.gif|PNG|*.png"
            };

            DialogResult df = ofd.ShowDialog();

            if (df == DialogResult.OK)
            {
                string[] files = ofd.FileNames;

                string targetDirectory = Path.Combine(AffairesPath + @"\" + txtCodeLocale.Text);

                if (!Directory.Exists(targetDirectory))
                {
                    Directory.CreateDirectory(targetDirectory);
                }

                try
                {
                    foreach (string file in files)
                    {
                        string fileName = Path.GetFileName(file);
                        string destinationPath = Path.Combine(targetDirectory, fileName);

                        File.Copy(file, destinationPath, true);
                    }

                    MessageBox.Show("Images copiées avec succès dans Success: " + targetDirectory, "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Une erreur s'est produite lors de la copie des fichiers: " + ex.Message, "Ereur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            try
            {
                string MAJListeCodeLocale_ = MAJListeCodeLocale.SelectedValue?.ToString() ?? string.Empty;
                string MAJListeEtat_ = MAJListeEtat.SelectedValue?.ToString() ?? string.Empty;
                string Ville_ = MAJListeVille.SelectedValue?.ToString() ?? string.Empty;
                string MAJListArrondissement_ = MAJListArrondissement.SelectedValue?.ToString() ?? string.Empty;
                string MAJListQuertier_ = MAJListQuertier.SelectedValue?.ToString() ?? string.Empty;
                string MAJListeTypeSol_ = MAJListeTypeSol.SelectedValue?.ToString() ?? string.Empty;
                string MAJListeEtageBien_ = MAJListeEtageBien.SelectedValue?.ToString() ?? string.Empty;
                string MAJListeSalleBains_ = MAJListeSalleBains.SelectedValue?.ToString() ?? string.Empty;
                string MAJListeChambres_ = MAJListeChambres.SelectedValue?.ToString() ?? string.Empty;
                string MAJListeCaracteriqtiques_ = MAJListeCaracteriqtiques.SelectedValue?.ToString() ?? string.Empty;
                string MAJListeInterieur_ = MAJListeInterieur.SelectedValue?.ToString() ?? string.Empty;
                string MAJListeOptionSupp_ = MAJListeOptionSupp.SelectedValue?.ToString() ?? string.Empty;
                string MAJtxtAdresseComplet_ = string.IsNullOrWhiteSpace(MAJtxtAdresseComplet.Text) ? string.Empty : MAJtxtAdresseComplet.Text;
                string MAJtxtDescription_ = string.IsNullOrWhiteSpace(MAJtxtDescription.Text) ? string.Empty : MAJtxtDescription.Text;
                int MAJtxtSurface_ = string.IsNullOrWhiteSpace(MAJtxtSurface.Text) ? 0 : Convert.ToInt32(MAJtxtSurface.Text);
                int MAJtxtPrix_ = string.IsNullOrWhiteSpace(MAJtxtPrix.Text) ? 0 : Convert.ToInt32(MAJtxtPrix.Text);

                string insertQuery = @"
        UPDATE Appartement 
        SET 
            [Etat] = @MAJListEtatLocale,
            [Arrondissement] = @MAJListArrondissement, 
            [Quartier] = @MAJListQuertier,
            [AdresseComplete] = @MAJtxtAdresse, 
            [TypeSol] = @MAJListTypeSol, 
            [EtageBien] = @MAJListEtage, 
            [Surface] = @MAJtxtSurface,
            [Prix] = @MAJtxtPrix, 
            [SallesBains] = @MAJListSalleBains, 
            [Chambres] = @MAJListChambres, 
            [CaractqGene] = @MAJListCaracteristiques,
            [Interieur] = @MAJListInterieur, 
            [OptionSupp] = @MAJListOptionSupp, 
            [Description] = @MAJtxtDescription, 
            [Ville] = @MAJListeVille
        WHERE [Id_Appartement] = @MAJtxtCodeLocale";

                using (OleDbCommand cmd = new OleDbCommand(insertQuery, _connection))
                {
                    cmd.Parameters.AddWithValue("@MAJListEtatLocale", MAJListeEtat_);
                    cmd.Parameters.AddWithValue("@MAJListArrondissement", MAJListArrondissement_);
                    cmd.Parameters.AddWithValue("@MAJListQuertier", MAJListQuertier_);
                    cmd.Parameters.AddWithValue("@MAJtxtAdresse", MAJtxtAdresseComplet_);
                    cmd.Parameters.AddWithValue("@MAJListTypeSol", MAJListeTypeSol_);
                    cmd.Parameters.AddWithValue("@MAJListEtage", MAJListeEtageBien_);
                    cmd.Parameters.AddWithValue("@MAJtxtSurface", MAJtxtSurface_);
                    cmd.Parameters.AddWithValue("@MAJtxtPrix", MAJtxtPrix_);
                    cmd.Parameters.AddWithValue("@MAJListSalleBains", MAJListeSalleBains_);
                    cmd.Parameters.AddWithValue("@MAJListChambres", MAJListeChambres_);
                    cmd.Parameters.AddWithValue("@MAJListCaracteristiques", MAJListeCaracteriqtiques_);
                    cmd.Parameters.AddWithValue("@MAJListInterieur", MAJListeInterieur_);
                    cmd.Parameters.AddWithValue("@MAJListOptionSupp", MAJListeOptionSupp_);
                    cmd.Parameters.AddWithValue("@MAJtxtDescription", MAJtxtDescription_);
                    cmd.Parameters.AddWithValue("@MAJListeVille", Ville_);
                    cmd.Parameters.AddWithValue("@MAJtxtCodeLocale", MAJListeCodeLocale_);

                    if (_connection.State != ConnectionState.Open)
                    {
                        _connection.Open();
                    }

                    int affectedRows = cmd.ExecuteNonQuery();

                    if (affectedRows > 0)
                    {
                        MessageBox.Show("Locale Bien Modifié.");
                        LoadData(); // Refresh the DataGridView
                    }
                    else
                    {
                        MessageBox.Show("Aucune mise à jour effectuée. Veuillez vérifier les données.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur: " + ex.Message);
            }
            finally
            {
                _connection.Close();
            }

        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            try
            {
                if (txtRecherch.Text != "")
                {
                    DataView dv = _dataTable.DefaultView;
                    dv.RowFilter = string.Format("[Code Locale] ='{0}'", txtRecherch.Text);
                    DataGridViewListeLocaux.DataSource = dv;
                }
                else
                {
                    DataView dv = _dataTable.DefaultView;
                    dv.RowFilter = string.Format("", "");
                    DataGridViewListeLocaux.DataSource = dv;
                }
            }
            catch
            {

            }

        }

        private void Supprimer_Click(object sender, EventArgs e)
        {
            if (DataGridViewListeLocaux.SelectedRows.Count > 0)
            {
                try
                {
                    DataGridViewRow selectedRow = DataGridViewListeLocaux.SelectedRows[0];

                    string CodeLocale = selectedRow.Cells["Code Locale"].Value.ToString();

                    DataGridViewListeLocaux.Rows.Remove(selectedRow);

                    string deleteQuery = "DELETE FROM Appartement WHERE [Id_Appartement] = @CodeLocale";

                    using (OleDbCommand cmd = new OleDbCommand(deleteQuery, _connection))
                    {
                        cmd.Parameters.AddWithValue("@CodeLocale", CodeLocale);

                        if (_connection.State != ConnectionState.Open)
                        {
                            _connection.Open();
                        }

                        int rowsAffected = cmd.ExecuteNonQuery();
                        if (rowsAffected > 0)
                        {
                            MessageBox.Show("Locale a été supprimé avec succès.");
                        }
                        else
                        {
                            MessageBox.Show("Erreur lors de la suppression du Locale.");
                        }
                    }

                    // Refresh the DataGridView after deletion
                    LoadData();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erreur lors de la suppression du Locale.");
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
                MessageBox.Show("Veillez séléctionner une ligne");
            }

        }

        private void ListeCodeLocaleDetails_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedListeCodeLocaleDetails = ListeCodeLocaleDetails.SelectedValue?.ToString();

            if (string.IsNullOrEmpty(selectedListeCodeLocaleDetails))
            {
                return;
            }

            string path = Path.Combine(AffairesPath, selectedListeCodeLocaleDetails);

            if (Directory.Exists(path))
            {
                // Clear existing images in the FlowLayoutPanel
                flowLayoutPanel1.Controls.Clear();

                string[] imageFiles = Directory.GetFiles(path, "*.jpg");

                foreach (var imageFile in imageFiles)
                {
                    try
                    {
                        // Test the image file before loading
                        if (!File.Exists(imageFile))
                        {
                            MessageBox.Show($"File does not exist: {imageFile}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            continue;
                        }

                        using (var fs = new FileStream(imageFile, FileMode.Open, FileAccess.Read))
                        using (var ms = new MemoryStream())
                        {
                            fs.CopyTo(ms);
                            ms.Seek(0, SeekOrigin.Begin);

                            // Check if the stream contains a valid image
                            using (var img = Image.FromStream(ms))
                            {
                                var pictureBox = new PictureBox
                                {
                                    Size = new Size(100, 100),
                                    Image = (Image)img.Clone(), // Clone the image to avoid locking issues
                                    SizeMode = PictureBoxSizeMode.StretchImage,
                                    BorderStyle = BorderStyle.FixedSingle,
                                    Tag = imageFile
                                };

                                pictureBox.Click += PictureBox_Click; // Event handler for click event
                                flowLayoutPanel1.Controls.Add(pictureBox); // Add to FlowLayoutPanel
                            }
                        }
                    }
                    catch (OutOfMemoryException)
                    {
                        MessageBox.Show($"The file is not a valid image or is too large: {imageFile}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"An error occurred while loading {imageFile}: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            else
            {
                MessageBox.Show("Image directory not found!");
            }
        }

        private void ListeVille_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            LoadArrondissement();
        }

        private void ListArrondissement_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            LoadQuarttier();
        }

        private void MAJListeVille_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadArrondissement();
        }

        private void MAJListArrondissement_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadQuarttier();
        }

        private void txtCodeLocale_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(txtCodeLocale.Text))
                {
                    textShowError.Text = string.Empty;
                    txtCodeLocale.BackColor = SystemColors.Window; // Default background color
                    _dataTable.DefaultView.RowFilter = string.Empty;
                    DataGridViewListeLocaux.DataSource = _dataTable.DefaultView;
                    return; // Exit the method
                }

                DataView dv = _dataTable.DefaultView;
                dv.RowFilter = string.Format("CONVERT([Code Locale], 'System.String') LIKE '%{0}%'", txtCodeLocale.Text.Trim());

                if (dv.Count > 0)
                {
                    textShowError.Text = "Code Locale déjà existe !";
                    txtCodeLocale.BackColor = Color.Red;
                }
                else
                {
                    textShowError.Text = string.Empty;
                    txtCodeLocale.BackColor = SystemColors.Window;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur lors de la recherche : " + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
