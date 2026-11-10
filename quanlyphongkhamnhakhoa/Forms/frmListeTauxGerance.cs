

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
    public partial class FormListeTauxGerance : Form
    {   
        private OleDbConnection _connection;
        private OleDbDataAdapter _dataAdapter;
        private DataTable _dataTable;

        public FormListeTauxGerance()
        {
            InitializeComponent();
            this.Icon = new Icon("AGI.ico");
            InitializeDatabase();
            LoadData();
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
                _dataAdapter = new OleDbDataAdapter("SELECT [code ger] AS [Code Gerance], " +
                    "[nom_cop] AS [Nom Proprietaire]," +
                    "[pqrt] & '/' & [den] AS [Part], " +
                    "[adresse1] AS [Adresse] FROM part", _connection);
                OleDbCommandBuilder commandBuilder = new OleDbCommandBuilder(_dataAdapter);

                _dataTable = new DataTable();
                _dataAdapter.Fill(_dataTable);

                // Bind the updated DataTable to the DataGridView
                DataGridViewListGerance.DataSource = _dataTable;

                //DEBUT : Ajustement automatique pour les colonnes restantes-------------------------------------------------
                DataGridViewListGerance.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                // Ajustement spécifique pour la colonne "Nom Locataire"
                DataGridViewListGerance.Columns["Adresse"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; // Désactiver l'ajustement automatique
                DataGridViewListGerance.Columns["Adresse"].Width = 270; // Définir la largeur souhaitée

                // Ajustement spécifique pour la colonne "Nom Locataire"
                DataGridViewListGerance.Columns["Nom Proprietaire"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; // Désactiver l'ajustement automatique
                DataGridViewListGerance.Columns["Nom Proprietaire"].Width = 400; // Définir la largeur souhaitée
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
    }
}
