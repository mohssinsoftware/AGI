
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
    public partial class FormParamsAvancees : Form
    {
        private OleDbConnection _connection;
        private OleDbDataAdapter _dataAdapter;
        private DataTable _dataTable;

        public FormParamsAvancees()
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
                // Define the SQL query with a WHERE clause to filter by the selected "code loc"
                string selectQuery = "SELECT [Id Param], [Enable Param], [Valeur] FROM [Params]";

                // Check if the connection is not already open
                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open(); // Open the connection
                }

                // Use OleDbCommand to fetch the data with parameterized query
                using (OleDbCommand command = new OleDbCommand(selectQuery, _connection))
                {
                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                    {
                        // Fill a DataTable with the results
                        DataTable dataTable = new DataTable();

                        adapter.Fill(dataTable); // Fill the DataTable

                        for (int i = 0; i < dataTable.Rows.Count; i++)
                        {
                            DataRow row = dataTable.Rows[i];

                            string ExpressionSwitch = row["Id Param"].ToString();

                            switch (ExpressionSwitch)
                            {
                                case "Encaissement":
                                    txtEncaissement_EMP.Text = row["Valeur"].ToString();
                                    break;
                                case "OrderVirement":
                                    txtOrderVirement_EMP.Text = row["Valeur"].ToString();
                                    break;
                                case "Affaires":
                                    txtAffairePhoto_EMP.Text = row["Valeur"].ToString();
                                    break;
                                case "Quittance":
                                    txtQuittance.Text = row["Valeur"].ToString();
                                    break;
                                case "Releve":
                                    txtReleve.Text = row["Valeur"].ToString();
                                    break;
                                case "ListeRetard":
                                    ListeRetard.Text = row["Valeur"].ToString();
                                    break;
                                case "RetardLocataire":
                                     RetardLocataire.Text = row["Valeur"].ToString();
                                    break;
                                case "EtatDetaille":
                                    EtatDetaille.Text = row["Valeur"].ToString();
                                    break;
                                default: break;
                                    //EtatDetaille
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur lors du chargement des données des emplacement des fichiers : " + ex.Message + " Merci de vérifier votre base de données !");
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

        private void btnLoad_Click(object sender, EventArgs e)
        {
            LoadData();
        }

        /****************************************************************************************************/

        private void FormParamsAvancees_Load(object sender, EventArgs e)
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


        private void btnModifierLemplacement_Click(object sender, EventArgs e)
        {
            try
            {

                string updateQuery_Encaissement    = "UPDATE Params SET [Valeur] = @Encaissement  WHERE [Id Param] = 'Encaissement'";
                string updateQuery_OrderVirement   = "UPDATE Params SET [Valeur] = @OrderVirement WHERE [Id Param] = 'OrderVirement'";
                string updateQuery_Affaires        = "UPDATE Params SET [Valeur] = @Affaires WHERE [Id Param] = 'Affaires'";
                string updateQuery_Quittance       = "UPDATE Params SET [Valeur] = @Affaires WHERE [Id Param] = 'Quittance'";
                string updateQuery_Releve          = "UPDATE Params SET [Valeur] = @Releve WHERE [Id Param] = 'Releve'";
                string updateQuery_ListeRetard     = "UPDATE Params SET [Valeur] = @ListeRetard WHERE [Id Param] = 'ListeRetard'";
                string updateQuery_RetardLocataire = "UPDATE Params SET [Valeur] = @RetardLocataire WHERE [Id Param] = 'RetardLocataire'";
                string updateQuery_EtatDetaille    = "UPDATE Params SET [Valeur] = @EtatDetaille WHERE [Id Param] = 'EtatDetaille'";

                //EtatDetaille

                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open();
                }

                using (OleDbCommand cmd_Encaissement = new OleDbCommand(updateQuery_Encaissement, _connection))
                {
                    cmd_Encaissement.Parameters.AddWithValue("@Encaissement", txtEncaissement_EMP.Text) ;
                    cmd_Encaissement.ExecuteNonQuery();
                }
                using (OleDbCommand cmd_OrderVirement = new OleDbCommand(updateQuery_OrderVirement, _connection))
                {
                    cmd_OrderVirement.Parameters.AddWithValue("@OrderVirement", txtOrderVirement_EMP.Text);
                    cmd_OrderVirement.ExecuteNonQuery();
                }
                using (OleDbCommand cmd_Affaires = new OleDbCommand(updateQuery_Affaires, _connection))
                {
                    cmd_Affaires.Parameters.AddWithValue("@Affaires", txtAffairePhoto_EMP.Text);
                    cmd_Affaires.ExecuteNonQuery();
                }
                using (OleDbCommand cmd_Quittance = new OleDbCommand(updateQuery_Quittance, _connection))
                {
                    cmd_Quittance.Parameters.AddWithValue("@Quittance", txtQuittance.Text);
                    cmd_Quittance.ExecuteNonQuery();
                }
                using (OleDbCommand cmd_Releve = new OleDbCommand(updateQuery_Releve, _connection))
                {
                    cmd_Releve.Parameters.AddWithValue("@Releve", txtReleve.Text);
                    cmd_Releve.ExecuteNonQuery();
                }
                using (OleDbCommand cmd_ListeRetard = new OleDbCommand(updateQuery_ListeRetard, _connection))
                {
                    cmd_ListeRetard.Parameters.AddWithValue("@ListeRetard", ListeRetard.Text);
                    cmd_ListeRetard.ExecuteNonQuery();
                }
                using (OleDbCommand cmd_RetardLocataire = new OleDbCommand(updateQuery_RetardLocataire, _connection))
                {
                    cmd_RetardLocataire.Parameters.AddWithValue("@RetardLocataire", RetardLocataire.Text);
                    cmd_RetardLocataire.ExecuteNonQuery();
                }

                using (OleDbCommand cmd_EtatDetaille = new OleDbCommand(updateQuery_EtatDetaille, _connection))
                {
                    cmd_EtatDetaille.Parameters.AddWithValue("@EtatDetaille", EtatDetaille.Text);
                    cmd_EtatDetaille.ExecuteNonQuery();

                    MessageBox.Show("l'emplacement est bien modifié !");
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

        private void RunQuery_Click(object sender, EventArgs e)
        {   
            string alterTableQuery = txtQuery.Text;

            try
            {
                using (OleDbConnection connection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=AGI_DB.mdb"))
                {   
                    connection.Open();

                    using (OleDbCommand command = new OleDbCommand(alterTableQuery, connection))
                    {
                        command.ExecuteNonQuery();
                        MessageBox.Show($"Successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error : {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
