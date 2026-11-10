
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
    public partial class FormMiseSoldeCommission : Form
    {   

        private OleDbConnection _connection;
        private OleDbDataAdapter _dataAdapter;
        private DataTable _dataTable;

        public FormMiseSoldeCommission()
        {
            InitializeComponent();
            this.Icon = new Icon("AGI.ico");
            InitializeDatabase();
            fnGetOldCommission();
        }
        private void InitializeDatabase()
        {
            string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+ AGI_DB.PathDataBase + "";
            _connection = new OleDbConnection(connectionString);
        }

        private void restaurerCommission_Click(object sender, EventArgs e)
        {
            try
            {
                string selectedNewCommission = txtNewCommission.Text?.ToString();

                if (string.IsNullOrEmpty(selectedNewCommission))
                {
                    MessageBox.Show("veuillez entrer une nouvelle valeur Commission.");
                    return;
                }

                string updateQuery = "UPDATE frais_ger SET [montant]=@Montant WHERE [depense] = 'Commission'";

                using (OleDbCommand cmd = new OleDbCommand(updateQuery, _connection))
                {
                    cmd.Parameters.AddWithValue("@Montant", Convert.ToDouble(txtNewCommission.Text));

                    if (_connection.State != ConnectionState.Open)
                    {
                        _connection.Open();
                    }
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Commission Bien Modifié");
                    textOldComs.Text = Convert.ToDouble(txtNewCommission.Text).ToString("N2");
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
        private void fnGetOldCommission()
        {
            try
            {
                string selectQuery = "SELECT [depense], [montant] FROM [frais_ger]";

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

                        for( int i = 0; i < dataTable.Rows.Count; i++)
                        {
                            DataRow row = dataTable.Rows[i];
                            if (row["depense"].ToString() == "Commission")
                            {
                                textOldComs.Text = Convert.ToDouble(row["montant"]).ToString("N2");
                            }
                        }
                    }
                }
                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open();
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
    }
}
