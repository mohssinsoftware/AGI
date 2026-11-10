using System;
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
    public partial class FormArrTrimestere : Form
    {
        private OleDbConnection _connection;
        private OleDbDataAdapter _dataAdapter;
        private DataTable _dataTable;

        public FormArrTrimestere()
        {
            InitializeComponent();
            this.Icon = new Icon("AGI.ico");
            InitializeDatabase();
        }

        /****************************************************************************************************/
        private void InitializeDatabase()
        {
            string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+ AGI_DB.PathDataBase + "";
            _connection = new OleDbConnection(connectionString);
        }


        /****************************************************************************************************/

        private void FormArrTrimestere_Load(object sender, EventArgs e)
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

        private void LancerNouvelleTrimester()
        {
            try
            {

            }
            catch(Exception ex)
            {
                
            }
        }

        private void RunQuery_Click(object sender, EventArgs e)
        {   
            string sourcePath = Path.GetFullPath(@"AGI_DB.mdb"); // Convert to absolute path
            string directoryPath = Path.GetFullPath(@"Trimestere"); // Convert to absolute path
            string destinationPath = Path.Combine(directoryPath, $"AGI_DB_{DateTime.Today:yyyyMMdd}.mdb");

            try 
            {   
                Console.WriteLine($"Starting the duplication process...");

                Console.WriteLine($"Source Path: {sourcePath}");
                Console.WriteLine($"Destination Path: {destinationPath}");

                if (!Directory.Exists(directoryPath))
                {
                    Console.WriteLine("Destination directory does not exist. Creating it...");
                    Directory.CreateDirectory(directoryPath);
                }

                if (File.Exists(sourcePath))
                {
                    Console.WriteLine("Source database found. Proceeding to copy...");
                    File.Copy(sourcePath, destinationPath, overwrite: true);
                    Console.WriteLine("Database duplicated successfully!");
                    InsertIntoHistorique($"AGI_DB_{DateTime.Today:yyyyMMdd}", DateTime.Today);
                }
                else
                {
                    Console.WriteLine("Source database file does not exist. Please check the source path.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while duplicating the database:");
                Console.WriteLine(ex.Message);
            }
            finally
            {
                Console.WriteLine("Process completed. Press Enter to exit.");
                Console.ReadLine();
            }
        }

        private void InsertIntoHistorique(string Intitule, DateTime date)
        {
            try
            {
                string insertQuery = "INSERT INTO Historique_Trimester ([Intitule], [date]) " +
                                     "VALUES (@Intitule, @date)";

                using (OleDbCommand cmd = new OleDbCommand(insertQuery, _connection))
                {
                    cmd.Parameters.AddWithValue("@Intitule", Intitule);
                    cmd.Parameters.AddWithValue("@date", date);
                    //Add more parameters as needed

                    if (_connection.State != ConnectionState.Open)
                    {
                        _connection.Open(); // Open the connection
                    }
                    cmd.ExecuteNonQuery();
                    UpdateMoiPayLoc();
                    TruncateTables();
                    MessageBox.Show("La nouvelle Trimestre Bien Ajouté dans l'historique de Trimester");
                }
            }
            catch (Exception ex)
            {
                
            }
            finally
            {
                _connection.Close();
            }
        }

        private void UpdateMoiPayLoc()
        {
            try
            {

                string updateQuery = "UPDATE loc SET [py] = 0 ";

                using (OleDbCommand cmd = new OleDbCommand(updateQuery, _connection))
                {
                    if (_connection.State != ConnectionState.Open)
                    {
                        _connection.Open();
                    }

                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
               
            }
            finally
            {
                _connection.Close();
            }
        }

        private void TruncateTables()
        {
            string[] tableNames = { "Encaissement", "loc_encs" };

            try
            {
                using (OleDbConnection connection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=AGI_DB.mdb"))
                {
                    connection.Open();
                    foreach (var tableName in tableNames)
                    {
                        string alterTableQuery = $"DELETE FROM [{tableName}]";

                        using (OleDbCommand command = new OleDbCommand(alterTableQuery, connection))
                        {
                            command.ExecuteNonQuery();
                        }
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
