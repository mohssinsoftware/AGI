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
            try
            {
                string sourcePath = Path.GetFullPath(AGI_DB.PathDataBase.Trim());
                string directoryPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Trimestere");
                string backupName = $"AGI_DB_{DateTime.Today:yyyyMMdd}.mdb";
                string destinationPath = Path.Combine(directoryPath, backupName);

                if (!File.Exists(sourcePath))
                {
                    throw new FileNotFoundException("La base de données à sauvegarder est introuvable.", sourcePath);
                }

                Directory.CreateDirectory(directoryPath);
                File.Copy(sourcePath, destinationPath, overwrite: true);
                InsertIntoHistorique(Path.GetFileNameWithoutExtension(backupName), DateTime.Today);
                MessageBox.Show("La sauvegarde trimestrielle a été créée : " + destinationPath,
                    "Sauvegarde terminée", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("La génération de la sauvegarde trimestrielle a échoué : " + ex.Message,
                    "Erreur de sauvegarde", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

                    if (_connection.State != ConnectionState.Open)
                    {
                        _connection.Open();
                    }
                    cmd.ExecuteNonQuery();
                    UpdateMoiPayLoc();
                    TruncateTables();
                }
            }
            finally
            {
                if (_connection.State == ConnectionState.Open)
                {
                    _connection.Close();
                }
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
            finally
            {
                if (_connection.State == ConnectionState.Open)
                {
                    _connection.Close();
                }
            }
        }

        private void TruncateTables()
        {
            string[] tableNames = { "Encaissement", "loc_encs" };

            foreach (string tableName in tableNames)
            {
                string deleteQuery = $"DELETE FROM [{tableName}]";
                using (OleDbCommand command = new OleDbCommand(deleteQuery, _connection))
                {
                    if (_connection.State != ConnectionState.Open)
                    {
                        _connection.Open();
                    }

                    command.ExecuteNonQuery();
                }
            }

        }
    }
}
