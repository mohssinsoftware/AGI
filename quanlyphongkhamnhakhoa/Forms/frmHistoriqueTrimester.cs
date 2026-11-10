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
    public partial class FormHistoriqueTrimester : Form
    {
        
        private OleDbConnection _connection;
        private OleDbDataAdapter _dataAdapter;
        private DataTable _dataTable;
        private DataTable _dataTableValideEncaissement;
        private string RelevePath = "";
        public FormHistoriqueTrimester()
        {
            InitializeComponent();
            this.Icon = new Icon("AGI.ico");
            InitializeDatabase();
            LoadData_ValiderEncaissement();
        }

        /****************************************************************************************************/
        private void InitializeDatabase()
        {
            string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+ AGI_DB.PathDataBase + "";
            _connection = new OleDbConnection(connectionString);
        }
        private void LoadData_ValiderEncaissement()
        {   
            try
            {
                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open();
                }

                _dataAdapter = new OleDbDataAdapter("SELECT [Id_Trimester] AS [Code Trimestere], [Intitule] AS [Intitulé Trimestere]," +
                    " [date] AS [Date Trimestere] FROM Historique_Trimester ORDER BY [Id_Trimester] ASC", _connection);
                OleDbCommandBuilder commandBuilder = new OleDbCommandBuilder(_dataAdapter);

                _dataTableValideEncaissement = new DataTable();
                _dataAdapter.Fill(_dataTableValideEncaissement);

                dataGridViewHisTrimestre.DataSource = _dataTableValideEncaissement;
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




        /*-- UPDATE DATETIME VALUE ----------------------------------------------------------------*/
        /*-- UPDATE DATETIME VALUE ----------------------------------------------------------------*/


        private void MAJCodeGerance_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        /****************************************************************************************************/

        private void FormHistoriqueTrimester_Load(object sender, EventArgs e)
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
        

        private void btnRechercheParCodeGerance_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtRechercheCodeTrimestere.Text != "")
                {
                    DataView dv = _dataTableValideEncaissement.DefaultView;
                    dv.RowFilter = string.Format("[Code Trimestere] ='{0}'", txtRechercheCodeTrimestere.Text);
                    dataGridViewHisTrimestre.DataSource = dv;
                }
                else
                {
                    DataView dv = _dataTableValideEncaissement.DefaultView;
                    dv.RowFilter = string.Format("", "");
                    dataGridViewHisTrimestre.DataSource = dv;
                }
            }
            catch (Exception ex)
            {

            }

        }

        private void textDateLoyer_ValueChanged(object sender, EventArgs e)
        {
            DataView dv = _dataTableValideEncaissement.DefaultView;
            dv.RowFilter = string.Format("[Date Trimestere] ='{0}'", textDateTrimestre.Text.Trim());
            dataGridViewHisTrimestre.DataSource = dv;
        }
    }
}
