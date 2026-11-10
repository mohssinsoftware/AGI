using iTextSharp.text;
using iTextSharp.text.pdf;
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
    public partial class FormListDesVirementBanquaire : Form
    {   
        
        private OleDbConnection _connection;
        private OleDbDataAdapter _dataAdapter;
        private DataTable _dataTable;
        private string VirementPath = "";
        public FormListDesVirementBanquaire()
        {
            InitializeComponent();
            this.Icon = new Icon("AGI.ico");
            InitializeDatabase();
            LoadData();
            FnGetQuittancesPath();
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

                _dataAdapter = new OleDbDataAdapter("SELECT [code ger] AS [Code Gerance], [nom_cop] AS [Nom Proprietaire]," +
                    "[agence] AS [Agence], [banque] AS [Banque], [compte] AS [Compte]," +
                    "[pqrt] & '/' & [den] AS [Part] FROM part", _connection);
                OleDbCommandBuilder commandBuilder = new OleDbCommandBuilder(_dataAdapter);

                _dataTable = new DataTable();
                _dataAdapter.Fill(_dataTable);

                DataGridViewCheckBoxColumn checkBoxColumn = new DataGridViewCheckBoxColumn();
                checkBoxColumn.HeaderText = "";
                checkBoxColumn.Name = "Select";
                checkBoxColumn.Width = 20;
                DataGridViewListGerance.Columns.Add(checkBoxColumn);
                
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

        private void FnGetQuittancesPath()
        {
            try
            {
                // Adjust the SQL query to match the exact table and column names
                string selectQuery = "SELECT [Valeur] FROM [Params] WHERE [Id Param] = 'OrderVirement'";

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

                            VirementPath = row["Valeur"].ToString();
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

        private void FnLoadDebutFin()
        {
            try
            {
                // Clear any previous items to ensure a fresh start
                SelDebut.Items.Clear();
                SelFin.Items.Clear();

                // Add placeholder item
                SelDebut.Items.Add("");
                SelFin.Items.Add("");

                // Add actual data items
                for (int i = 1; i <= 12; i++)
                {

                    SelDebut.Items.Add(i.ToString());
                    SelFin.Items.Add(i.ToString());
                }

                // Select the placeholder by default
                SelDebut.SelectedIndex = 0;
                SelFin.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            
        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

        /****************************************************************************************************/

        private void FormListDesVirementBanquaire_Load(object sender, EventArgs e)
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
            
        }

        private void clearValue()
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

        private void txtRecherch_TextChanged_1(object sender, EventArgs e)
        {
            DataView dv = _dataTable.DefaultView;
            dv.RowFilter = string.Format("[Nom Proprietaire] LIKE '%{0}%'", txtRecherch.Text.Trim());
            DataGridViewListGerance.DataSource = dv;
        }

        private void btnRechercheCodeGerance_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtRecherchCodeGerance.Text != "")
                {
                    DataView dv = _dataTable.DefaultView;
                    dv.RowFilter = string.Format("[Code Gerance] = '{0}'", txtRecherchCodeGerance.Text);
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

        private DataTable GetDatLocation(string Codeger)
        {
            DataTable dataTable = new DataTable();

            try
            {
                string selectQuery_loc_encs = "SELECT [code loc], [nom loc], [datloc], [total], [impy] FROM [loc] WHERE [CODE GER] = ? ORDER BY [code loc] ASC";

                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open();
                }

                using (OleDbCommand command_loc_encs = new OleDbCommand(selectQuery_loc_encs, _connection))
                {
                    //dataTable.Rows.Add(1, "Alice", 25, "alice@example.com");

                    command_loc_encs.Parameters.Add(new OleDbParameter("?", Codeger));

                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(command_loc_encs))
                    {
                        adapter.Fill(dataTable);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading date location data: " + ex.Message);
            }
            finally
            {
                if (_connection.State == ConnectionState.Open)
                {
                    _connection.Close();
                }
            }
            return dataTable;
        }

        private DataTable GetEncaissementLocataire(string CodeGerance, int Debut, int Fin, DataTable tableDateLocation)
        {
            DataTable TableLocataire = new DataTable();

            try
            {

                TableLocataire.Columns.Add("ListCodeLocataires", typeof(string));
                TableLocataire.Columns.Add("ListNomLocataires", typeof(string));
                TableLocataire.Columns.Add("LitsLoyer", typeof(double));
                TableLocataire.Columns.Add("ListPerdioPay", typeof(string));
                TableLocataire.Columns.Add("ListPerdioImpy", typeof(string));
                TableLocataire.Columns.Add("ListMontantPay", typeof(double));
                TableLocataire.Columns.Add("ListMontantImp", typeof(double));

                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open();
                }

                for (int k = 0; k < tableDateLocation.Rows.Count; k++)
                {
                    DataRow RowLocataire = tableDateLocation.Rows[k];

                    string CodeLoc = RowLocataire["code loc"]?.ToString() ?? string.Empty;
                    string NomLoc = RowLocataire["nom loc"]?.ToString() ?? string.Empty;

                    string selectQuery_loc_encs = @"SELECT [Person_Paye], [code_loc], [total], [Nom_Locataire], [intitulé] 
                    FROM [Encaissement] WHERE [code_ger] = @CodeGerance AND [code_loc] = @CodeLoc
                    AND [intitulé] BETWEEN @Debut AND @Fin ORDER BY [intitulé] ASC";

                    using (OleDbCommand command_loc_encs = new OleDbCommand(selectQuery_loc_encs, _connection))
                    {
                        // Add parameters safely
                        command_loc_encs.Parameters.Add(new OleDbParameter("@CodeGerance", OleDbType.VarChar)).Value = CodeGerance;
                        command_loc_encs.Parameters.Add(new OleDbParameter("@CodeLoc", OleDbType.VarChar)).Value = CodeLoc;
                        command_loc_encs.Parameters.Add(new OleDbParameter("@Debut", OleDbType.Integer)).Value = Debut;
                        command_loc_encs.Parameters.Add(new OleDbParameter("@Fin", OleDbType.Integer)).Value = Fin;

                        using (OleDbDataAdapter adapter = new OleDbDataAdapter(command_loc_encs))
                        {
                            DataTable dataTable = new DataTable();
                            adapter.Fill(dataTable);

                            double total = 0;
                            string ListPerdioPay = "";
                            string ListPerdioImpy = "";
                            double ListMontantPay = 0;
                            double ListMontantImp = 0;

                            total = Convert.ToDouble(RowLocataire["total"].ToString());
                            int Impy = 0;
                            string LastMonthImp = "";

                            if (Fin == DateTime.Today.Month)
                            {
                                Impy = Convert.ToInt32(RowLocataire["impy"].ToString());
                                LastMonthImp = DateTime.Today.Month.ToString();
                            }
                            else
                            {
                                if (Convert.ToInt32(RowLocataire["impy"].ToString()) != 0)
                                {
                                    Impy = (Convert.ToInt32(RowLocataire["impy"].ToString()) - 1);
                                    LastMonthImp = GetCalculMonth(DateTime.Today.Month, 1).Month.ToString();
                                }
                            }

                            if (Impy != 0)
                            {
                                ListPerdioImpy = (GetCalculMonth(DateTime.Today.Month, Impy)).Month + "/" + getLast2Yers(DateTime.Today.Year) + " au " +
                                (GetCalculMonth(DateTime.Today.Month, 1)).Month + "/" + getLast2Yers(DateTime.Today.Year);
                                ListMontantImp = (Impy * total);

                                if (Fin == DateTime.Today.Month)
                                {
                                    ListPerdioImpy = (GetCalculMonth(DateTime.Today.Month, (Impy - 1))).Month + "/" + getLast2Yers(DateTime.Today.Year) + " au " +
                                    DateTime.Today.Month + "/" + getLast2Yers(DateTime.Today.Year);
                                }
                            }
                            else
                            {
                                ListPerdioImpy = "//////////";
                                ListMontantImp = 0;
                            }

                            if (dataTable.Rows.Count > 0)
                            {
                                for (int i = 0; i < dataTable.Rows.Count; i++)
                                {
                                    DataRow row = dataTable.Rows[i];

                                    int FinPeriode = Convert.ToInt32(row["intitulé"].ToString());


                                    DateTime DateLocation = Convert.ToDateTime(RowLocataire["datloc"].ToString());

                                    if (Debut == Fin)
                                    {
                                        ListPerdioPay = Debut + "/" + getLast2Yers(DateTime.Today.Year);
                                        ListMontantPay = total;
                                    }
                                    else if (Fin > Debut)
                                    {
                                        ListPerdioPay = Debut + "/" + getLast2Yers(DateTime.Today.Year) + " au " +
                                        FinPeriode + "/" + getLast2Yers(DateTime.Today.Year);
                                        ListMontantPay = (ListMontantPay + total);
                                    }
                                }

                            }
                            else
                            {
                                ListPerdioPay = "//////////";
                            }
                            TableLocataire.Rows.Add(CodeLoc, NomLoc, total, ListPerdioPay, ListPerdioImpy, ListMontantPay, ListMontantImp);
                        }
                    }
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

            return TableLocataire;
        }

        private DateTime GetCalculMonth(int Month, int Num)
        {
            DateTime adjustedDate = DateTime.Now;

            try
            {
                DateTime currentDate = new DateTime(DateTime.Now.Year, Month, 1);
                adjustedDate = currentDate.AddMonths(-Num);
            }
            catch (Exception ex)
            {

            }
            return adjustedDate;
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

        private List<string> GetTVACommission(string Codeger)
        {
            List<string> TVAcommision = new List<string>();

            try
            {
                string selectQuery_loc_encs = "SELECT [comission], [TVA] FROM [ger] WHERE [CODE GER] = ?";

                if (_connection.State != ConnectionState.Open)
                {
                    _connection.Open();
                }

                using (OleDbCommand command_loc_encs = new OleDbCommand(selectQuery_loc_encs, _connection))
                {
                    DataTable dataTable = new DataTable();
                    command_loc_encs.Parameters.Add(new OleDbParameter("?", Codeger));

                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(command_loc_encs))
                    {
                        adapter.Fill(dataTable);

                        for (int i = 0; i < dataTable.Rows.Count; i++)
                        {
                            TVAcommision.Add(dataTable.Rows[0]["TVA"].ToString());
                            TVAcommision.Add(dataTable.Rows[0]["comission"].ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading date location data: " + ex.Message);
            }
            finally
            {
                if (_connection.State == ConnectionState.Open)
                {
                    _connection.Close();
                }
            }
            return TVAcommision;
        }

        private double getMontantVirement(string Codeger, string NomCopor, int Deb, int Fin, string Part)
        {
            double result = 0;
            double PartDouble = 0;
            try
            {
                List<string> LstTVAcomm = GetTVACommission(Codeger);
                double TVA = Convert.ToDouble(LstTVAcomm[0].ToString());
                double comm = Convert.ToDouble(LstTVAcomm[1].ToString());

                double depend = (TVA + comm);

                try
                {
                    string cleanedInput = Part.Replace("%", "").Trim();
                    string[] parts = cleanedInput.Split('/');
                    if (parts.Length == 2)
                    {
                        double numerator = double.Parse(parts[0]);
                        double denominator = double.Parse(parts[1]);
                        PartDouble = numerator / denominator;
                    }
                    else
                    {
                        Console.WriteLine("Merci de vérifier la part de " + NomCopor + "dans votre base de données !");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Merci de vérifier la part de " + NomCopor + "dans votre base de données !");
                }

                DataTable DT_Locataires = GetDatLocation(Codeger);

                DataTable DTLocataire = GetEncaissementLocataire(Codeger, Deb, Fin, DT_Locataires);

                double TotalPay = CalculateTotalSum(DTLocataire, "ListMontantPay");

                TotalPay = (TotalPay - depend);

                result = (TotalPay * PartDouble);
            }
            catch
            {

            }

            return result;
        }

        private string getMaxId()
        {
            string MAX_ID_Virement = "1";

            try
            {
                string Query_MAX_ID_Credit = "SELECT MAX([N°]) AS N° FROM recap_ord_virm";

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
                            MAX_ID_Virement = Res_Int.ToString();
                        }
                        else
                        {
                            // MessageBox.Show("Impossible de récupérer le MAX N° Crédit !!");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("La table des Virement dans la base de données est vide, le numéro de virement sera être 1 " + ex.Message);
            }
            finally
            {
                // Ensure the connection is closed even if an error occurs
                if (_connection.State == ConnectionState.Open)
                {
                    _connection.Close();
                }
            }
            return MAX_ID_Virement;
        }

        private void fnAjouter(string CodeGer, string NomGer, string Montant, string Banque)
        {
            string MAX_ID_Virement = "0";
            MAX_ID_Virement =  getMaxId();

            try
            {
                string insertQuery = "INSERT INTO recap_ord_virm ([N°], [code_ger], [nom_ger],  [date], [banque], [montant]) " +
                                     "VALUES (@N, @codeGer, @NomGer, @date, @baqnue, @montant)";

                using (OleDbCommand cmd = new OleDbCommand(insertQuery, _connection))
                {
                    cmd.Parameters.AddWithValue("@N", MAX_ID_Virement);
                    cmd.Parameters.AddWithValue("@codeGer", CodeGer);
                    cmd.Parameters.AddWithValue("@NomGer", NomGer);
                    cmd.Parameters.AddWithValue("@date", DateTime.Today.ToString("yyyy/MM/dd"));
                    cmd.Parameters.AddWithValue("@baqnue", Banque);
                    cmd.Parameters.AddWithValue("@montant", Montant);
                    //Add more parameters as needed

                    if (_connection.State != ConnectionState.Open)
                    {
                        _connection.Open(); // Open the connection
                    }
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Veuillez verifier Votre system les insértions des virement ne fonctionne pas ! : " + ex.Message);
            }
            finally
            {
                _connection.Close();
            }
        }

        private void btnGenerateOrderVirement_Click(object sender, EventArgs e)
        {
            try
            {

                string SelDebut_ = SelDebut.SelectedItem?.ToString();
                string SelFin_ = SelFin.SelectedItem?.ToString();

                if (SelDebut.SelectedItem != "" && SelFin.SelectedItem != "")
                {
                    int iSelDebut_ = Convert.ToInt32(SelDebut_);
                    int iSelFin_ = Convert.ToInt32(SelFin_);

                    if (iSelDebut_ > iSelFin_)
                    {
                        MessageBox.Show("début de période doit être toujours supérieur à fin de période !");
                        return;
                    }

                    var selectedRows = new List<DataRow>();

                    foreach (DataGridViewRow row in DataGridViewListGerance.Rows)
                    {
                        if (Convert.ToBoolean(row.Cells["Select"].Value) == true)
                        {
                            selectedRows.Add(((DataRowView)row.DataBoundItem).Row);
                        }
                    }

                    foreach (var row in selectedRows)
                    {
                        string Codeger = row["Code Gerance"].ToString();
                        string NomCopor = row["Nom Proprietaire"].ToString();
                        string Banque = row["Banque"].ToString();
                        string Compte = row["Compte"].ToString();
                        string Part = row["Part"].ToString();

                        double MontantVir = getMontantVirement(Codeger, NomCopor, iSelDebut_, iSelFin_, Part);

                        int integerPart = (int)MontantVir;
                        int decimalPart = (int)((MontantVir - integerPart) * 100);

                        string MontantString = NumberToWords(integerPart);
                        string MontantCentim = NumberToWords(decimalPart);

                        string PathComplet = GeneratedFile.GetOutputPath(
                            VirementPath, $"OrderVirement_{Codeger}_{NomCopor}.pdf");

                        string IdVirement = getMaxId();

                       //GeneratePdfTable(PathComplet, Codeger, NomCopor, Banque, Compte, MontantVir, iSelDebut_, iSelFin_, MontantString, MontantCentim, IdVirement);

                        if (MontantVir <= 0)
                        {
                            MessageBox.Show("Le copropriétaire " + NomCopor + " n'a aucun montant à verser !");
                        }
                        else
                        {
                            GeneratePdfTable(PathComplet, Codeger, NomCopor, Banque, Compte, MontantVir, iSelDebut_, iSelFin_, MontantString, MontantCentim, IdVirement);
                            fnAjouter(Codeger, NomCopor, MontantVir.ToString(), Banque);
                            ShowFile(PathComplet);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Merci de sélectionner au moins un Coproprietaire ! Erreur: ");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Merci de sélectionner au moins un Coproprietaire ! Erreur: " + ex.Message);
            }
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

        static string NumberToWords(long number)
        {
            if (number == 0)
                return "zéro";

            if (number < 0)
                return "moins " + NumberToWords(Math.Abs(number));

            string words = "";

            if ((number / 1000000000) > 0)
            {
                words += NumberToWords(number / 1000000000) + " milliard ";
                number %= 1000000000;
            }

            if ((number / 1000000) > 0)
            {
                words += NumberToWords(number / 1000000) + " million ";
                number %= 1000000;
            }

            if ((number / 1000) > 0)
            {
                words += NumberToWords(number / 1000) + " mille ";
                number %= 1000;
            }

            if ((number / 100) > 0)
            {
                words += NumberToWords(number / 100) + " cent ";
                number %= 100;
            }

            if (number > 0)
            {
                if (words != "")
                    words += "et ";

                string[] unitsMap = { "zéro", "un", "deux", "trois", "quatre", "cinq", "six", "sept", "huit", "neuf", "dix",
                                  "onze", "douze", "treize", "quatorze", "quinze", "seize", "dix-sept", "dix-huit", "dix-neuf" };
                string[] tensMap = { "zéro", "dix", "vingt", "trente", "quarante", "cinquante", "soixante", "soixante-dix", "quatre-vingt", "quatre-vingt-dix" };

                if (number < 20)
                    words += unitsMap[number];
                else
                {
                    words += tensMap[number / 10];
                    if ((number % 10) > 0)
                        words += "-" + unitsMap[number % 10];
                }
            }

            return words.Trim();
        }

        public void GeneratePdfTable(string filePath, string CodeGer, string NomCop, string Banque, string Compte, double MontantVir, int Deb, int Fin, string MontantString, string MontantCentim, string IdVirement)
        {
            Document document = new Document(PageSize.A4);

            if (MontantCentim != "")
            {
                MontantCentim = "et "+ MontantCentim + " cts";
            }

            try 
            {   
                string Mois = "mois "+ Deb + "/" + getLast2Yers(DateTime.Today.Year) + "";

                if (Deb < Fin)
                {
                    for(int i = Deb+1; i <= Fin; i++)
                    {
                        Mois = Mois + " et "+ i + "/" + getLast2Yers(DateTime.Today.Year) + "";
                    }
                }

                PdfWriter.GetInstance(document, new FileStream(filePath, FileMode.Create));
                document.Open();

                // Define custom font
                var titleFont = FontFactory.GetFont(FontFactory.COURIER_BOLD, 12, BaseColor.BLACK);
                var normalFont = FontFactory.GetFont(FontFactory.COURIER, 10, BaseColor.BLACK);

                // Add Title
                document.Add(new Paragraph("Banque Populaire — Agence Sahraoui", titleFont));
                document.Add(new Paragraph("ORDRE DE VIREMENT", titleFont));
                document.Add(new Paragraph("n°"+ IdVirement + "   B.P DH :"+ MontantVir.ToString("00.00") + " dhs\n", normalFont));

                // Adding content with some alignment and spacing
                document.Add(new Paragraph("\n", normalFont));
                document.Add(new Paragraph("DONNEUR D'ORDRE\n", titleFont));
                document.Add(new Paragraph("Numero de compte à débiter : 190 780 21211 9306585 0029 73\nIntitulé du compte : agence de gestion immobilière — A.G.I SARL\nMontant du virement : "+ MontantString + " dhs "+ MontantCentim + ".\nMotif : loyer du "+Mois+" - GERANCE N°"+ CodeGer + ".\n", normalFont));

                // Beneficiary details
                document.Add(new Paragraph("BENEFICIAIRE\n", titleFont));
                document.Add(new Paragraph("Banque domiciliataire : "+ Banque + ".\nNuméro de compte à créditer : "+ Compte + "\nIntitulé du compte : "+ NomCop + ".\n", normalFont));

                // Footer with date
                document.Add(new Paragraph("\n\n\n", titleFont));
                document.Add(new Paragraph($"Casablanca le "+ DateTime.Today.ToString("yyyy/MM/dd") +"\n", normalFont));

                Console.WriteLine("fichier d'Order de virement généré avec succès dans le Path : " + filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            finally
            {
                document.Close();
            }
        }

    }
}
