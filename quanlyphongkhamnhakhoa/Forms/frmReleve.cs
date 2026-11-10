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
    public partial class FormReleve : Form
    {
        private OleDbConnection _connection;
        private OleDbDataAdapter _dataAdapter;
        private DataTable _dataTable;
        private DataTable _dataTableValideEncaissement;
        private string RelevePath = "";
        public FormReleve()
        {
            InitializeComponent();
            this.Icon = new Icon("AGI.ico");
            InitializeDatabase();
            LoadData_ValiderEncaissement();
            FnGetQuittancesPath();
            FnLoadDebutFin();
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

                _dataAdapter = new OleDbDataAdapter("SELECT [nom_cop] AS [Nom Coproprietaire], [code ger] AS [Code Gerance], " +
                "[adresse1] AS [Adresse 1],[adresse2] AS [Adresse 2],[adresse3] AS [Adresse 3]," +
                " [autre_frais] AS [Autre frais] FROM part ORDER BY [code ger] ASC", _connection);
                OleDbCommandBuilder commandBuilder = new OleDbCommandBuilder(_dataAdapter);

                _dataTableValideEncaissement = new DataTable();
                _dataAdapter.Fill(_dataTableValideEncaissement);

                // Add checkbox column
                DataGridViewCheckBoxColumn checkBoxColumn = new DataGridViewCheckBoxColumn();
                checkBoxColumn.HeaderText = "";
                checkBoxColumn.Name = "Select";
                checkBoxColumn.Width = 10;
                dataGridViewValiderEncaissement.Columns.Add(checkBoxColumn);

                dataGridViewValiderEncaissement.DataSource = _dataTableValideEncaissement;
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
                string selectQuery = "SELECT [Valeur] FROM [Params] WHERE [Id Param] = 'Releve'";

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

                            RelevePath = row["Valeur"].ToString();
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

        private void MAJCodeGerance_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        /****************************************************************************************************/

        private void FormReleve_Load(object sender, EventArgs e)
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
        
        public void GeneratePdfEncaissement(DataTable dataTable_Periode, DataTable DTLocataire, DataTable DTCoporp)
        {   
            Document document = new Document(PageSize.A4, 30, 30, 10, 10);

            try
            {
                string filePath = dataTable_Periode.Rows[0]["filePath"].ToString();
                string imagePath = dataTable_Periode.Rows[0]["imagePath"].ToString();
                string codeGer = dataTable_Periode.Rows[0]["codeGer"].ToString();
                string NomCoprop = dataTable_Periode.Rows[0]["NomCoprop"].ToString();
                string Adresse1 = dataTable_Periode.Rows[0]["Adresse1"].ToString();
                string Adresse2 = dataTable_Periode.Rows[0]["Adresse2"].ToString();
                string Adresse3 = dataTable_Periode.Rows[0]["Adresse3"].ToString();
                double commission = Convert.ToDouble(dataTable_Periode.Rows[0]["commission"].ToString());
                double TVA = Convert.ToDouble(dataTable_Periode.Rows[0]["TVA"].ToString());
                double Total = Convert.ToDouble(dataTable_Periode.Rows[0]["Total"].ToString());
                double Sold = Convert.ToDouble(dataTable_Periode.Rows[0]["Sold"].ToString());

                // Define the fonts
                iTextSharp.text.Font headerFont = FontFactory.GetFont(FontFactory.COURIER_BOLD, 8);
                iTextSharp.text.Font regularFont = FontFactory.GetFont(FontFactory.COURIER, 8);
                iTextSharp.text.Font boldFont = FontFactory.GetFont(FontFactory.COURIER, 6);

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

                    for (int i = 0; i < DTLocataire.Rows.Count; i++)
                    {
                        table.AddCell(new PdfPCell(new Phrase(DTLocataire.Rows[i]["ListCodeLocataires"].ToString(), boldFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                        table.AddCell(new PdfPCell(new Phrase(DTLocataire.Rows[i]["ListNomLocataires"].ToString(), boldFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                        table.AddCell(new PdfPCell(new Phrase(Convert.ToDouble(DTLocataire.Rows[i]["LitsLoyer"].ToString()).ToString("00.00"), boldFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                        table.AddCell(new PdfPCell(new Phrase(DTLocataire.Rows[i]["ListPerdioPay"].ToString(), boldFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                        table.AddCell(new PdfPCell(new Phrase(Convert.ToDouble(DTLocataire.Rows[i]["ListMontantPay"].ToString()).ToString("00.00"), boldFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                        table.AddCell(new PdfPCell(new Phrase(DTLocataire.Rows[i]["ListPerdioImpy"].ToString(), boldFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                        table.AddCell(new PdfPCell(new Phrase(Convert.ToDouble(DTLocataire.Rows[i]["ListMontantImp"].ToString()).ToString("00.00"), boldFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                    }

                    double TotalPay = CalculateTotalSum(DTLocataire, "ListMontantPay");
                    double TotalImp = CalculateTotalSum(DTLocataire, "ListMontantImp");

                    table.AddCell(new PdfPCell(new Phrase("", boldFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                    table.AddCell(new PdfPCell(new Phrase("", boldFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                    table.AddCell(new PdfPCell(new Phrase("", boldFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                    table.AddCell(new PdfPCell(new Phrase("Total : ", FontFactory.GetFont(FontFactory.COURIER_BOLD, 14))) { HorizontalAlignment = Element.ALIGN_LEFT });
                    table.AddCell(new PdfPCell(new Phrase(TotalPay.ToString("00.00"), FontFactory.GetFont(FontFactory.COURIER_BOLD, 14))) { HorizontalAlignment = Element.ALIGN_LEFT });
                    table.AddCell(new PdfPCell(new Phrase("")) { HorizontalAlignment = Element.ALIGN_RIGHT });
                    table.AddCell(new PdfPCell(new Phrase(TotalImp.ToString("00.00"), FontFactory.GetFont(FontFactory.COURIER_BOLD, 14))) { HorizontalAlignment = Element.ALIGN_RIGHT });

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

                double TotalDeponse = (commission + TVA);

                // Add data rows for the "Dépense" table
                expenseTable.AddCell(new PdfPCell(new Phrase("Commission", regularFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                expenseTable.AddCell(new PdfPCell(new Phrase(new Phrase(commission.ToString("00.00"), regularFont))));

                expenseTable.AddCell(new PdfPCell(new Phrase("TVA", regularFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                expenseTable.AddCell(new PdfPCell(new Phrase(new Phrase(TVA.ToString("00.00"), regularFont))));

                // Add Total row
                PdfPCell totalLabelCellExpense = new PdfPCell(new Phrase("Total :", boldFont)) { HorizontalAlignment = Element.ALIGN_RIGHT };
                totalLabelCellExpense.Colspan = 1; // Span over one cell
                expenseTable.AddCell(totalLabelCellExpense);
                expenseTable.AddCell(new PdfPCell(new Phrase(new Phrase((TotalDeponse).ToString("00.00"), regularFont))));

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
                PdfPCell soldeLabelCell = new PdfPCell(new Phrase("Solde :", FontFactory.GetFont(FontFactory.COURIER_BOLD, 14)))
                {
                    Border = iTextSharp.text.Rectangle.NO_BORDER, // No border to match the styling
                    HorizontalAlignment = Element.ALIGN_RIGHT
                };

                double TotalSolde = (TotalPay - TotalDeponse);

                PdfPCell soldeAmountCell = new PdfPCell(new Phrase(TotalSolde.ToString("00.00"), FontFactory.GetFont(FontFactory.COURIER_BOLD, 14)))
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
                    for(int j = 0; j < DTCoporp.Rows.Count; j++)
                    {
                        double pqrt = Convert.ToDouble(DTCoporp.Rows[j]["pqrt"].ToString());
                        double den = Convert.ToDouble(DTCoporp.Rows[j]["den"].ToString());
                        double SoldCop = TotalSolde * (pqrt / den);
                        coproprietaireTable.AddCell(new PdfPCell(new Phrase(DTCoporp.Rows[j]["nom_cop"].ToString(), regularFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                        coproprietaireTable.AddCell(new PdfPCell(new Phrase(pqrt + "/"+ den, regularFont)) { HorizontalAlignment = Element.ALIGN_LEFT });
                        coproprietaireTable.AddCell(new PdfPCell(new Phrase(SoldCop.ToString("00.00"), regularFont)) { HorizontalAlignment = Element.ALIGN_RIGHT });
                    }

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
                Console.WriteLine("Erreur de génération de Relevé : " + ex.Message);
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
                mainContentTable.AddCell(new PdfPCell(new Phrase(Usage.ToString(), boldFont)) { HorizontalAlignment = Element.ALIGN_LEFT });

                // Add main content table to the document
                document.Add(mainContentTable);

                Paragraph rightSectionContent = new Paragraph("Additional Information", boldFont)
                {
                    Alignment = Element.ALIGN_LEFT, 
                    SpacingBefore = 0f,
                    IndentationLeft = separationPosition + 10 
                };

                document.Add(rightSectionContent);
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


        private void btnValiderencaissement_Click(object sender, EventArgs e)
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
                else if ((iSelDebut_ == 1 || iSelDebut_ == 2 || iSelDebut_ == 3) &&
                         (iSelFin_ != 1 && iSelFin_ != 2 && iSelFin_ != 3))
                {
                    MessageBox.Show("vous ne pouvez pas sélectionner une période qui dépasse la premier trimestre !");
                    return;
                }
                else if ((iSelDebut_ == 4 || iSelDebut_ == 5 || iSelDebut_ == 6) &&
                         (iSelFin_ != 4 && iSelFin_ != 5 && iSelFin_ != 6))
                {
                    MessageBox.Show("vous ne pouvez pas sélectionner une période qui dépasse la deuxième trimestre !");
                    return;
                }
                else if ((iSelDebut_ == 7 || iSelDebut_ == 8 || iSelDebut_ == 9) &&
                (iSelFin_ != 7 && iSelFin_ != 8 && iSelFin_ != 9))
                {
                    MessageBox.Show("vous ne pouvez pas sélectionner une période qui dépasse la troisième trimestre !");
                    return;
                }
                else if ((iSelDebut_ == 10 || iSelDebut_ == 11 || iSelDebut_ == 12) &&
                (iSelFin_ != 10 && iSelFin_ != 11 && iSelFin_ != 12))
                {
                    MessageBox.Show("vous ne pouvez pas sélectionner une période qui dépasse la quatrième trimestre !");
                    return;
                }

                try
                {
                    var selectedRows = new List<DataRow>();

                    foreach (DataGridViewRow row in dataGridViewValiderEncaissement.Rows)
                    {
                        if (Convert.ToBoolean(row.Cells["Select"].Value) == true)
                        {
                            selectedRows.Add(((DataRowView)row.DataBoundItem).Row);
                        }
                    }

                    if (selectedRows.Count == 0)
                    {
                        MessageBox.Show("Veuillez sélectionner des copropriétaires !");
                        return;
                    }

                    foreach (var row in selectedRows)
                    {
                        string Codeger = row["Code Gerance"].ToString();
                        string NomCopor = row["Nom Coproprietaire"].ToString();
                        string Adresse1 = row["Adresse 1"].ToString();
                        string Adresse2 = row["Adresse 2"].ToString();
                        string Adresse3 = row["Adresse 3"].ToString();
                        List<string> LstTVAcomm = new List<string>();
                        LstTVAcomm = GetTVACommission(Codeger);
                        double TVA = Convert.ToDouble(LstTVAcomm[0].ToString());
                        double comm = Convert.ToDouble(LstTVAcomm[1].ToString());
                        fnValiderEncaissementParCoprop(Codeger, NomCopor, Adresse1, Adresse2, Adresse3, iSelDebut_, iSelFin_, TVA, comm);
                    }
                    MessageBox.Show("Les relevés sont générés avec succès, vous pouvez les trouver dans le dossier :" + RelevePath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Merci de sélectionner au moins un Coproprietaire ! Erreur: " + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("merci de selectionner debut et fin période !");
                return;
            }
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
                                (GetCalculMonth(DateTime.Today.Month,1)).Month + "/" + getLast2Yers(DateTime.Today.Year);
                                ListMontantImp = (Impy * total);

                                if (Fin == DateTime.Today.Month)
                                {
                                    ListPerdioImpy = (GetCalculMonth(DateTime.Today.Month, (Impy-1))).Month + "/" + getLast2Yers(DateTime.Today.Year) + " au " +
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
                                 ListPerdioPay  = "//////////";
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

        private string getLast2Yers(int year)
        {
            string lastTwoDigitsString = "";
            try
            {
                string yearString = year.ToString();
                lastTwoDigitsString = yearString.Substring(yearString.Length - 2);
            }
            catch(Exception ex)
            {

            }
            return lastTwoDigitsString;
        }
        private DateTime GetCalculMonth(int Month, int Num)
        {
            DateTime adjustedDate = DateTime.Now;

            try
            {
                DateTime currentDate = new DateTime(DateTime.Now.Year, Month, 1);
                adjustedDate = currentDate.AddMonths(-Num);
            }
            catch(Exception ex)
            {

            }
            return adjustedDate;
        }

        private void fnValiderEncaissementParCoprop(string CodeGerance, string NomCoprop, string Adresse1, string Adresse2, 
        string Adresse3, int Deb, int Fin, double TVA, double commission)
        {   
            double Total = 0;
            double Sold = 0;

            try
            {
                DataTable tableDateLocation = GetDatLocation(CodeGerance);

                DataTable DTLocataire =  GetEncaissementLocataire(CodeGerance, Deb, Fin, tableDateLocation);

                DataTable DTCop = GetCoprop(CodeGerance);

                DataTable dataTable_Details = new DataTable();
                dataTable_Details.Columns.Add("filePath", typeof(string));
                dataTable_Details.Columns.Add("imagePath", typeof(string));
                dataTable_Details.Columns.Add("codeGer", typeof(string));
                dataTable_Details.Columns.Add("NomCoprop", typeof(string));
                dataTable_Details.Columns.Add("Adresse1", typeof(string));
                dataTable_Details.Columns.Add("Adresse2", typeof(string));
                dataTable_Details.Columns.Add("Adresse3", typeof(string));
                dataTable_Details.Columns.Add("commission", typeof(double));
                dataTable_Details.Columns.Add("TVA", typeof(double));
                dataTable_Details.Columns.Add("Total", typeof(double));
                dataTable_Details.Columns.Add("Sold", typeof(double));

                string PathComplet = RelevePath + $@"\Encaissement_{CodeGerance}_{NomCoprop}.pdf";

                dataTable_Details.Rows.Add(PathComplet,
                @"Logo_AGI_page-0001-removebg-preview.png",Adresse1, Adresse2, Adresse3, CodeGerance, NomCoprop,
                commission,TVA,Total,Sold);

                GeneratePdfEncaissement(dataTable_Details, DTLocataire, DTCop);

                ShowFile(PathComplet);
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
        private DataTable GetCoprop(string Codeger)
        {
            DataTable dataTable = new DataTable();

            try
            {
                string selectQuery_loc_encs = "SELECT [nom_cop], [pqrt], [den] FROM [part] WHERE [code ger] = ?";

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

                        for(int i = 0; i < dataTable.Rows.Count; i++)
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

        private void btnRechercheParCodeGerance_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtRechercheGerance_Copro.Text != "")
                {
                    DataView dv = _dataTableValideEncaissement.DefaultView;
                    dv.RowFilter = string.Format("[Code Gerance] ='{0}'", txtRechercheGerance_Copro.Text);
                    dataGridViewValiderEncaissement.DataSource = dv;
                }
                else
                {
                    DataView dv = _dataTableValideEncaissement.DefaultView;
                    dv.RowFilter = string.Format("", "");
                    dataGridViewValiderEncaissement.DataSource = dv;
                }
            }
            catch (Exception ex)
            {

            }

        }

        private void textNomCprop_TextChanged(object sender, EventArgs e)
        {
            DataView dv = _dataTableValideEncaissement.DefaultView;
            dv.RowFilter = string.Format("[Nom Coproprietaire] LIKE '%{0}%'", textNomCprop.Text.Trim());
            dataGridViewValiderEncaissement.DataSource = dv;
        }
    }
}
