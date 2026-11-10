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
    public partial class FormEtatDetaille : Form
    {   
        
        private OleDbConnection _connection;
        private OleDbDataAdapter _dataAdapter;
        private DataTable _dataTable;

        public FormEtatDetaille()
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

        private void btnTaiLai_Click(object sender, EventArgs e)
        {
            
        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

        /****************************************************************************************************/

        private void FormEtatDetaille_Load(object sender, EventArgs e)
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

        private void btnGenereEtatDetaile_Click(object sender, EventArgs e)
        {
            try
            {
                string deleteQuery = "Update Quartiers set Ville='Casablanca'";

                using (OleDbCommand cmd = new OleDbCommand(deleteQuery, _connection))
                {
                    if (_connection.State != ConnectionState.Open)
                    {
                        _connection.Open();
                    }
                }
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

                for (int j = 0; j < ListCodeGer.Count; j++)
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

        public void GeneratePdfTable(string filePath)
        {
            // Define the fonts
            iTextSharp.text.Font headerFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8);
            iTextSharp.text.Font regularFont = FontFactory.GetFont(FontFactory.HELVETICA, 8);
            iTextSharp.text.Font boldFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8);

            // Create a new document
            Document document = new Document(PageSize.A4);

            try
            {
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
                Paragraph dateAndRecipient = new Paragraph("Casablanca, le 05/01/2024\n\n" +
                                                           "Madame ZIANE Lamia\n" +
                                                           "Rue BACHIR EL ALJ\n" +
                                                           "Résidence NADIA Imb B\n" +
                                                           "Appt n°8 - CASABLANCA\n" +
                                                           "N° Gérance : 18\n" +
                                                           "10 RUE POUPONNIERE\n\n",
                                                           regularFont);
                dateAndRecipient.Alignment = Element.ALIGN_LEFT;
                dateAndRecipient.IndentationLeft = PageSize.A4.Width / 2;
                document.Add(dateAndRecipient);

                // Opening statement
                Paragraph opening = new Paragraph("Madame,\n" +
                                                  "Nous vous adressons ci-dessous relevé des loyers du Quatrième Trimestre 2023\n\n",
                                                  regularFont);
                document.Add(opening);

                // Table: Code Loc, Loyer, etc.
                PdfPTable table = new PdfPTable(5);
                table.WidthPercentage = 100;
                
                table.SetWidths(new float[] { 2, 2, 2, 2, 2 });
                
                

                // Add table headers
                table.AddCell(new PdfPCell(new Phrase("Code Loc", boldFont)));
                table.AddCell(new PdfPCell(new Phrase("Loyer", boldFont)));
                table.AddCell(new PdfPCell(new Phrase("Période payé", boldFont)));
                table.AddCell(new PdfPCell(new Phrase("Montant payé", boldFont)));
                table.AddCell(new PdfPCell(new Phrase("Période impy", boldFont)));
                table.AddCell(new PdfPCell(new Phrase("Montant impy", boldFont)));

                // Add table data
                table.AddCell("LAHRICHI Hadj Mohamed");
                table.AddCell("475.00");
                table.AddCell("////////////////////////");
                table.AddCell("0.00");
                table.AddCell("10/23 au 12/23");
                table.AddCell("1,425.00");

                table.AddCell("GARAGE ABBES");
                table.AddCell("150.00");
                table.AddCell("09/23 au 10/23");
                table.AddCell("300.00");
                table.AddCell("11/23 au 12/23");
                table.AddCell("300.00");

                table.AddCell("IMADID Jillali");
                table.AddCell("920.00");
                table.AddCell("08/23 au 12/23");
                table.AddCell("4,600.00");
                table.AddCell("à jour");
                table.AddCell("0.00");

                // Total row
                table.AddCell(new PdfPCell(new Phrase("Total", boldFont)) { Colspan = 4, HorizontalAlignment = Element.ALIGN_RIGHT });
                table.AddCell("4,900.00");
                table.AddCell("1,725.00");
                document.Add(table);

                // Expenses section
                Paragraph expenses = new Paragraph("\nDépense          Montant\n" +
                                                   "Commission       490.00\n" +
                                                   "TVA              98.00\n" +
                                                   "Total :          588.00\n\n" +
                                                   "Solde :          4,312.00\n",
                                                   regularFont);
                document.Add(expenses);

                // Coproprietor section
                Paragraph coproprietor = new Paragraph("Copropriétaire       Taux      Montant\n" +
                                                       "Madame ZIANE Lamia   35/64     2,358.13\n" +
                                                       "Madame JAWAANI Fatima 22/64    1,482.25\n" +
                                                       "Madame ZIANE Latifa  7/64      471.63\n\n",
                                                       regularFont);
                document.Add(coproprietor);

                // Closing statement
                Paragraph closing = new Paragraph("Veuillez agréer, Madame, l'expression de nos meilleures salutations.\n\n" +
                                                  "Le gérant\n\n" +
                                                  "<< Sauf erreur ou omission >>",
                                                  regularFont);
                document.Add(closing);
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
    }
}
