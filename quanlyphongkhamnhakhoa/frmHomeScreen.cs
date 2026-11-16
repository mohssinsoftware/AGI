using PKNK_CNPM;
using PKNK_CNPM.Forms;
using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace PKNK_CNPM
{
    public partial class frmHomeScreen : Form
    {
        bool sidebarExpand = true;
        bool customerCollapse = true;
        bool quanlyCollapse = true;
        bool settingCollapse = true;
        bool ReleveCollapse = true;
        bool FraisCollapse = true;
        bool CeditCollapse = true;
        bool GeranceCollapse = true;
        bool AppartementCollapse = true;
        //fiedls
        private Random random;
        private Button currentButton;
        private Form activeForm;
        private int tempIndex;
        /*public object SiderbarTimer { get; private set; }*/
        //contruction
        
        public frmHomeScreen()
        {
            this.WindowState = FormWindowState.Maximized;
            InitializeComponent();
            sidebar.AutoScroll = true;
            AddNewScreensToSettingsMenu();
            string displaySettingsError;
            if (!DisplayPreferences.Load(out displaySettingsError))
            {
                MessageBox.Show("Impossible de lire les préférences d'affichage : " + displaySettingsError,
                    "Préférences d'affichage", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            this.Icon = new Icon("AGI.ico");
            random = new Random();
            btnCloseChildForm.Visible = false;
            this.Text = string.Empty;
            this.ControlBox = false;
            this.MaximizedBounds = Screen.FromHandle(this.Handle).WorkingArea;
        }
        private void AddNewScreensToSettingsMenu()
        {
            const int menuItemHeight = 44;
            int firstMenuItemTop = panel50.Bottom + btnNouveauTrimester.Bottom - panel50.Height;
            int requiredHeight = firstMenuItemTop + 10 * menuItemHeight;
            int minimumHeight = SettingContainer.MinimumSize.Height;
            int maximumHeight = minimumHeight + ((requiredHeight - minimumHeight + 9) / 10) * 10;
            SettingContainer.MaximumSize = new Size(405, maximumHeight);

            AddSettingsScreenButton("Tableau de bord", sender => OpenChildForm(() => new FormPilotage(), sender), 0);
            AddSettingsScreenButton("Suivi des impayés", sender => OpenChildForm(() => new FormSuiviImpayes(), sender), 1);
            AddSettingsScreenButton("Fiche unique par logement", sender => OpenChildForm(() => new FormFicheLogement(), sender), 2);
            AddSettingsScreenButton("Recherche et filtres", sender => OpenChildForm(() => new FormRechercheHomogene(), sender), 3);
            AddSettingsScreenButton("Encaissements plus sécurisés", sender => OpenChildForm(() => new FormEncaissementsSecurises(), sender), 4);
            AddSettingsScreenButton("Documents plus fiables", sender => OpenChildForm(() => new FormDocumentsFiables(), sender), 5);
            AddSettingsScreenButton("Historique des actions", sender => OpenChildForm(() => new FormHistoriqueActions(), sender), 6);
            AddSettingsScreenButton("Sauvegarde rassurante", sender => OpenChildForm(() => new FormSauvegardeRassurante(), sender), 7);
            AddSettingsScreenButton("Affichage plus pratique", sender => OpenChildForm(() => new FormAffichagePratique(), sender), 8);
            AddSettingsScreenButton("Alertes utiles", sender => OpenChildForm(() => new FormAlertesUtiles(), sender), 9);
        }

        private void AddSettingsScreenButton(string text, Action<object> openScreen, int index)
        {
            Button screenButton = new Button
            {
                Name = "btnEcranAjoute" + index.ToString(CultureInfo.InvariantCulture),
                Text = "     " + text,
                Location = new Point(0, panel50.Bottom + btnNouveauTrimester.Bottom - panel50.Height + index * 44),
                Size = new Size(405, 44),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(28, 0, 0, 0),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.Teal,
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 10F, FontStyle.Regular),
                Cursor = Cursors.Hand,
                UseVisualStyleBackColor = false
            };
            screenButton.FlatAppearance.BorderSize = 0;
            screenButton.Click += (sender, e) => openScreen(sender);
            SettingContainer.Controls.Add(screenButton);
            screenButton.BringToFront();
        }

        //Drag form
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);
        //methods
        private Color SelectThemeColor()
        {
            int index = random.Next(Themecolor.ColorList.Count);
            while (tempIndex == index)
            {
                index = random.Next(Themecolor.ColorList.Count);
            }
            tempIndex = index;
            string color = Themecolor.ColorList[index];
            return ColorTranslator.FromHtml(color);
        }
        private void ActivateButton(object btnSender)
        {   
            if (btnSender != null)
            {
                
                if (currentButton != (Button)btnSender)
                {
                    DisableButton();

                    Color color = Color.Teal;
                    //Color color = SelectThemeColor();
                    //color = Color.CadetBlue;
                    //color = Color.DarkCyan;

                    currentButton = (Button)btnSender;
                    currentButton.BackColor = color;
                    currentButton.ForeColor = Color.White;
                    //currentButton.Font = new System.Drawing.Font("Segoe UI", 13.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
                    panelTitleBar.BackColor = color;
                    panelMenu.BackColor = Themecolor.ChangeColorBrightness(color, -0.3);
                    Themecolor.PrimaryColor = color;
                    Themecolor.SecondaryColor = Themecolor.ChangeColorBrightness(color, -0.3);
                    btnCloseChildForm.Visible = true;
                    btnCloseChildForm.BackColor = color;
                }
                
            }
        }
 
        private void DisableButton()
        {
            foreach (Control previousBtn in sidebar.Controls)
            {
                    previousBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(55)))), ((int)(((byte)(51)))), ((int)(((byte)(76)))));
                    previousBtn.ForeColor = Color.Gainsboro;
                    previousBtn.Font = new System.Drawing.Font("Segoe UI", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
                    
            }
        }
        //Show Forms
        private void OpenChildForm<TForm>(Func<TForm> createChildForm, object btnSender) where TForm : Form
        {
            if (activeForm is TForm && !activeForm.IsDisposed)
            {
                activeForm.BringToFront();
                return;
            }

            TForm childForm = createChildForm();
            DisplayPreferences.ApplyTo(childForm);
            panelDesktoppanel.SuspendLayout();
            try
            {
                if (activeForm != null)
                {
                    activeForm.Close();
                }

                // Reset();
                //ActivateButton(btnSender);
                activeForm = childForm;
                childForm.TopLevel = false;
                childForm.FormBorderStyle = FormBorderStyle.None;
                childForm.MinimumSize = Size.Empty;
                childForm.Dock = DockStyle.Fill;
                panelDesktoppanel.Controls.Add(childForm);
                panelDesktoppanel.Tag = childForm;
                childForm.BringToFront();
                childForm.Show();
                ConfigureResponsiveScrollableGroups(childForm);
                lblTitle.Text = childForm.Text;
            }
            finally
            {
                panelDesktoppanel.ResumeLayout(true);
            }
        }

        private static void ConfigureResponsiveScrollableGroups(Control parent)
        {
            FlowLayoutPanel flowPanel = parent as FlowLayoutPanel;
            if (flowPanel != null && flowPanel.AutoScroll && flowPanel.Controls.Count == 1 &&
                flowPanel.Controls[0] is GroupBox && flowPanel.Parent != null && flowPanel.Parent.Controls.Count == 1)
            {
                GroupBox content = (GroupBox)flowPanel.Controls[0];
                flowPanel.Dock = DockStyle.Fill;
                EventHandler resizeContent = (sender, e) =>
                {
                    int width = Math.Max(1, flowPanel.ClientSize.Width - flowPanel.Padding.Horizontal - content.Margin.Horizontal);
                    int height = Math.Max(1, flowPanel.ClientSize.Height - flowPanel.Padding.Vertical - content.Margin.Vertical);
                    content.Size = new Size(width, height);
                };
                flowPanel.SizeChanged += resizeContent;
                resizeContent(flowPanel, EventArgs.Empty);
            }

            foreach (Control child in parent.Controls)
            {
                ConfigureResponsiveScrollableGroups(child);
            }
        }

        private void btnCloseChildForm_Click_1(object sender, EventArgs e)
        {
            if (activeForm != null)
                activeForm.Close();
            Reset();
        }
        private void Reset()
        {
            DisableButton();
            lblTitle.Text = "HOME";
            panelTitleBar.BackColor = Color.FromArgb(0, 150, 136);
            panelMenu.BackColor = Color.FromArgb(39, 39, 58);
            btnCustomer.BackColor = Color.FromArgb(55, 51, 76);
            btnQuanLy.BackColor = Color.FromArgb(55, 51, 76);
            btnSetting.BackColor = Color.FromArgb(55, 51, 76);
            btnReleve.BackColor = Color.FromArgb(55, 51, 76);
            btnCredit.BackColor = Color.FromArgb(55, 51, 76);
            btnCreateLocataire.BackColor = Color.FromArgb(55, 51, 76);
            btnListeEtatsJournaliers.BackColor = Color.FromArgb(55, 51, 76);
            btnListeEncaissements.BackColor = Color.FromArgb(55, 51, 76);
            btnSaisirEncaissement.BackColor = Color.FromArgb(55, 51, 76);
            btnModifierEncaissement.BackColor = Color.FromArgb(55, 51, 76);
            if (currentButton != null)
                currentButton.Font = new System.Drawing.Font("Segoe UI", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(163))); 
            btnCloseChildForm.Visible = false;
            currentButton = null;


        }
        //Drag Form
        private void panelTitleBar_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }


        //Button Max, close, minimize
        private void btnClosePage_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        private void btnMaximine_Click(object sender, EventArgs e)
        {
            WindowState = WindowState == FormWindowState.Maximized
                ? FormWindowState.Normal
                : FormWindowState.Maximized;
        }
        private void btnMinimize_Click_1(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void sidebarTimer_Tick(object sender, EventArgs e)
        {
            if (sidebarExpand)
            {
                sidebar.Width = Math.Min(sidebar.MaximumSize.Width, sidebar.Width + 10);
                if (sidebar.Width >= sidebar.MaximumSize.Width)
                {
                    sidebar.Width = sidebar.MaximumSize.Width;
                    sidebarTimer.Stop();
                }
            }
            else
            {
                sidebar.Width = Math.Max(sidebar.MinimumSize.Width, sidebar.Width - 10);
                if (sidebar.Width <= sidebar.MinimumSize.Width)
                {
                    sidebar.Width = sidebar.MinimumSize.Width;
                    sidebarTimer.Stop();
                }
            }
        }

        private void menuButton_Click_1(object sender, EventArgs e)
        {
            sidebarExpand = !sidebarExpand;
            sidebarTimer.Start();
        }

        //Set Timer các chức năng trong toolBar
        private void btnCustomer_Click_1(object sender, EventArgs e)
        {
             CustomerTimer.Start();
        }

        private void CustomerTimer_Tick_1(object sender, EventArgs e)
        {
            sidebar.SuspendLayout();
            try
            {
                if (customerCollapse)
                {
                    CustomerContainer.Height += 10;
                    if (CustomerContainer.Height == CustomerContainer.MaximumSize.Height)
                    {
                        customerCollapse = false;
                        CustomerTimer.Stop();
                    }
                    QuanlyContainer.Height -= 20;
                    CollapseSettingContainer();
                    CreditContainer.Height -= 20;
                    ReleveContainer.Height -= 20;
                    FraisContainer.Height -= 20;
                    GeranceContainer.Height -= 20;
                    AppartementContainer.Height -= 20;

                }
                else
                {
                    CustomerContainer.Height -= 10;
                    if (CustomerContainer.Height == CustomerContainer.MinimumSize.Height)
                    {
                        customerCollapse = true;
                        CustomerTimer.Stop();
                    }
                }
            }
            finally
            {
                sidebar.ResumeLayout(true);
            }
        }
        private void QuanlyTimer_Tick_1(object sender, EventArgs e)
        {
            sidebar.SuspendLayout();
            try
            {
                if (quanlyCollapse)
                {
                    QuanlyContainer.Height += 10;
                    if (QuanlyContainer.Height == QuanlyContainer.MaximumSize.Height)
                    {
                        quanlyCollapse = false;
                        QuanlyTimer.Stop();
                    }
                    CustomerContainer.Height -= 20;
                    CollapseSettingContainer();
                    CreditContainer.Height -= 20;
                    ReleveContainer.Height -= 20;
                    FraisContainer.Height -= 20;
                    GeranceContainer.Height -= 20;
                    AppartementContainer.Height -= 20;
                }
                else
                { 
                    QuanlyContainer.Height -= 10;
                    if (QuanlyContainer.Height == QuanlyContainer.MinimumSize.Height)
                    {
                        quanlyCollapse = true;
                        QuanlyTimer.Stop();

                    }
                }

            }
            finally
            {
                sidebar.ResumeLayout(true);
            }
        }
        private void btnQuanLy_Click_1(object sender, EventArgs e)
        {   
            QuanlyTimer.Start();
        }
        
        private void CollapseSettingContainer()
        {
            SettingContainer.Height = Math.Max(SettingContainer.MinimumSize.Height, SettingContainer.Height - 20);
            if (SettingContainer.Height == SettingContainer.MinimumSize.Height)
            {
                settingCollapse = true;
                SettingTimer.Stop();
            }
        }

        private void SettingTimer_Tick_1(object sender, EventArgs e)
        {
            sidebar.SuspendLayout();
            try
            {
                if (settingCollapse)
                {
                    SettingContainer.Height += 10;
                    if (SettingContainer.Height == SettingContainer.MaximumSize.Height)
                    {
                        settingCollapse = false;
                        SettingTimer.Stop();
                    }
                    CustomerContainer.Height -= 20;
                    QuanlyContainer.Height -= 20;
                    CreditContainer.Height -= 20;
                    ReleveContainer.Height -= 20;
                    FraisContainer.Height -= 20;
                    GeranceContainer.Height -= 20;
                    AppartementContainer.Height -= 20;
                }
                else
                {
                    SettingContainer.Height -= 10;
                    if (SettingContainer.Height == SettingContainer.MinimumSize.Height)
                    {
                        settingCollapse = true;
                        SettingTimer.Stop();

                    }
                }

            }
            finally
            {
                sidebar.ResumeLayout(true);
            }
        }
        //
        private void ReleveTimer_Tick_1(object sender, EventArgs e)
        {
            sidebar.SuspendLayout();
            try
            {
                if (ReleveCollapse)
                {
                    ReleveContainer.Height += 10;
                    if (ReleveContainer.Height == ReleveContainer.MaximumSize.Height)
                    {
                        ReleveCollapse = false;
                        ReleveTimer.Stop();
                    }
                    CustomerContainer.Height -= 20;
                    QuanlyContainer.Height -= 20;
                    CollapseSettingContainer();
                    CreditContainer.Height -= 20;
                    FraisContainer.Height -= 20;
                    GeranceContainer.Height -= 20;
                    AppartementContainer.Height -= 20;
                }
                else
                {
                    ReleveContainer.Height -= 10;
                    if (ReleveContainer.Height == ReleveContainer.MinimumSize.Height)
                    {
                        ReleveCollapse = true;
                        ReleveTimer.Stop();

                    }
                }

            }
            finally
            {
                sidebar.ResumeLayout(true);
            }
        }
        private void FraisTimer_Tick_1(object sender, EventArgs e)
        {
            sidebar.SuspendLayout();
            try
            {
                if (FraisCollapse)
                {
                    FraisContainer.Height += 10;
                    if (FraisContainer.Height == FraisContainer.MaximumSize.Height)
                    {
                        FraisCollapse = false;
                        FraisTimer.Stop();
                    }
                    CustomerContainer.Height -= 20;
                    QuanlyContainer.Height -= 20;
                    CollapseSettingContainer();
                    CreditContainer.Height -= 20;
                    ReleveContainer.Height -= 20;
                    GeranceContainer.Height -= 20;
                    AppartementContainer.Height -= 20;
                }
                else
                {
                    FraisContainer.Height -= 10;
                    if (FraisContainer.Height == FraisContainer.MinimumSize.Height)
                    {
                        FraisCollapse = true;
                        FraisTimer.Stop();

                    }
                }

            }
            finally
            {
                sidebar.ResumeLayout(true);
            }
        }

        private void AppartementTimer_Tick_1(object sender, EventArgs e)
        {   
            sidebar.SuspendLayout();
            try
            {
                if (AppartementCollapse)
                {
                    AppartementContainer.Height += 10;
                    if (AppartementContainer.Height == AppartementContainer.MaximumSize.Height)
                    {
                        AppartementCollapse = false;
                        AppartementTimer.Stop();
                    }
                    CustomerContainer.Height -= 20;
                    QuanlyContainer.Height -= 20;
                    CollapseSettingContainer();
                    CreditContainer.Height -= 20;
                    ReleveContainer.Height -= 20;
                    FraisContainer.Height -= 20;
                    GeranceContainer.Height -= 20;
                }
                else
                {
                    AppartementContainer.Height -= 10;
                    if (AppartementContainer.Height == AppartementContainer.MinimumSize.Height)
                    {
                        AppartementCollapse = true;
                        AppartementTimer.Stop();
                    }
                }
            }
            finally
            {
                sidebar.ResumeLayout(true);
            }
        }


        private void GeranceTimer_Tick_1(object sender, EventArgs e)
        {
            sidebar.SuspendLayout();
            try
            {
                if (GeranceCollapse)
                {
                    GeranceContainer.Height += 10;
                    if (GeranceContainer.Height == GeranceContainer.MaximumSize.Height)
                    {
                        GeranceCollapse = false;
                        GeranceTimer.Stop();
                    }
                    CustomerContainer.Height -= 20;
                    QuanlyContainer.Height -= 20;
                    CollapseSettingContainer();
                    CreditContainer.Height -= 20;
                    ReleveContainer.Height -= 20;
                    FraisContainer.Height -= 20;
                    AppartementContainer.Height -= 20;
                }
                else
                {
                    GeranceContainer.Height -= 10;
                    if (GeranceContainer.Height == GeranceContainer.MinimumSize.Height)
                    {
                        GeranceCollapse = true;
                        GeranceTimer.Stop();
                    }
                }
            }
            finally
            {
                sidebar.ResumeLayout(true);
            }
        }
        
        private void CeditTimer_Tick_1(object sender, EventArgs e)
        {   
            sidebar.SuspendLayout();
            try
            {
                if (CeditCollapse)
                {
                    CreditContainer.Height += 10;
                    if (CreditContainer.Height == CreditContainer.MaximumSize.Height)
                    {
                        CeditCollapse = false;
                        CeditTimer.Stop();
                    }
                    CustomerContainer.Height -= 20;
                    QuanlyContainer.Height -= 20;
                    CollapseSettingContainer();
                    ReleveContainer.Height -= 20;
                    FraisContainer.Height -= 20;
                    GeranceContainer.Height -= 20;
                    AppartementContainer.Height -= 20;
                }
                else
                {
                    CreditContainer.Height -= 10;
                    if (CreditContainer.Height == CreditContainer.MinimumSize.Height)
                    {
                        CeditCollapse = true;
                        CeditTimer.Stop();

                    }
                }

            }
            finally
            {
                sidebar.ResumeLayout(true);
            }
        }

        private void btnSetting_Click_1(object sender, EventArgs e)
        {
           /* OpenChildForm(new PKNK_CNPM.Forms.FormSetting(), sender);*/
            //set timer interval to lowest to make it smoother 
            SettingTimer.Start();
        }

        private void btnReleve_Click(object sender, EventArgs e)
        {
            ReleveTimer.Start();
        }

        private void btnFrais_Click(object sender, EventArgs e)
        {
            FraisTimer.Start();
        }

        private void btnGerance_Click(object sender, EventArgs e)
        {
            GeranceTimer.Start();
        }
        private void btnAppartement_Click(object sender, EventArgs e)
        {
            AppartementTimer.Start();
        }

        private void btnCredit_Click(object sender, EventArgs e)
        {
            CeditTimer.Start();
        }

        private void btnSaisirEncaissement_Click(object sender, EventArgs e)
        {
            //OpenChildForm(new frmDanhSachNV(), sender);
        }

        private void btnHoaDon_Click(object sender, EventArgs e)
        {
        }

        private void btnDonXuatHang_Click(object sender, EventArgs e)
        {
            //OpenChildForm(new frmHoaDonXuatNhap(), sender);
        }


        private void btnListeEncaissements_Click(object sender, EventArgs e)
        {
            //OpenChildForm(new frmThongKeVatLieu(), sender);
        }

        private void btnThongKe_Click(object sender, EventArgs e)
        {
            //OpenChildForm(new frmThongKe(), sender);
        }
        
        private void btnModifierLaTaxte_Click(object sender, EventArgs e)
        {
            OpenChildForm(() => new PKNK_CNPM.Forms.FormParamsAvancees(), sender);
        }
        private void btnNouveautrimestere_Click(object sender, EventArgs e)
        {
            OpenChildForm(() => new PKNK_CNPM.Forms.FormArrTrimestere(), sender);
        }
        private void btnGestionCredit_Click(object sender, EventArgs e)
        {
            OpenChildForm(() => new PKNK_CNPM.Forms.FormCredit(), sender);
        }
        private void btnMiseAjourPeriode_Click(object sender, EventArgs e)
        {
            OpenChildForm(() => new PKNK_CNPM.Forms.FormMiseAjourPeriode(), sender);
        }
        private void btnEditerTableVirement_Click(object sender, EventArgs e)
        {
            OpenChildForm(() => new PKNK_CNPM.Forms.FormReleve(), sender);
        }
        private void btnReleveCompte_Click(object sender, EventArgs e)
        {
            OpenChildForm(() => new PKNK_CNPM.Forms.FormReleve(), sender);
        }
        private void btnHistoriqueVirement_Click(object sender, EventArgs e)
        {
            OpenChildForm(() => new PKNK_CNPM.Forms.FormHistoriqueTrimester(), sender);
        }
        private void btnEtatDetaile_Click(object sender, EventArgs e)
        {
            OpenChildForm(() => new PKNK_CNPM.Forms.FormEtatDetaille(), sender);
        }
        private void btnMiseSoldeComission_Click(object sender, EventArgs e)
        {
            OpenChildForm(() => new PKNK_CNPM.Forms.FormMiseSoldeCommission(), sender);
        }
        private void btnListVirements_Click(object sender, EventArgs e)
        {
            OpenChildForm(() => new PKNK_CNPM.Forms.FormListDesVirementBanquaire(), sender);
        }
        private void btnOrderVirement_Click(object sender, EventArgs e)
        {
            OpenChildForm(() => new PKNK_CNPM.Forms.FormOrderVirement(), sender);
        }


        private void btnListeDraisGerance_Click(object sender, EventArgs e)
        {
            //OpenChildForm(new frmDanhSachThuoc(), sender);
        }

        private void btnModifierFrais_Click(object sender, EventArgs e)
        {
            //OpenChildForm(new frmDanhSachThuoc(), sender);
        }

        private void btnSaisirFrais_Click(object sender, EventArgs e)
        {
            OpenChildForm(() => new PKNK_CNPM.Forms.FormFrais(), sender);
        }
        private void btnGestionGerance_Click(object sender, EventArgs e)
        {
            OpenChildForm(() => new PKNK_CNPM.Forms.FormGerance(), sender);
        }
        private void btnAppartementLoure_Click(object sender, EventArgs e)
        {
            OpenChildForm(() => new PKNK_CNPM.Forms.FormLocaleLouer(), sender);
        }
        private void btnAppartementvendre_Click(object sender, EventArgs e)
        {
            OpenChildForm(() => new PKNK_CNPM.Forms.FormLocaleVendr(), sender);
        }
        private void btnCreerGerance_Click(object sender, EventArgs e)
        {
            //OpenChildForm(new frmDanhSachThuoc(), sender);
        }

        private void btnCreateLocataire_Click(object sender, EventArgs e)
        {
            OpenChildForm(() => new PKNK_CNPM.Forms.FormCustomer(), sender);
            //set timer interval to lowest to make it smoother  
        }
        private void btnLocataireGerance_Click(object sender, EventArgs e)
        {
            //OpenChildForm(new frmDanhSachThuoc(), sender);
        }
        private void btnVerifierAugmentation_Click(object sender, EventArgs e)
        {
            OpenChildForm(() => new FormVerifierAugmentation(), sender);
        }
        private void btnListTauxGerance_Click(object sender, EventArgs e)
        {
            OpenChildForm(() => new FormListeTauxGerance(), sender); 
        }
        private void btnGestionParts_Click(object sender, EventArgs e)
        {
            OpenChildForm(() => new PKNK_CNPM.Forms.FormPart(), sender);
        }
        private void btnMisAjourLocataire_Click(object sender, EventArgs e)
        {
            OpenChildForm(() => new PKNK_CNPM.Forms.FormCustomer(), sender);
            //set timer interval to lowest to make it smoother 
        }
        private void btnGestionEncaissement_Click_1(object sender, EventArgs e)
        {
            OpenChildForm(() => new PKNK_CNPM.Forms.FormEncaissement(), sender);
        }
        private void btnListEtatJounalie_Click(object sender, EventArgs e)
        {
            OpenChildForm(() => new PKNK_CNPM.Forms.FormListEtatJournalie(), sender);
        }
        private void btnListeEtatsJournaliers_Click_1(object sender, EventArgs e)
        {
            //OpenChildForm(new PKNK_CNPM.Forms.frmThongKe(), sender);
        }

        // NGODAT


        public DataTable GetDataFromTable(string tableName)
        {
            DataTable dataTable = new DataTable();

            try
            {
                using (OleDbConnection connection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+ AGI_DB.PathDataBase + ";"))
                {
                    connection.Open();
                    // Select only the required columns
                    string query = $"SELECT Id_Notification, Notification, Date_notification FROM {tableName}";
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        OleDbDataAdapter adapter = new OleDbDataAdapter(command);
                        adapter.Fill(dataTable);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }

            return dataTable;
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("vous voulez vraiment quitter ? ", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                this.Close();
            }
        }
    }
}
