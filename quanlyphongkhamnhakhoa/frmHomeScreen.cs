using PKNK_CNPM;
using PKNK_CNPM.Forms;
using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace PKNK_CNPM
{
    public partial class frmHomeScreen : Form
    {
        private static readonly Font StandardInputFont = new Font("Segoe UI", 9F, FontStyle.Regular);
        private static readonly Font StandardButtonFont = new Font("Segoe UI", 9F, FontStyle.Bold);
        private static readonly Font StandardHeaderFont = new Font("Segoe UI", 9F, FontStyle.Bold);

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
        private int lastContentLayoutSidebarWidth = -1;
        /*public object SiderbarTimer { get; private set; }*/
        //contruction
        
        public frmHomeScreen()
        {
            this.WindowState = FormWindowState.Maximized;
            InitializeComponent();
            this.Icon = new Icon("AGI.ico");
            random = new Random();
            btnCloseChildForm.Visible = false;
            this.Text = string.Empty;
            this.ControlBox = false;
            this.MaximizedBounds = Screen.FromHandle(this.Handle).WorkingArea;

            panelTitleBar.Dock = DockStyle.None;
            panelDesktoppanel.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            grpTatCa.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            this.Resize += (sender, e) => LayoutMainContent(true);
            sidebar.SizeChanged += (sender, e) => LayoutMainContent(false);
            LayoutMainContent(true);
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
        private void StopMenuTimers()
        {
            CustomerTimer.Stop();
            QuanlyTimer.Stop();
            SettingTimer.Stop();
            ReleveTimer.Stop();
            FraisTimer.Stop();
            CeditTimer.Stop();
            GeranceTimer.Stop();
            AppartementTimer.Stop();
        }

        private void ToggleMenu(Panel container, ref bool isCollapsed, Timer timer)
        {
            StopMenuTimers();

            if (container.Height <= container.MinimumSize.Height)
            {
                isCollapsed = true;
            }
            else if (container.Height >= container.MaximumSize.Height)
            {
                isCollapsed = false;
            }
            else
            {
                isCollapsed = !isCollapsed;
            }

            timer.Start();
        }

        private void CollapseOtherMenus(Panel activeMenu)
        {
            Panel[] containers =
            {
                CustomerContainer, QuanlyContainer, SettingContainer, CreditContainer,
                ReleveContainer, FraisContainer, GeranceContainer, AppartementContainer
            };

            foreach (Panel container in containers)
            {
                if (container == activeMenu)
                {
                    continue;
                }

                container.Height = Math.Max(container.MinimumSize.Height, container.Height - 20);
                if (container.Height == container.MinimumSize.Height)
                {
                    SetMenuCollapsed(container);
                }
            }
        }

        private void SetMenuCollapsed(Panel container)
        {
            if (container == CustomerContainer) customerCollapse = true;
            else if (container == QuanlyContainer) quanlyCollapse = true;
            else if (container == SettingContainer) settingCollapse = true;
            else if (container == CreditContainer) CeditCollapse = true;
            else if (container == ReleveContainer) ReleveCollapse = true;
            else if (container == FraisContainer) FraisCollapse = true;
            else if (container == GeranceContainer) GeranceCollapse = true;
            else if (container == AppartementContainer) AppartementCollapse = true;
        }

        private void LayoutMainContent(bool force)
        {
            if (panelTitleBar == null || panelDesktoppanel == null || sidebar == null)
            {
                return;
            }

            bool animationFinished = sidebar.Width == sidebar.MinimumSize.Width ||
                sidebar.Width == sidebar.MaximumSize.Width;
            if (!force && !animationFinished &&
                Math.Abs(sidebar.Width - lastContentLayoutSidebarWidth) < 40)
            {
                return;
            }

            lastContentLayoutSidebarWidth = sidebar.Width;
            int contentLeft = Math.Min(sidebar.Width, ClientSize.Width);
            int contentWidth = Math.Max(0, ClientSize.Width - contentLeft);
            int titleBarHeight = panelTitleBar.Height;

            panelTitleBar.SetBounds(contentLeft, 0, contentWidth, titleBarHeight);
            panelDesktoppanel.SetBounds(
                contentLeft,
                titleBarHeight,
                contentWidth,
                Math.Max(0, ClientSize.Height - titleBarHeight));
        }

        private static void ApplyScreenStyle(Control control)
        {
            Color accentColor = Color.FromArgb(15, 118, 110);
            Color accentHoverColor = Color.FromArgb(17, 94, 89);
            Color textColor = Color.FromArgb(39, 52, 67);

            if (control is Form form)
            {
                form.BackColor = Color.FromArgb(245, 248, 250);
            }
            else if (control is TabPage tabPage)
            {
                tabPage.BackColor = Color.White;
            }

            if (control is DataGridView grid)
            {
                grid.BackgroundColor = Color.White;
                grid.BorderStyle = BorderStyle.FixedSingle;
                grid.GridColor = Color.FromArgb(226, 232, 240);
                grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
                grid.RowHeadersVisible = false;
                grid.EnableHeadersVisualStyles = false;
                grid.ColumnHeadersDefaultCellStyle.BackColor = accentColor;
                grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                grid.ColumnHeadersDefaultCellStyle.Font = StandardHeaderFont;
                grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = accentColor;
                grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                grid.DefaultCellStyle.BackColor = Color.White;
                grid.DefaultCellStyle.ForeColor = textColor;
                grid.DefaultCellStyle.Font = StandardInputFont;
                grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(204, 235, 231);
                grid.DefaultCellStyle.SelectionForeColor = textColor;
                grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 248, 250);
                grid.RowTemplate.Height = Math.Max(grid.RowTemplate.Height, 28);
                grid.ColumnHeadersHeight = Math.Max(grid.ColumnHeadersHeight, 36);
            }
            else if (control is Button button)
            {
                bool isDestructive = button.Name.IndexOf("supprim", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    button.Text.IndexOf("supprim", StringComparison.OrdinalIgnoreCase) >= 0;
                bool isSecondary = button.Name.IndexOf("annul", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    button.Text.IndexOf("annul", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    button.Name.IndexOf("fermer", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    button.Text.IndexOf("fermer", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    button.Name.IndexOf("close", StringComparison.OrdinalIgnoreCase) >= 0;
                Color buttonColor = isDestructive
                    ? Color.FromArgb(185, 55, 55)
                    : isSecondary ? Color.FromArgb(226, 232, 240) : accentColor;
                Color buttonHoverColor = isDestructive
                    ? Color.FromArgb(153, 42, 42)
                    : isSecondary ? Color.FromArgb(203, 213, 225) : accentHoverColor;

                button.FlatStyle = FlatStyle.Flat;
                button.FlatAppearance.BorderColor = isSecondary
                    ? Color.FromArgb(203, 213, 225)
                    : buttonColor;
                button.FlatAppearance.BorderSize = 1;
                button.FlatAppearance.MouseOverBackColor = buttonHoverColor;
                button.FlatAppearance.MouseDownBackColor = isDestructive
                    ? Color.FromArgb(125, 35, 35)
                    : isSecondary ? Color.FromArgb(186, 198, 213) : Color.FromArgb(13, 80, 75);
                button.BackColor = buttonColor;
                button.ForeColor = isSecondary ? textColor : Color.White;
                button.Font = StandardButtonFont;
                button.Cursor = Cursors.Hand;
                button.UseVisualStyleBackColor = false;
            }
            else if (control is TextBox textBox)
            {
                textBox.Font = StandardInputFont;
                textBox.ForeColor = textColor;
                textBox.BorderStyle = BorderStyle.FixedSingle;
            }
            else if (control is ListBox listBox)
            {
                listBox.Font = StandardInputFont;
                listBox.ForeColor = textColor;
                listBox.BackColor = Color.White;
                listBox.BorderStyle = BorderStyle.FixedSingle;
            }
            else if (control is ComboBox comboBox)
            {
                comboBox.Font = StandardInputFont;
                comboBox.ForeColor = textColor;
                comboBox.BackColor = Color.White;
                comboBox.FlatStyle = FlatStyle.Standard;
                ComboBoxBorderWindow borderWindow = new ComboBoxBorderWindow(comboBox);
                comboBox.Disposed += (sender, e) => borderWindow.Dispose();
            }
            else if (control is DateTimePicker dateTimePicker)
            {
                dateTimePicker.Font = StandardInputFont;
                dateTimePicker.CalendarTitleBackColor = accentColor;
            }
            else if (control is TabControl tabControl)
            {
                tabControl.Font = StandardHeaderFont;
            }
            else if (control is GroupBox groupBox)
            {
                groupBox.ForeColor = accentColor;
                groupBox.Font = StandardHeaderFont;
            }
            else if (control is Label label)
            {
                label.ForeColor = textColor;
                if (label.Font.Size >= 14F)
                {
                    label.ForeColor = accentColor;
                }
                else
                {
                    label.Font = StandardInputFont;
                }
            }
            else if (control is CheckBox checkBox)
            {
                checkBox.ForeColor = textColor;
                checkBox.Font = StandardInputFont;
            }
            else if (control is RadioButton radioButton)
            {
                radioButton.ForeColor = textColor;
                radioButton.Font = StandardInputFont;
            }

            foreach (Control child in control.Controls)
            {
                ApplyScreenStyle(child);
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
                ApplyScreenStyle(childForm);
                childForm.TopLevel = false;
                childForm.FormBorderStyle = FormBorderStyle.None;
                childForm.Dock = DockStyle.Fill;
                panelDesktoppanel.Controls.Add(childForm);
                panelDesktoppanel.Tag = childForm;
                childForm.BringToFront();
                childForm.Show();
                lblTitle.Text = childForm.Text;
            }
            finally
            {
                panelDesktoppanel.ResumeLayout(true);
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
            if (WindowState == FormWindowState.Normal)
                this.WindowState = FormWindowState.Maximized;
            else
                this.WindowState = FormWindowState.Maximized;
        }
        private void btnMinimize_Click_1(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void sidebarTimer_Tick(object sender, EventArgs e)
        {
            if (sidebarExpand)
            {
                //if siderbar is Expand, maximine
                sidebar.Width = Math.Max(sidebar.MinimumSize.Width, sidebar.Width - 20);
                if (sidebar.Width <= sidebar.MinimumSize.Width)
                {
                    sidebar.Width = sidebar.MinimumSize.Width;
                    sidebarExpand = false;
                    sidebarTimer.Stop();
                    LayoutMainContent(true);
                }
            }
            else
            {
                //if siderbar is Expand, minimine
                sidebar.Width = Math.Min(sidebar.MaximumSize.Width, sidebar.Width + 20);
                if (sidebar.Width >= sidebar.MaximumSize.Width)
                {
                    sidebar.Width = sidebar.MaximumSize.Width;
                    sidebarExpand = true;
                    sidebarTimer.Stop();
                    LayoutMainContent(true);
                }
            }
        }
        

        private void menuButton_Click_1(object sender, EventArgs e)
        {
            //set timer interval to lowest to make it smoother 
            sidebarTimer.Start();
        }

        //Set Timer các chức năng trong toolBar
        private void btnCustomer_Click_1(object sender, EventArgs e)
        {
            ToggleMenu(CustomerContainer, ref customerCollapse, CustomerTimer);
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
                    CollapseOtherMenus(CustomerContainer);

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
                    CollapseOtherMenus(QuanlyContainer);
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
            ToggleMenu(QuanlyContainer, ref quanlyCollapse, QuanlyTimer);
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
                    CollapseOtherMenus(SettingContainer);
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
                    CollapseOtherMenus(ReleveContainer);
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
                    CollapseOtherMenus(FraisContainer);
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
                    CollapseOtherMenus(AppartementContainer);
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
                    CollapseOtherMenus(GeranceContainer);
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
                    CollapseOtherMenus(CreditContainer);
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
            ToggleMenu(SettingContainer, ref settingCollapse, SettingTimer);
        }

        private void btnReleve_Click(object sender, EventArgs e)
        {
            ToggleMenu(ReleveContainer, ref ReleveCollapse, ReleveTimer);
        }

        private void btnFrais_Click(object sender, EventArgs e)
        {
            ToggleMenu(FraisContainer, ref FraisCollapse, FraisTimer);
        }

        private void btnGerance_Click(object sender, EventArgs e)
        {
            ToggleMenu(GeranceContainer, ref GeranceCollapse, GeranceTimer);
        }
        private void btnAppartement_Click(object sender, EventArgs e)
        {
            ToggleMenu(AppartementContainer, ref AppartementCollapse, AppartementTimer);
        }

        private void btnCredit_Click(object sender, EventArgs e)
        {
            ToggleMenu(CreditContainer, ref CeditCollapse, CeditTimer);
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

        private sealed class ComboBoxBorderWindow : NativeWindow, IDisposable
        {
            private const int WmPaint = 0x000F;
            private readonly ComboBox _comboBox;

            public ComboBoxBorderWindow(ComboBox comboBox)
            {
                _comboBox = comboBox;
                _comboBox.HandleCreated += ComboBox_HandleCreated;
                _comboBox.HandleDestroyed += ComboBox_HandleDestroyed;
                if (_comboBox.IsHandleCreated)
                {
                    AssignHandle(_comboBox.Handle);
                }
            }

            protected override void WndProc(ref Message message)
            {
                base.WndProc(ref message);

                if (message.Msg == WmPaint && !_comboBox.IsDisposed)
                {
                    int width = _comboBox.ClientSize.Width;
                    int height = _comboBox.ClientSize.Height;
                    if (width > 1 && height > 1)
                    {
                        using (Graphics graphics = Graphics.FromHwnd(message.HWnd))
                        using (Pen borderPen = new Pen(Color.Black))
                        {
                            graphics.DrawRectangle(borderPen, 0, 0, width - 1, height - 1);
                        }
                    }
                }
            }

            public void Dispose()
            {
                _comboBox.HandleCreated -= ComboBox_HandleCreated;
                _comboBox.HandleDestroyed -= ComboBox_HandleDestroyed;
                if (Handle != IntPtr.Zero)
                {
                    ReleaseHandle();
                }
            }

            private void ComboBox_HandleCreated(object sender, EventArgs e)
            {
                AssignHandle(_comboBox.Handle);
            }

            private void ComboBox_HandleDestroyed(object sender, EventArgs e)
            {
                if (Handle != IntPtr.Zero)
                {
                    ReleaseHandle();
                }
            }
        }
    }
}
