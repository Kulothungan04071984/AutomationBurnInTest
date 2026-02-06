using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BurnInTestValidate
{
    public partial class FrmProductSelection : Form
    {
        private readonly IUserService _userService;
        private readonly IServiceProvider _serviceProvider;
        private ComboBox cmbCustomer;
        private ComboBox cmbProductType;
        private ComboBox cmbFGName;
        private Button btnShow;
        public FrmProductSelection(IUserService userService, IServiceProvider serviceProvider)
        {
            InitializeComponent();
            _userService = userService;
            SetupUI();
            LoadCustomer();
            LoadProductTypes();
            _serviceProvider = serviceProvider;
        }
        private void LoadCustomer()
        {
            cmbCustomer.Items.Clear();
            cmbCustomer.Items.Add("Essencore");
            cmbCustomer.SelectedIndex = 0;
        }
        private void SetupUI()
        {

            this.Text = "Product Selection";
            this.Size = new Size(520, 340);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.BackColor = Color.White;

            Panel headerPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 60,
                BackColor = Color.White
            };
            // TITLE
            Label lblTitle = new Label
            {
                Text = "Select Product Details",
                Font = new Font("Segoe UI", 15, FontStyle.Bold),
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter
            };
            // Controls.Add(lblTitle);

          

            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 4,
                Padding = new Padding(30),
                AutoSize = false // IMPORTANT
            };

            //Controls.Add(layout);

            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 60));

            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 45));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 45));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 45));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 60));
            headerPanel.Controls.Add(lblTitle);
            Controls.Add(headerPanel);
            // CUSTOMER
            layout.Controls.Add(CreateLabel("Customer"), 0, 1);
            cmbCustomer = CreateComboBox();
            layout.Controls.Add(cmbCustomer, 1, 1);

            // PRODUCT TYPE
            layout.Controls.Add(CreateLabel("Product Type"), 0, 2);
            cmbProductType = CreateComboBox();
            cmbProductType.SelectedIndexChanged += cmbProductType_SelectedIndexChanged;
            layout.Controls.Add(cmbProductType, 1, 2);

            //// FG NAME
            layout.Controls.Add(CreateLabel("FG Name"), 0, 3);
            cmbFGName = CreateComboBox();
            cmbFGName.Margin = new Padding(0, 15, 0, 0);
            layout.Controls.Add(cmbFGName, 1, 3);

            // BUTTON
            btnShow = new Button
            {
                Text = "SHOW",
                Height = 35,
                Width = 120,
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Anchor = AnchorStyles.None
            };
            btnShow.FlatAppearance.BorderSize = 0;
            btnShow.Click += btnShow_Click;

            layout.Controls.Add(btnShow, 1, 4);

            Controls.Add(layout);
            
        }

        private Label CreateLabel(string text)
        {
            return new Label
            {
                Text = text,
                Font = new Font("Segoe UI", 10),
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleRight
            };
        }
        private ComboBox CreateComboBox()
        {
            return new ComboBox
            {
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 10),
                Margin = new Padding(0, 10, 0, 0),
                DropDownStyle = ComboBoxStyle.DropDownList
                
            };
        }
       

        private void cmbProductType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbProductType.SelectedIndex == 0) return;

            int productTypeId = Convert.ToInt32(cmbProductType.SelectedValue);
            var dt = _userService.GetFGNames(productTypeId);

            cmbFGName.DataSource = dt;
            cmbFGName.DisplayMember = "Fg_Name";
            cmbFGName.ValueMember = "Fg_Name";
        }

        private void btnShow_Click(object sender, EventArgs e)
        {
            if(cmbProductType.SelectedIndex == 0)
            {
                MessageBox.Show("Please select a Product Type.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string customer = cmbCustomer.Text.ToString();
            string productType = cmbProductType.Text.ToString();
            string fgName = cmbFGName.Text.ToString();
            //var burnInForm = _serviceProvider.GetRequiredService<FrmProductSelection>();
            var burnInFormshow = ActivatorUtilities.CreateInstance<FrmBurnIntest>(
    _serviceProvider,
    customer,
    productType,
    fgName

);
            burnInFormshow.Show();
           // var result = _userService.Check_Curr_Stage(string.Empty, "262", "Performance Test", false);
            //MessageBox.Show(
            //    $"Customer: {cmbCustomer.Text}\n" +
            //    $"Product Type: {cmbProductType.Text}\n" +
            //    $"FG Name: {cmbFGName.Text}",
            //    "Selection",
            //    MessageBoxButtons.OK,
            //    MessageBoxIcon.Information);
        }

        private void LoadProductTypes()
        {
            var dt = _userService.GetProductTypes();

            cmbProductType.DataSource = dt;
            cmbProductType.DisplayMember = "ProductName";
            cmbProductType.ValueMember = "ProductID";
        }
    }
}
