
namespace BessolitsynExcelToolAdd_in
{
    partial class RibbonBessolitsyn : Microsoft.Office.Tools.Ribbon.RibbonBase
    {

        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonBessolitsyn()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.button17 = this.Factory.CreateRibbonButton();
            this.button16 = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.button2 = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.button4 = this.Factory.CreateRibbonButton();
            this.button5 = this.Factory.CreateRibbonButton();
            this.button6 = this.Factory.CreateRibbonButton();
            this.button7 = this.Factory.CreateRibbonButton();
            this.dropDown1 = this.Factory.CreateRibbonDropDown();
            this.button8 = this.Factory.CreateRibbonButton();
            this.checkBox1 = this.Factory.CreateRibbonCheckBox();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.button10 = this.Factory.CreateRibbonButton();
            this.button9 = this.Factory.CreateRibbonButton();
            this.editBox1 = this.Factory.CreateRibbonEditBox();
            this.separator4 = this.Factory.CreateRibbonSeparator();
            this.button11 = this.Factory.CreateRibbonButton();
            this.button12 = this.Factory.CreateRibbonButton();
            this.button13 = this.Factory.CreateRibbonButton();
            this.separator5 = this.Factory.CreateRibbonSeparator();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button14 = this.Factory.CreateRibbonButton();
            this.checkBox2 = this.Factory.CreateRibbonCheckBox();
            this.checkBox3 = this.Factory.CreateRibbonCheckBox();
            this.checkBox4 = this.Factory.CreateRibbonCheckBox();
            this.button15 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button1);
            this.group1.Items.Add(this.button17);
            this.group1.Items.Add(this.button16);
            this.group1.Items.Add(this.separator1);
            this.group1.Items.Add(this.button2);
            this.group1.Items.Add(this.button3);
            this.group1.Items.Add(this.button4);
            this.group1.Items.Add(this.button5);
            this.group1.Items.Add(this.button6);
            this.group1.Items.Add(this.button7);
            this.group1.Items.Add(this.dropDown1);
            this.group1.Items.Add(this.button8);
            this.group1.Items.Add(this.checkBox1);
            this.group1.Items.Add(this.separator3);
            this.group1.Items.Add(this.button10);
            this.group1.Items.Add(this.button9);
            this.group1.Items.Add(this.editBox1);
            this.group1.Items.Add(this.separator4);
            this.group1.Items.Add(this.button11);
            this.group1.Items.Add(this.button12);
            this.group1.Items.Add(this.button13);
            this.group1.Items.Add(this.separator5);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // button1
            // 
            this.button1.Label = "IntersectionR1...Rn";
            this.button1.Name = "button1";
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.IntersectionR1Rn_Click);
            // 
            // button17
            // 
            this.button17.Label = "SumR1...Rn";
            this.button17.Name = "button17";
            this.button17.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SumR1Rn_Click);
            // 
            // button16
            // 
            this.button16.Label = "Paste Results";
            this.button16.Name = "button16";
            this.button16.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button16_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // button2
            // 
            this.button2.Label = "Set R1";
            this.button2.Name = "button2";
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Label = "Set R2";
            this.button3.Name = "button3";
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.Label = "Run";
            this.button4.Name = "button4";
            this.button4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button4_Click);
            // 
            // button5
            // 
            this.button5.Label = "R1-R2";
            this.button5.Name = "button5";
            this.button5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button5_Click);
            // 
            // button6
            // 
            this.button6.Label = "R2-R1";
            this.button6.Name = "button6";
            this.button6.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button6_Click);
            // 
            // button7
            // 
            this.button7.Label = "Intersection(R1,R2)";
            this.button7.Name = "button7";
            this.button7.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button7_Click);
            // 
            // dropDown1
            // 
            ribbonDropDownItemImpl1.Label = "R1-R2";
            ribbonDropDownItemImpl2.Label = "R2-R1";
            ribbonDropDownItemImpl3.Label = "Intersection";
            this.dropDown1.Items.Add(ribbonDropDownItemImpl1);
            this.dropDown1.Items.Add(ribbonDropDownItemImpl2);
            this.dropDown1.Items.Add(ribbonDropDownItemImpl3);
            this.dropDown1.Label = "dropDown1";
            this.dropDown1.Name = "dropDown1";
            // 
            // button8
            // 
            this.button8.Label = "Paste results";
            this.button8.Name = "button8";
            this.button8.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button8_Click);
            // 
            // checkBox1
            // 
            this.checkBox1.Label = "Highlightb Results";
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBox1_Click);
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            // 
            // button10
            // 
            this.button10.Label = "SetRange[ID,Value]";
            this.button10.Name = "button10";
            this.button10.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button10_Click);
            // 
            // button9
            // 
            this.button9.Label = "Get Value";
            this.button9.Name = "button9";
            this.button9.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button9_Click);
            // 
            // editBox1
            // 
            this.editBox1.Label = "+NnextColumns";
            this.editBox1.Name = "editBox1";
            this.editBox1.Text = "1";
            // 
            // separator4
            // 
            this.separator4.Name = "separator4";
            // 
            // button11
            // 
            this.button11.Label = "SetRange[A,B]";
            this.button11.Name = "button11";
            this.button11.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button11_Click);
            // 
            // button12
            // 
            this.button12.Label = "Get A ";
            this.button12.Name = "button12";
            this.button12.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button12_Click);
            // 
            // button13
            // 
            this.button13.Label = "Get B";
            this.button13.Name = "button13";
            this.button13.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button13_Click);
            // 
            // separator5
            // 
            this.separator5.Name = "separator5";
            // 
            // group2
            // 
            this.group2.Items.Add(this.button14);
            this.group2.Items.Add(this.checkBox2);
            this.group2.Items.Add(this.checkBox3);
            this.group2.Items.Add(this.checkBox4);
            this.group2.Items.Add(this.button15);
            this.group2.Label = "ECE2 Phase";
            this.group2.Name = "group2";
            // 
            // button14
            // 
            this.button14.Label = "Read PBS2 Matrix";
            this.button14.Name = "button14";
            this.button14.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button14_Click);
            // 
            // checkBox2
            // 
            this.checkBox2.Label = "Attributes in L2";
            this.checkBox2.Name = "checkBox2";
            // 
            // checkBox3
            // 
            this.checkBox3.Label = "Attributes in L3";
            this.checkBox3.Name = "checkBox3";
            // 
            // checkBox4
            // 
            this.checkBox4.Label = "Attributes in L4";
            this.checkBox4.Name = "checkBox4";
            // 
            // button15
            // 
            this.button15.Label = "Paste Results";
            this.button15.Name = "button15";
            this.button15.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button15_Click);
            // 
            // RibbonBessolitsyn
            // 
            this.Name = "RibbonBessolitsyn";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonBessolitsyn_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button7;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDown1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button9;
        public Microsoft.Office.Tools.Ribbon.RibbonButton button8;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button10;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button11;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button12;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button13;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button14;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button15;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox1;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox2;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox3;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox4;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button16;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button17;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox1;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonBessolitsyn RibbonBessolitsyn
        {
            get { return this.GetRibbon<RibbonBessolitsyn>(); }
        }
    }
}
