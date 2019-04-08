namespace OutlookAddInDecoderRing
{
    partial class RibbonDemo : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonDemo()
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
            this.tabDecode = this.Factory.CreateRibbonTab();
            this.Decode = this.Factory.CreateRibbonGroup();
            this.btnDecodeMessage = this.Factory.CreateRibbonButton();
            this.tabDecode.SuspendLayout();
            this.Decode.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabDecode
            // 
            this.tabDecode.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabDecode.Groups.Add(this.Decode);
            this.tabDecode.Label = "Decode Tab";
            this.tabDecode.Name = "tabDecode";
            // 
            // Decode
            // 
            this.Decode.Items.Add(this.btnDecodeMessage);
            this.Decode.Label = "Decode";
            this.Decode.Name = "Decode";
            // 
            // btnDecodeMessage
            // 
            this.btnDecodeMessage.Label = "Decode Message";
            this.btnDecodeMessage.Name = "btnDecodeMessage";
            this.btnDecodeMessage.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDecodeMessage_Click);
            // 
            // RibbonDemo
            // 
            this.Name = "RibbonDemo";
            this.RibbonType = "Microsoft.Outlook.Mail.Read";
            this.Tabs.Add(this.tabDecode);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tabDecode.ResumeLayout(false);
            this.tabDecode.PerformLayout();
            this.Decode.ResumeLayout(false);
            this.Decode.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabDecode;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Decode;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDecodeMessage;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonDemo Ribbon1
        {
            get { return this.GetRibbon<RibbonDemo>(); }
        }
    }
}
