namespace CBSys
{
    partial class RibCBSys : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibCBSys()
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
            this.CBSys = this.Factory.CreateRibbonTab();
            this.GrpTipo = this.Factory.CreateRibbonGroup();
            this.drSymbology = this.Factory.CreateRibbonDropDown();
            this.drFormato = this.Factory.CreateRibbonDropDown();
            this.btnRun = this.Factory.CreateRibbonButton();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.CBSys.SuspendLayout();
            this.GrpTipo.SuspendLayout();
            this.SuspendLayout();
            // 
            // CBSys
            // 
            this.CBSys.Groups.Add(this.GrpTipo);
            this.CBSys.Label = "Codigos de Barra";
            this.CBSys.Name = "CBSys";
            // 
            // GrpTipo
            // 
            this.GrpTipo.Items.Add(this.drSymbology);
            this.GrpTipo.Items.Add(this.drFormato);
            this.GrpTipo.Items.Add(this.btnRun);
            this.GrpTipo.Label = "Tipo";
            this.GrpTipo.Name = "GrpTipo";
            // 
            // drSymbology
            // 
            this.drSymbology.Label = "Symbology";
            this.drSymbology.Name = "drSymbology";
            this.drSymbology.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.drSymbology_SelectionChanged);
            // 
            // drFormato
            // 
            this.drFormato.Label = "Formato";
            this.drFormato.Name = "drFormato";
            this.drFormato.Visible = false;
            // 
            // btnRun
            // 
            this.btnRun.Label = "Actualizar Archivos";
            this.btnRun.Name = "btnRun";
            this.btnRun.Visible = false;
            this.btnRun.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRun_Click);
            // 
            // RibCBSys
            // 
            this.Name = "RibCBSys";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.CBSys);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.CBSys.ResumeLayout(false);
            this.CBSys.PerformLayout();
            this.GrpTipo.ResumeLayout(false);
            this.GrpTipo.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup GrpTipo;
        public Microsoft.Office.Tools.Ribbon.RibbonTab CBSys;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drSymbology;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drFormato;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRun;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
    }

    partial class ThisRibbonCollection
    {
        internal RibCBSys Ribbon1
        {
            get { return this.GetRibbon<RibCBSys>(); }
        }
    }
}
