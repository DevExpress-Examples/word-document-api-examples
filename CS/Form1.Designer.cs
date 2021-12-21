namespace RichEditDocumentServerAPIExample
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.sidePanel1 = new DevExpress.XtraEditors.SidePanel();
            this.treeList1 = new DevExpress.XtraTreeList.TreeList();
            this.sidePanel2 = new DevExpress.XtraEditors.SidePanel();
            this.codeExampleNameLbl = new DevExpress.XtraEditors.LabelControl();
            this.sidePanel3 = new DevExpress.XtraEditors.SidePanel();
            this.simpleButton1 = new DevExpress.XtraEditors.SimpleButton();
            this.sidePanel4 = new DevExpress.XtraEditors.SidePanel();
            this.sidePanel5 = new DevExpress.XtraEditors.SidePanel();
            this.sidePanel6 = new DevExpress.XtraEditors.SidePanel();
            this.xtraTabControl1 = new DevExpress.XtraTab.XtraTabControl();
            this.xtraTabPage1 = new DevExpress.XtraTab.XtraTabPage();
            this.richEditControlCS = new DevExpress.XtraRichEdit.RichEditControl();
            this.xtraTabPage2 = new DevExpress.XtraTab.XtraTabPage();
            this.richEditControlVB = new DevExpress.XtraRichEdit.RichEditControl();
            this.sidePanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.treeList1)).BeginInit();
            this.sidePanel2.SuspendLayout();
            this.sidePanel3.SuspendLayout();
            this.sidePanel6.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.xtraTabControl1)).BeginInit();
            this.xtraTabControl1.SuspendLayout();
            this.xtraTabPage1.SuspendLayout();
            this.xtraTabPage2.SuspendLayout();
            this.SuspendLayout();
            // 
            // sidePanel1
            // 
            this.sidePanel1.Controls.Add(this.treeList1);
            this.sidePanel1.Dock = System.Windows.Forms.DockStyle.Right;
            this.sidePanel1.Location = new System.Drawing.Point(1322, 0);
            this.sidePanel1.MinimumSize = new System.Drawing.Size(100, 600);
            this.sidePanel1.Name = "sidePanel1";
            this.sidePanel1.Size = new System.Drawing.Size(469, 1258);
            this.sidePanel1.TabIndex = 0;
            this.sidePanel1.Text = "sidePanel1";
            // 
            // treeList1
            // 
            this.treeList1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeList1.Location = new System.Drawing.Point(2, 0);
            this.treeList1.Name = "treeList1";
            this.treeList1.Size = new System.Drawing.Size(467, 1258);
            this.treeList1.TabIndex = 0;
            // 
            // sidePanel2
            // 
            this.sidePanel2.AllowResize = false;
            this.sidePanel2.Controls.Add(this.codeExampleNameLbl);
            this.sidePanel2.Dock = System.Windows.Forms.DockStyle.Top;
            this.sidePanel2.Location = new System.Drawing.Point(0, 0);
            this.sidePanel2.Name = "sidePanel2";
            this.sidePanel2.Size = new System.Drawing.Size(1322, 167);
            this.sidePanel2.TabIndex = 1;
            this.sidePanel2.Text = "sidePanel2";
            // 
            // codeExampleNameLbl
            // 
            this.codeExampleNameLbl.Appearance.Font = new System.Drawing.Font("Tahoma", 16.125F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.codeExampleNameLbl.Appearance.Options.UseFont = true;
            this.codeExampleNameLbl.Location = new System.Drawing.Point(55, 58);
            this.codeExampleNameLbl.Name = "codeExampleNameLbl";
            this.codeExampleNameLbl.Size = new System.Drawing.Size(446, 52);
            this.codeExampleNameLbl.TabIndex = 0;
            this.codeExampleNameLbl.Text = "Examples Not Found";
            // 
            // sidePanel3
            // 
            this.sidePanel3.AllowResize = false;
            this.sidePanel3.Controls.Add(this.simpleButton1);
            this.sidePanel3.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.sidePanel3.Location = new System.Drawing.Point(0, 1091);
            this.sidePanel3.Name = "sidePanel3";
            this.sidePanel3.Size = new System.Drawing.Size(1322, 167);
            this.sidePanel3.TabIndex = 2;
            this.sidePanel3.Text = "sidePanel3";
            // 
            // simpleButton1
            // 
            this.simpleButton1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.simpleButton1.Location = new System.Drawing.Point(1072, 22);
            this.simpleButton1.Name = "simpleButton1";
            this.simpleButton1.Size = new System.Drawing.Size(228, 78);
            this.simpleButton1.TabIndex = 0;
            this.simpleButton1.Text = "Run";
            this.simpleButton1.Click += new System.EventHandler(this.OnRunButtonClick);
            // 
            // sidePanel4
            // 
            this.sidePanel4.AllowResize = false;
            this.sidePanel4.Dock = System.Windows.Forms.DockStyle.Right;
            this.sidePanel4.Location = new System.Drawing.Point(1302, 167);
            this.sidePanel4.Name = "sidePanel4";
            this.sidePanel4.Size = new System.Drawing.Size(20, 924);
            this.sidePanel4.TabIndex = 3;
            this.sidePanel4.Text = "sidePanel4";
            // 
            // sidePanel5
            // 
            this.sidePanel5.AllowResize = false;
            this.sidePanel5.Dock = System.Windows.Forms.DockStyle.Left;
            this.sidePanel5.Location = new System.Drawing.Point(0, 167);
            this.sidePanel5.Name = "sidePanel5";
            this.sidePanel5.Size = new System.Drawing.Size(18, 924);
            this.sidePanel5.TabIndex = 4;
            this.sidePanel5.Text = "sidePanel5";
            // 
            // sidePanel6
            // 
            this.sidePanel6.Controls.Add(this.xtraTabControl1);
            this.sidePanel6.Dock = System.Windows.Forms.DockStyle.Fill;
            this.sidePanel6.Location = new System.Drawing.Point(18, 167);
            this.sidePanel6.MinimumSize = new System.Drawing.Size(140, 180);
            this.sidePanel6.Name = "sidePanel6";
            this.sidePanel6.Size = new System.Drawing.Size(1284, 924);
            this.sidePanel6.TabIndex = 5;
            this.sidePanel6.Text = "sidePanel6";
            // 
            // xtraTabControl1
            // 
            this.xtraTabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.xtraTabControl1.Location = new System.Drawing.Point(0, 0);
            this.xtraTabControl1.Name = "xtraTabControl1";
            this.xtraTabControl1.SelectedTabPage = this.xtraTabPage1;
            this.xtraTabControl1.Size = new System.Drawing.Size(1284, 924);
            this.xtraTabControl1.TabIndex = 0;
            this.xtraTabControl1.TabPages.AddRange(new DevExpress.XtraTab.XtraTabPage[] {
            this.xtraTabPage1,
            this.xtraTabPage2});
            // 
            // xtraTabPage1
            // 
            this.xtraTabPage1.Controls.Add(this.richEditControlCS);
            this.xtraTabPage1.Name = "xtraTabPage1";
            this.xtraTabPage1.Size = new System.Drawing.Size(1280, 875);
            this.xtraTabPage1.Tag = "CS";
            this.xtraTabPage1.Text = "C#";
            // 
            // richEditControlCS
            // 
            this.richEditControlCS.Dock = System.Windows.Forms.DockStyle.Fill;
            this.richEditControlCS.Location = new System.Drawing.Point(0, 0);
            this.richEditControlCS.Name = "richEditControlCS";
            this.richEditControlCS.Options.HorizontalRuler.Visibility = DevExpress.XtraRichEdit.RichEditRulerVisibility.Hidden;
            this.richEditControlCS.Options.VerticalRuler.Visibility = DevExpress.XtraRichEdit.RichEditRulerVisibility.Hidden;
            this.richEditControlCS.ReadOnly = true;
            this.richEditControlCS.Size = new System.Drawing.Size(1280, 875);
            this.richEditControlCS.TabIndex = 0;
            // 
            // xtraTabPage2
            // 
            this.xtraTabPage2.Controls.Add(this.richEditControlVB);
            this.xtraTabPage2.Name = "xtraTabPage2";
            this.xtraTabPage2.Size = new System.Drawing.Size(1280, 875);
            this.xtraTabPage2.Tag = "VB";
            this.xtraTabPage2.Text = "VB";
            // 
            // richEditControlVB
            // 
            this.richEditControlVB.Dock = System.Windows.Forms.DockStyle.Fill;
            this.richEditControlVB.Location = new System.Drawing.Point(0, 0);
            this.richEditControlVB.Name = "richEditControlVB";
            this.richEditControlVB.Options.HorizontalRuler.Visibility = DevExpress.XtraRichEdit.RichEditRulerVisibility.Hidden;
            this.richEditControlVB.Options.VerticalRuler.Visibility = DevExpress.XtraRichEdit.RichEditRulerVisibility.Hidden;
            this.richEditControlVB.ReadOnly = true;
            this.richEditControlVB.Size = new System.Drawing.Size(1280, 875);
            this.richEditControlVB.TabIndex = 0;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1791, 1258);
            this.Controls.Add(this.sidePanel6);
            this.Controls.Add(this.sidePanel5);
            this.Controls.Add(this.sidePanel4);
            this.Controls.Add(this.sidePanel3);
            this.Controls.Add(this.sidePanel2);
            this.Controls.Add(this.sidePanel1);
            this.Margin = new System.Windows.Forms.Padding(6);
            this.MinimumSize = new System.Drawing.Size(765, 560);
            this.Name = "Form1";
            this.Text = "Form1";
            this.sidePanel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.treeList1)).EndInit();
            this.sidePanel2.ResumeLayout(false);
            this.sidePanel2.PerformLayout();
            this.sidePanel3.ResumeLayout(false);
            this.sidePanel6.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.xtraTabControl1)).EndInit();
            this.xtraTabControl1.ResumeLayout(false);
            this.xtraTabPage1.ResumeLayout(false);
            this.xtraTabPage2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraEditors.SidePanel sidePanel1;
        private DevExpress.XtraTreeList.TreeList treeList1;
        private DevExpress.XtraEditors.SidePanel sidePanel2;
        private DevExpress.XtraEditors.LabelControl codeExampleNameLbl;
        private DevExpress.XtraEditors.SidePanel sidePanel3;
        private DevExpress.XtraEditors.SimpleButton simpleButton1;
        private DevExpress.XtraEditors.SidePanel sidePanel4;
        private DevExpress.XtraEditors.SidePanel sidePanel5;
        private DevExpress.XtraEditors.SidePanel sidePanel6;
        private DevExpress.XtraTab.XtraTabControl xtraTabControl1;
        private DevExpress.XtraTab.XtraTabPage xtraTabPage1;
        private DevExpress.XtraTab.XtraTabPage xtraTabPage2;
        private DevExpress.XtraRichEdit.RichEditControl richEditControlCS;
        private DevExpress.XtraRichEdit.RichEditControl richEditControlVB;
    }
}

