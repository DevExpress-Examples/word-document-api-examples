Namespace RichEditDocumentServerAPIExample

    Partial Class Form1

        ''' <summary>
        ''' Required designer variable.
        ''' </summary>
        Private components As System.ComponentModel.IContainer = Nothing

        ''' <summary>
        ''' Clean up any resources being used.
        ''' </summary>
        ''' <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing AndAlso (Me.components IsNot Nothing) Then
                Me.components.Dispose()
            End If

            MyBase.Dispose(disposing)
        End Sub

#Region "Windows Form Designer generated code"
        ''' <summary>
        ''' Required method for Designer support - do not modify
        ''' the contents of this method with the code editor.
        ''' </summary>
        Private Sub InitializeComponent()
            Me.sidePanel1 = New DevExpress.XtraEditors.SidePanel()
            Me.treeList1 = New DevExpress.XtraTreeList.TreeList()
            Me.sidePanel2 = New DevExpress.XtraEditors.SidePanel()
            Me.codeExampleNameLbl = New DevExpress.XtraEditors.LabelControl()
            Me.sidePanel3 = New DevExpress.XtraEditors.SidePanel()
            Me.simpleButton1 = New DevExpress.XtraEditors.SimpleButton()
            Me.sidePanel4 = New DevExpress.XtraEditors.SidePanel()
            Me.sidePanel5 = New DevExpress.XtraEditors.SidePanel()
            Me.sidePanel6 = New DevExpress.XtraEditors.SidePanel()
            Me.xtraTabControl1 = New DevExpress.XtraTab.XtraTabControl()
            Me.xtraTabPage1 = New DevExpress.XtraTab.XtraTabPage()
            Me.richEditControlCS = New DevExpress.XtraRichEdit.RichEditControl()
            Me.xtraTabPage2 = New DevExpress.XtraTab.XtraTabPage()
            Me.richEditControlVB = New DevExpress.XtraRichEdit.RichEditControl()
            Me.sidePanel1.SuspendLayout()
            CType((Me.treeList1), System.ComponentModel.ISupportInitialize).BeginInit()
            Me.sidePanel2.SuspendLayout()
            Me.sidePanel3.SuspendLayout()
            Me.sidePanel6.SuspendLayout()
            CType((Me.xtraTabControl1), System.ComponentModel.ISupportInitialize).BeginInit()
            Me.xtraTabControl1.SuspendLayout()
            Me.xtraTabPage1.SuspendLayout()
            Me.xtraTabPage2.SuspendLayout()
            Me.SuspendLayout()
            ' 
            ' sidePanel1
            ' 
            Me.sidePanel1.Controls.Add(Me.treeList1)
            Me.sidePanel1.Dock = System.Windows.Forms.DockStyle.Right
            Me.sidePanel1.Location = New System.Drawing.Point(1322, 0)
            Me.sidePanel1.MinimumSize = New System.Drawing.Size(100, 600)
            Me.sidePanel1.Name = "sidePanel1"
            Me.sidePanel1.Size = New System.Drawing.Size(469, 1258)
            Me.sidePanel1.TabIndex = 0
            Me.sidePanel1.Text = "sidePanel1"
            ' 
            ' treeList1
            ' 
            Me.treeList1.Dock = System.Windows.Forms.DockStyle.Fill
            Me.treeList1.Location = New System.Drawing.Point(2, 0)
            Me.treeList1.Name = "treeList1"
            Me.treeList1.Size = New System.Drawing.Size(467, 1258)
            Me.treeList1.TabIndex = 0
            ' 
            ' sidePanel2
            ' 
            Me.sidePanel2.AllowResize = False
            Me.sidePanel2.Controls.Add(Me.codeExampleNameLbl)
            Me.sidePanel2.Dock = System.Windows.Forms.DockStyle.Top
            Me.sidePanel2.Location = New System.Drawing.Point(0, 0)
            Me.sidePanel2.Name = "sidePanel2"
            Me.sidePanel2.Size = New System.Drawing.Size(1322, 167)
            Me.sidePanel2.TabIndex = 1
            Me.sidePanel2.Text = "sidePanel2"
            ' 
            ' codeExampleNameLbl
            ' 
            Me.codeExampleNameLbl.Appearance.Font = New System.Drawing.Font("Tahoma", 16.125F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte((0))))
            Me.codeExampleNameLbl.Appearance.Options.UseFont = True
            Me.codeExampleNameLbl.Location = New System.Drawing.Point(55, 58)
            Me.codeExampleNameLbl.Name = "codeExampleNameLbl"
            Me.codeExampleNameLbl.Size = New System.Drawing.Size(446, 52)
            Me.codeExampleNameLbl.TabIndex = 0
            Me.codeExampleNameLbl.Text = "Examples Not Found"
            ' 
            ' sidePanel3
            ' 
            Me.sidePanel3.AllowResize = False
            Me.sidePanel3.Controls.Add(Me.simpleButton1)
            Me.sidePanel3.Dock = System.Windows.Forms.DockStyle.Bottom
            Me.sidePanel3.Location = New System.Drawing.Point(0, 1091)
            Me.sidePanel3.Name = "sidePanel3"
            Me.sidePanel3.Size = New System.Drawing.Size(1322, 167)
            Me.sidePanel3.TabIndex = 2
            Me.sidePanel3.Text = "sidePanel3"
            ' 
            ' simpleButton1
            ' 
            Me.simpleButton1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)), System.Windows.Forms.AnchorStyles)
            Me.simpleButton1.Location = New System.Drawing.Point(1072, 22)
            Me.simpleButton1.Name = "simpleButton1"
            Me.simpleButton1.Size = New System.Drawing.Size(228, 78)
            Me.simpleButton1.TabIndex = 0
            Me.simpleButton1.Text = "Run"
            AddHandler Me.simpleButton1.Click, New System.EventHandler(AddressOf Me.OnRunButtonClick)
            ' 
            ' sidePanel4
            ' 
            Me.sidePanel4.AllowResize = False
            Me.sidePanel4.Dock = System.Windows.Forms.DockStyle.Right
            Me.sidePanel4.Location = New System.Drawing.Point(1302, 167)
            Me.sidePanel4.Name = "sidePanel4"
            Me.sidePanel4.Size = New System.Drawing.Size(20, 924)
            Me.sidePanel4.TabIndex = 3
            Me.sidePanel4.Text = "sidePanel4"
            ' 
            ' sidePanel5
            ' 
            Me.sidePanel5.AllowResize = False
            Me.sidePanel5.Dock = System.Windows.Forms.DockStyle.Left
            Me.sidePanel5.Location = New System.Drawing.Point(0, 167)
            Me.sidePanel5.Name = "sidePanel5"
            Me.sidePanel5.Size = New System.Drawing.Size(18, 924)
            Me.sidePanel5.TabIndex = 4
            Me.sidePanel5.Text = "sidePanel5"
            ' 
            ' sidePanel6
            ' 
            Me.sidePanel6.Controls.Add(Me.xtraTabControl1)
            Me.sidePanel6.Dock = System.Windows.Forms.DockStyle.Fill
            Me.sidePanel6.Location = New System.Drawing.Point(18, 167)
            Me.sidePanel6.MinimumSize = New System.Drawing.Size(140, 180)
            Me.sidePanel6.Name = "sidePanel6"
            Me.sidePanel6.Size = New System.Drawing.Size(1284, 924)
            Me.sidePanel6.TabIndex = 5
            Me.sidePanel6.Text = "sidePanel6"
            ' 
            ' xtraTabControl1
            ' 
            Me.xtraTabControl1.Dock = System.Windows.Forms.DockStyle.Fill
            Me.xtraTabControl1.Location = New System.Drawing.Point(0, 0)
            Me.xtraTabControl1.Name = "xtraTabControl1"
            Me.xtraTabControl1.SelectedTabPage = Me.xtraTabPage1
            Me.xtraTabControl1.Size = New System.Drawing.Size(1284, 924)
            Me.xtraTabControl1.TabIndex = 0
            Me.xtraTabControl1.TabPages.AddRange(New DevExpress.XtraTab.XtraTabPage() {Me.xtraTabPage1, Me.xtraTabPage2})
            ' 
            ' xtraTabPage1
            ' 
            Me.xtraTabPage1.Controls.Add(Me.richEditControlCS)
            Me.xtraTabPage1.Name = "xtraTabPage1"
            Me.xtraTabPage1.Size = New System.Drawing.Size(1280, 875)
            Me.xtraTabPage1.Tag = "CS"
            Me.xtraTabPage1.Text = "C#"
            ' 
            ' richEditControlCS
            ' 
            Me.richEditControlCS.Dock = System.Windows.Forms.DockStyle.Fill
            Me.richEditControlCS.Location = New System.Drawing.Point(0, 0)
            Me.richEditControlCS.Name = "richEditControlCS"
            Me.richEditControlCS.Options.HorizontalRuler.Visibility = DevExpress.XtraRichEdit.RichEditRulerVisibility.Hidden
            Me.richEditControlCS.Options.VerticalRuler.Visibility = DevExpress.XtraRichEdit.RichEditRulerVisibility.Hidden
            Me.richEditControlCS.[ReadOnly] = True
            Me.richEditControlCS.Size = New System.Drawing.Size(1280, 875)
            Me.richEditControlCS.TabIndex = 0
            ' 
            ' xtraTabPage2
            ' 
            Me.xtraTabPage2.Controls.Add(Me.richEditControlVB)
            Me.xtraTabPage2.Name = "xtraTabPage2"
            Me.xtraTabPage2.Size = New System.Drawing.Size(1280, 875)
            Me.xtraTabPage2.Tag = "VB"
            Me.xtraTabPage2.Text = "VB"
            ' 
            ' richEditControlVB
            ' 
            Me.richEditControlVB.Dock = System.Windows.Forms.DockStyle.Fill
            Me.richEditControlVB.Location = New System.Drawing.Point(0, 0)
            Me.richEditControlVB.Name = "richEditControlVB"
            Me.richEditControlVB.Options.HorizontalRuler.Visibility = DevExpress.XtraRichEdit.RichEditRulerVisibility.Hidden
            Me.richEditControlVB.Options.VerticalRuler.Visibility = DevExpress.XtraRichEdit.RichEditRulerVisibility.Hidden
            Me.richEditControlVB.[ReadOnly] = True
            Me.richEditControlVB.Size = New System.Drawing.Size(1280, 875)
            Me.richEditControlVB.TabIndex = 0
            ' 
            ' Form1
            ' 
            Me.AutoScaleDimensions = New System.Drawing.SizeF(12F, 25F)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(1791, 1258)
            Me.Controls.Add(Me.sidePanel6)
            Me.Controls.Add(Me.sidePanel5)
            Me.Controls.Add(Me.sidePanel4)
            Me.Controls.Add(Me.sidePanel3)
            Me.Controls.Add(Me.sidePanel2)
            Me.Controls.Add(Me.sidePanel1)
            Me.Margin = New System.Windows.Forms.Padding(6)
            Me.MinimumSize = New System.Drawing.Size(765, 560)
            Me.Name = "Form1"
            Me.Text = "Form1"
            Me.sidePanel1.ResumeLayout(False)
            CType((Me.treeList1), System.ComponentModel.ISupportInitialize).EndInit()
            Me.sidePanel2.ResumeLayout(False)
            Me.sidePanel2.PerformLayout()
            Me.sidePanel3.ResumeLayout(False)
            Me.sidePanel6.ResumeLayout(False)
            CType((Me.xtraTabControl1), System.ComponentModel.ISupportInitialize).EndInit()
            Me.xtraTabControl1.ResumeLayout(False)
            Me.xtraTabPage1.ResumeLayout(False)
            Me.xtraTabPage2.ResumeLayout(False)
            Me.ResumeLayout(False)
        End Sub

#End Region
        Private sidePanel1 As DevExpress.XtraEditors.SidePanel

        Private treeList1 As DevExpress.XtraTreeList.TreeList

        Private sidePanel2 As DevExpress.XtraEditors.SidePanel

        Private codeExampleNameLbl As DevExpress.XtraEditors.LabelControl

        Private sidePanel3 As DevExpress.XtraEditors.SidePanel

        Private simpleButton1 As DevExpress.XtraEditors.SimpleButton

        Private sidePanel4 As DevExpress.XtraEditors.SidePanel

        Private sidePanel5 As DevExpress.XtraEditors.SidePanel

        Private sidePanel6 As DevExpress.XtraEditors.SidePanel

        Private xtraTabControl1 As DevExpress.XtraTab.XtraTabControl

        Private xtraTabPage1 As DevExpress.XtraTab.XtraTabPage

        Private xtraTabPage2 As DevExpress.XtraTab.XtraTabPage

        Private richEditControlCS As DevExpress.XtraRichEdit.RichEditControl

        Private richEditControlVB As DevExpress.XtraRichEdit.RichEditControl
    End Class
End Namespace
