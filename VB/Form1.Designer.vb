Namespace RichEditDocumentServerAPIExample
    Partial Public Class Form1
        ''' <summary>
        ''' Required designer variable.
        ''' </summary>
        Private components As System.ComponentModel.IContainer = Nothing

        ''' <summary>
        ''' Clean up any resources being used.
        ''' </summary>
        ''' <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing AndAlso (components IsNot Nothing) Then
                components.Dispose()
            End If
            MyBase.Dispose(disposing)
        End Sub

        #Region "Windows Form Designer generated code"

        ''' <summary>
        ''' Required method for Designer support - do not modify
        ''' the contents of this method with the code editor.
        ''' </summary>
        Private Sub InitializeComponent()
            Dim resources As New System.ComponentModel.ComponentResourceManager(GetType(Form1))
            Me.splitContainerControl1 = New DevExpress.XtraEditors.SplitContainerControl()
            Me.verticalSplitContainerControl = New DevExpress.XtraEditors.SplitContainerControl()
            Me.codeExampleNameLbl = New DevExpress.XtraEditors.LabelControl()
            Me.xtraTabControl1 = New DevExpress.XtraTab.XtraTabControl()
            Me.xtraTabPage2 = New DevExpress.XtraTab.XtraTabPage()
            Me.richEditControlVB = New DevExpress.XtraRichEdit.RichEditControl()
            Me.xtraTabPage1 = New DevExpress.XtraTab.XtraTabPage()
            Me.richEditControlCS = New DevExpress.XtraRichEdit.RichEditControl()
            Me.xtraTabPage3 = New DevExpress.XtraTab.XtraTabPage()
            Me.richEditControlCSClass = New DevExpress.XtraRichEdit.RichEditControl()
            Me.xtraTabPage4 = New DevExpress.XtraTab.XtraTabPage()
            Me.richEditControlVBClass = New DevExpress.XtraRichEdit.RichEditControl()
            Me.btnRun = New DevExpress.XtraEditors.SimpleButton()
            Me.treeList1 = New DevExpress.XtraTreeList.TreeList()
            DirectCast(Me.splitContainerControl1, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.splitContainerControl1.SuspendLayout()
            DirectCast(Me.verticalSplitContainerControl, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.verticalSplitContainerControl.SuspendLayout()
            DirectCast(Me.xtraTabControl1, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.xtraTabControl1.SuspendLayout()
            Me.xtraTabPage2.SuspendLayout()
            Me.xtraTabPage1.SuspendLayout()
            Me.xtraTabPage3.SuspendLayout()
            Me.xtraTabPage4.SuspendLayout()
            DirectCast(Me.treeList1, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.SuspendLayout()
            ' 
            ' splitContainerControl1
            ' 
            Me.splitContainerControl1.Anchor = (CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
            Me.splitContainerControl1.Location = New System.Drawing.Point(0, 0)
            Me.splitContainerControl1.Name = "splitContainerControl1"
            Me.splitContainerControl1.Panel1.Controls.Add(Me.verticalSplitContainerControl)
            Me.splitContainerControl1.Panel1.Text = "Panel1"
            Me.splitContainerControl1.Panel2.Controls.Add(Me.treeList1)
            Me.splitContainerControl1.Panel2.Text = "Panel2"
            Me.splitContainerControl1.Size = New System.Drawing.Size(829, 654)
            Me.splitContainerControl1.SplitterPosition = 569
            Me.splitContainerControl1.TabIndex = 0
            Me.splitContainerControl1.Text = "splitContainerControl1"
            ' 
            ' verticalSplitContainerControl
            ' 
            Me.verticalSplitContainerControl.Anchor = (CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
            Me.verticalSplitContainerControl.Horizontal = False
            Me.verticalSplitContainerControl.Location = New System.Drawing.Point(0, 0)
            Me.verticalSplitContainerControl.Name = "verticalSplitContainerControl"
            Me.verticalSplitContainerControl.Panel1.Controls.Add(Me.codeExampleNameLbl)
            Me.verticalSplitContainerControl.Panel1.Controls.Add(Me.xtraTabControl1)
            Me.verticalSplitContainerControl.Panel1.Text = "Panel1"
            Me.verticalSplitContainerControl.Panel2.Controls.Add(Me.btnRun)
            Me.verticalSplitContainerControl.Panel2.Text = "Panel2"
            Me.verticalSplitContainerControl.Size = New System.Drawing.Size(569, 654)
            Me.verticalSplitContainerControl.SplitterPosition = 570
            Me.verticalSplitContainerControl.TabIndex = 0
            Me.verticalSplitContainerControl.Text = "splitContainerControl2"
            ' 
            ' codeExampleNameLbl
            ' 
            Me.codeExampleNameLbl.Anchor = (CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
            Me.codeExampleNameLbl.Appearance.Font = New System.Drawing.Font("Tahoma", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, (CByte(0)))
            Me.codeExampleNameLbl.Appearance.Options.UseFont = True
            Me.codeExampleNameLbl.Location = New System.Drawing.Point(12, 29)
            Me.codeExampleNameLbl.LookAndFeel.SkinName = "Office 2016 Colorful"
            Me.codeExampleNameLbl.Name = "codeExampleNameLbl"
            Me.codeExampleNameLbl.Size = New System.Drawing.Size(291, 39)
            Me.codeExampleNameLbl.TabIndex = 4
            Me.codeExampleNameLbl.Text = "Examples Not Found"
            ' 
            ' xtraTabControl1
            ' 
            Me.xtraTabControl1.Anchor = (CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
            Me.xtraTabControl1.HeaderAutoFill = DevExpress.Utils.DefaultBoolean.True
            Me.xtraTabControl1.Location = New System.Drawing.Point(3, 86)
            Me.xtraTabControl1.Name = "xtraTabControl1"
            Me.xtraTabControl1.SelectedTabPage = Me.xtraTabPage2
            Me.xtraTabControl1.Size = New System.Drawing.Size(569, 481)
            Me.xtraTabControl1.TabIndex = 0
            Me.xtraTabControl1.TabPages.AddRange(New DevExpress.XtraTab.XtraTabPage() { Me.xtraTabPage1, Me.xtraTabPage2, Me.xtraTabPage3, Me.xtraTabPage4})
            ' 
            ' xtraTabPage2
            ' 
            Me.xtraTabPage2.Controls.Add(Me.richEditControlVB)
            Me.xtraTabPage2.Name = "xtraTabPage2"
            Me.xtraTabPage2.Size = New System.Drawing.Size(563, 453)
            Me.xtraTabPage2.Tag = "VB"
            Me.xtraTabPage2.Text = "VB"
            ' 
            ' richEditControlVB
            ' 
            Me.richEditControlVB.ActiveViewType = DevExpress.XtraRichEdit.RichEditViewType.Simple
            Me.richEditControlVB.Anchor = (CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
            Me.richEditControlVB.Location = New System.Drawing.Point(0, 0)
            Me.richEditControlVB.Name = "richEditControlVB"
            Me.richEditControlVB.Size = New System.Drawing.Size(563, 453)
            Me.richEditControlVB.TabIndex = 0
            ' 
            ' xtraTabPage1
            ' 
            Me.xtraTabPage1.Controls.Add(Me.richEditControlCS)
            Me.xtraTabPage1.Name = "xtraTabPage1"
            Me.xtraTabPage1.Size = New System.Drawing.Size(563, 453)
            Me.xtraTabPage1.Tag = "CS"
            Me.xtraTabPage1.Text = "C#"
            ' 
            ' richEditControlCS
            ' 
            Me.richEditControlCS.ActiveViewType = DevExpress.XtraRichEdit.RichEditViewType.Simple
            Me.richEditControlCS.Dock = System.Windows.Forms.DockStyle.Fill
            Me.richEditControlCS.Location = New System.Drawing.Point(0, 0)
            Me.richEditControlCS.Name = "richEditControlCS"
            Me.richEditControlCS.Size = New System.Drawing.Size(563, 453)
            Me.richEditControlCS.TabIndex = 0
            ' 
            ' xtraTabPage3
            ' 
            Me.xtraTabPage3.Controls.Add(Me.richEditControlCSClass)
            Me.xtraTabPage3.Name = "xtraTabPage3"
            Me.xtraTabPage3.Size = New System.Drawing.Size(563, 453)
            Me.xtraTabPage3.Tag = "CS"
            Me.xtraTabPage3.Text = "C# Helper"
            ' 
            ' richEditControlCSClass
            ' 
            Me.richEditControlCSClass.ActiveViewType = DevExpress.XtraRichEdit.RichEditViewType.Simple
            Me.richEditControlCSClass.Dock = System.Windows.Forms.DockStyle.Fill
            Me.richEditControlCSClass.Location = New System.Drawing.Point(0, 0)
            Me.richEditControlCSClass.Name = "richEditControlCSClass"
            Me.richEditControlCSClass.Size = New System.Drawing.Size(563, 453)
            Me.richEditControlCSClass.TabIndex = 0
            ' 
            ' xtraTabPage4
            ' 
            Me.xtraTabPage4.Controls.Add(Me.richEditControlVBClass)
            Me.xtraTabPage4.Name = "xtraTabPage4"
            Me.xtraTabPage4.Size = New System.Drawing.Size(563, 453)
            Me.xtraTabPage4.Tag = "VB"
            Me.xtraTabPage4.Text = "VB Helper"
            ' 
            ' richEditControlVBClass
            ' 
            Me.richEditControlVBClass.ActiveViewType = DevExpress.XtraRichEdit.RichEditViewType.Simple
            Me.richEditControlVBClass.Dock = System.Windows.Forms.DockStyle.Fill
            Me.richEditControlVBClass.Location = New System.Drawing.Point(0, 0)
            Me.richEditControlVBClass.Name = "richEditControlVBClass"
            Me.richEditControlVBClass.Size = New System.Drawing.Size(563, 453)
            Me.richEditControlVBClass.TabIndex = 0
            ' 
            ' btnRun
            ' 
            Me.btnRun.Image = (DirectCast(resources.GetObject("btnRun.Image"), System.Drawing.Image))
            Me.btnRun.Location = New System.Drawing.Point(12, 16)
            Me.btnRun.Name = "btnRun"
            Me.btnRun.Size = New System.Drawing.Size(177, 31)
            Me.btnRun.TabIndex = 0
            Me.btnRun.Text = "Open in Microsoft Word"
            ' 
            ' treeList1
            ' 
            Me.treeList1.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
            Me.treeList1.Location = New System.Drawing.Point(0, 0)
            Me.treeList1.Name = "treeList1"
            Me.treeList1.Size = New System.Drawing.Size(255, 654)
            Me.treeList1.TabIndex = 0
            ' 
            ' Form1
            ' 
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(829, 654)
            Me.Controls.Add(Me.splitContainerControl1)
            Me.Name = "Form1"
            Me.Text = "Form1"
            DirectCast(Me.splitContainerControl1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.splitContainerControl1.ResumeLayout(False)
            DirectCast(Me.verticalSplitContainerControl, System.ComponentModel.ISupportInitialize).EndInit()
            Me.verticalSplitContainerControl.ResumeLayout(False)
            DirectCast(Me.xtraTabControl1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.xtraTabControl1.ResumeLayout(False)
            Me.xtraTabPage2.ResumeLayout(False)
            Me.xtraTabPage1.ResumeLayout(False)
            Me.xtraTabPage3.ResumeLayout(False)
            Me.xtraTabPage4.ResumeLayout(False)
            DirectCast(Me.treeList1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.ResumeLayout(False)

        End Sub

        #End Region

        Private splitContainerControl1 As DevExpress.XtraEditors.SplitContainerControl
        Private verticalSplitContainerControl As DevExpress.XtraEditors.SplitContainerControl
        Private WithEvents btnRun As DevExpress.XtraEditors.SimpleButton
        Private codeExampleNameLbl As DevExpress.XtraEditors.LabelControl
        Private xtraTabControl1 As DevExpress.XtraTab.XtraTabControl
        Private xtraTabPage1 As DevExpress.XtraTab.XtraTabPage
        Private richEditControlCS As DevExpress.XtraRichEdit.RichEditControl
        Private xtraTabPage2 As DevExpress.XtraTab.XtraTabPage
        Private richEditControlVB As DevExpress.XtraRichEdit.RichEditControl
        Private WithEvents treeList1 As DevExpress.XtraTreeList.TreeList
        Private xtraTabPage3 As DevExpress.XtraTab.XtraTabPage
        Private xtraTabPage4 As DevExpress.XtraTab.XtraTabPage
        Private richEditControlCSClass As DevExpress.XtraRichEdit.RichEditControl
        Private richEditControlVBClass As DevExpress.XtraRichEdit.RichEditControl
    End Class
End Namespace

