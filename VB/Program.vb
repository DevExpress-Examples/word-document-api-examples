Imports DevExpress.XtraEditors
Imports System
Imports System.Windows.Forms

Namespace RichEditDocumentServerAPIExample

    Friend Module Program

        ''' <summary>
        ''' The main entry point for the application.
        ''' </summary>
        <STAThread>
        Sub Main()
            Call Application.EnableVisualStyles()
            Application.SetCompatibleTextRenderingDefault(False)
            Call WindowsFormsSettings.SetPerMonitorDpiAware()
            Call Application.Run(New Form1())
        End Sub
    End Module
End Namespace
