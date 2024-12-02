Imports DevExpress.XtraEditors
Imports System.Linq
Imports System.Threading.Tasks

Namespace RichEditDocumentServerAPIExample

    Friend Module Program

        ''' <summary>
        ''' The main entry point for the application.
        ''' </summary>
        <STAThread>
        Sub Main()
            Application.EnableVisualStyles()
            Application.SetCompatibleTextRenderingDefault(False)
            Call WindowsFormsSettings.SetPerMonitorDpiAware()
            Application.Run(New Form1())
        End Sub
    End Module
End Namespace
