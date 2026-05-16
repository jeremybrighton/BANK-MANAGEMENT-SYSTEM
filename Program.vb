' ============================================================
'  BANK MANAGEMENT SYSTEM - Program.vb
'  Application entry point
' ============================================================
Imports System.Windows.Forms

Module Program

    <STAThread>
    Sub Main()
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)

        ' Initialise data storage and seed default admin
        InitialiseStorage()

        ' Launch login screen
        Application.Run(New frmLogin())
    End Sub

End Module
