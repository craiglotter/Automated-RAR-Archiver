Imports System.Diagnostics
Imports System.IO

Module Startup_Module

    Private Sub Error_Handler(ByVal ex As String)
        Try
            Dim Display_Message1 As New Display_Message(ex.ToString)
            Display_Message1.ShowDialog()
            Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
            If dir.Exists = False Then
                dir.Create()
            End If
            Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
            filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & ex.ToString)
            filewriter.Flush()
            filewriter.Close()
        Catch except As Exception
            MsgBox("An error occurred in Automated RAR Archiver's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Sub main()
        Dim ApplicationName As String
        ApplicationName = "Automated RAR Archiver"
        Try
            Dim aModuleName As String = Diagnostics.Process.GetCurrentProcess.MainModule.ModuleName
            Dim aProcName As String = System.IO.Path.GetFileNameWithoutExtension(aModuleName)

            If Process.GetProcessesByName(aProcName).Length > 1 Then
                Dim Display_Message1 As New Display_Message("Another Instance of " & ApplicationName & " is already running. Only one instance of " & ApplicationName & " is allowed to run at any time")
                Display_Message1.ShowDialog()
                Application.Exit()
            Else
                Dim Splash As New Splash_Screen
                Splash.Show()
                While Splash.dataloaded = False
                End While
                Dim monitor As New Main_Screen(Splash)
                monitor.ShowDialog()
                monitor.Visible = False
                While monitor.dataloaded = False
                End While
                'monitor.Visible = True
                Application.Exit()

            End If
        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while launching " & ApplicationName & ". The program will attempt to recover shortly.")
        End Try
    End Sub
End Module
