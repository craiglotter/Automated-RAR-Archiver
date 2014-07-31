Imports System.IO
Imports Microsoft.Win32

Public Class Worker

    Inherits System.ComponentModel.Component

    ' Declares the variables you will use to hold your thread objects.

    Public WorkerThread As System.Threading.Thread

    Public foldername As String
    Public showwindow As Boolean

    Public result As String = ""

    Public Event WorkerErrorEncountered(ByVal ex As Exception)
    Public Event WorkerComplete(ByVal Result As String)
    Public Event WorkerFileCount()
    Public Event WorkerFolderCount()
    Public Event WorkerMessageExtracted(ByVal str As String)


#Region " Component Designer generated code "

    Public Sub New(ByVal Container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        Container.Add(Me)
    End Sub

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
      
        foldername = ""
        showwindow = False
    End Sub

    'Component overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region

    Private Sub Error_Handler(ByVal exc As Exception)
        Try
            RaiseEvent WorkerErrorEncountered(exc)
        Catch ex As Exception
            MsgBox("An error occurred in Automated RAR Archiver's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub


    Public Sub ChooseThreads(ByVal threadNumber As Integer)
        Try
            ' Determines which thread to start based on the value it receives.
            Select Case threadNumber
                Case 1
                    ' Sets the thread using the AddressOf the subroutine where
                    ' the thread will start.
                    WorkerThread = New System.Threading.Thread(AddressOf WorkerExecute)
                    ' Starts the thread.
                    WorkerThread.Start()
            End Select
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub



    Public Sub WorkerExecute()
        Try
          



            FolderCountRunner(foldername)
            FileCountRunner(foldername)
            Dim apptorun As String
            Dim finfo As FileInfo = New FileInfo((Application.StartupPath & "\Rar.exe").Replace("\\", "\"))
            If finfo.Exists = True Then

                Dim prevstamp, stamp As String
                stamp = Format(Now(), "yyyyMMdd")
                Dim fi As FileInfo
                For drun As Integer = 1 To 60
                    prevstamp = Format(Now().AddDays((drun * -1)), "yyyyMMdd")
                    fi = New FileInfo(foldername & "_" & prevstamp & ".rar")
                    If fi.Exists = True Then
                        fi.Delete()
                    End If
                Next


                fi = New FileInfo(foldername & "_" & stamp & ".rar")
                If fi.Exists = True Then
                    fi.Delete()
                End If
                fi = Nothing
                apptorun = """" & (Application.StartupPath & "\Rar.exe").Replace("\\", "\") & """ a """ & foldername & "_" & stamp & ".rar""" & " """ & foldername & """"

                DosShellCommand(apptorun, showwindow)
            End If
            finfo = Nothing



            result = "success"
        Catch ex As Exception
            Error_Handler(ex)
            result = "fail"
        Finally
            RaiseEvent WorkerComplete(result)
        End Try
    End Sub

  



    Private Function DosShellCommand(ByVal AppToRun As String, Optional ByVal silent As Boolean = True) As String
        Dim s As String = ""
        Try
            Dim myProcess As Process = New Process

            myProcess.StartInfo.FileName = "cmd.exe"
            myProcess.StartInfo.UseShellExecute = False

            Dim sErr As StreamReader
            Dim sOut As StreamReader
            Dim sIn As StreamWriter

            If silent = True Then
                myProcess.StartInfo.CreateNoWindow = True

                myProcess.StartInfo.RedirectStandardInput = True
                myProcess.StartInfo.RedirectStandardOutput = True
                myProcess.StartInfo.RedirectStandardError = True

                myProcess.StartInfo.FileName = AppToRun

                myProcess.Start()
                sIn = myProcess.StandardInput
                sIn.AutoFlush = True

                sOut = myProcess.StandardOutput()
                sErr = myProcess.StandardError

                sIn.Write(AppToRun & _
       System.Environment.NewLine)
                sIn.Write("exit" & System.Environment.NewLine)
                s = sOut.ReadToEnd()

                If Not myProcess.HasExited Then
                    myProcess.Kill()
                End If



                sIn.Close()
                sOut.Close()
                sErr.Close()
                myProcess.Close()
            Else
                myProcess.StartInfo.CreateNoWindow = True

                myProcess.StartInfo.RedirectStandardInput = True
                myProcess.StartInfo.RedirectStandardOutput = True
                myProcess.StartInfo.RedirectStandardError = True

                myProcess.StartInfo.FileName = AppToRun

                myProcess.Start()
                sIn = myProcess.StandardInput
                sIn.AutoFlush = True

                sOut = myProcess.StandardOutput()
                sErr = myProcess.StandardError

                sIn.Write(AppToRun & _
       System.Environment.NewLine)
                sIn.Write("cls" & System.Environment.NewLine)
                sIn.Write("exit" & System.Environment.NewLine)
                s = sOut.ReadLine
                Dim dirty As Boolean = True
                ' While Not s = "cls"
                While sOut.Peek > -1
                    If dirty = True Then
                        RaiseEvent WorkerMessageExtracted(s & vbCrLf)
                    End If
                    If sOut.Peek > -1 Then
                        s = sOut.ReadLine
                        dirty = True
                    Else
                        dirty = False
                    End If
                End While
                s = sOut.ReadToEnd()
                RaiseEvent WorkerMessageExtracted(s)
                If Not myProcess.HasExited Then
                    myProcess.Kill()
                End If



                sIn.Close()
                sOut.Close()
                sErr.Close()
                myProcess.Close()
            End If

        Catch ex As Exception
            Error_Handler(ex)
        End Try
        Return s
    End Function

    Private Sub FolderCountRunner(ByVal targetDirectory As String)
        Try
            Dim subdirectoryEntries As String() = Directory.GetDirectories(targetDirectory)
            Dim subdirectory As String
            For Each subdirectory In subdirectoryEntries
                RaiseEvent WorkerFolderCount()
                FolderCountRunner(subdirectory)
            Next subdirectory
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub FileCountRunner(ByVal targetDirectory As String)
        Try
            Dim fileEntries As String() = Directory.GetFiles(targetDirectory)
            Dim fileName As String
            For Each fileName In fileEntries
                RaiseEvent WorkerFileCount()
            Next fileName

            Dim subdirectoryEntries As String() = Directory.GetDirectories(targetDirectory)
            Dim subdirectory As String
            For Each subdirectory In subdirectoryEntries
                FileCountRunner(subdirectory)
            Next subdirectory
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub 'ProcessDirectory
End Class
