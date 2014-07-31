Imports System.IO

Public Class Display_Message
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    Public Sub New(ByVal display As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        label1.Text = display
        label1.Select(label1.Text.Length - 1, 0)
    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Protected Friend WithEvents label1 As System.Windows.Forms.TextBox
    Protected Friend WithEvents Timer1 As System.Windows.Forms.Timer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Display_Message))
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Label2 = New System.Windows.Forms.Label
        Me.label1 = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'Timer1
        '
        Me.Timer1.Interval = 1000
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label2.Location = New System.Drawing.Point(6, 136)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(303, 24)
        Me.Label2.TabIndex = 1
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'label1
        '
        Me.label1.BackColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(192, Byte), CType(0, Byte))
        Me.label1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.label1.ForeColor = System.Drawing.Color.Black
        Me.label1.Location = New System.Drawing.Point(6, 8)
        Me.label1.Multiline = True
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(303, 128)
        Me.label1.TabIndex = 2
        Me.label1.Text = ""
        '
        'Display_Message
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(192, Byte), CType(0, Byte))
        Me.ClientSize = New System.Drawing.Size(314, 160)
        Me.Controls.Add(Me.label1)
        Me.Controls.Add(Me.Label2)
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(320, 192)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(320, 192)
        Me.Name = "Display_Message"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Automated RAR Archiver"
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private elapsed As Integer

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

    Private Sub Display_Message_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            elapsed = 0
            Label2.Text = "Closing in " & (10 - elapsed).ToString & " secs"
            Timer1.Start()
        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while displaying the required error message. The program will attempt to recover shortly.")
        End Try
    End Sub

    Private Sub TickHandler(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try
            elapsed = elapsed + 1
            Label2.Text = "Closing in " & (10 - elapsed).ToString & " secs"
            If elapsed = 10 Then
                Timer1.Stop()
                Me.Close()
            End If
        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while displaying the required error message. The program will attempt to recover shortly.")
        End Try
    End Sub
End Class
