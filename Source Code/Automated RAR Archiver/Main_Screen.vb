Imports Microsoft.Win32
Imports System.IO

Public Class Main_Screen
    Inherits System.Windows.Forms.Form

    Dim WithEvents Worker1 As Worker
    Public Delegate Sub WorkerhHandler(ByVal Result As String)
    Public Delegate Sub WorkerProgresshHandler()
    Public Delegate Sub WorkerFileCreatedhHandler()
    Public Delegate Sub WorkerErrorEncountered(ByVal ex As Exception)
    Public Delegate Sub WorkerMessageExtractedhHandler(ByVal str As String)


    Private splash_loader As Splash_Screen
    Public dataloaded As Boolean = False

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Worker1 = New Worker
        AddHandler Worker1.WorkerComplete, AddressOf WorkerHandler
        AddHandler Worker1.WorkerFolderCount, AddressOf WorkerFolderCountHandler
        AddHandler Worker1.WorkerFileCount, AddressOf WorkerFileCountHandler
        AddHandler Worker1.WorkerMessageExtracted, AddressOf WorkerMessageExtractedHandler
        AddHandler Worker1.WorkerErrorEncountered, AddressOf WorkerErrorEncounteredHandler
    End Sub

    Public Sub New(ByVal splash As Splash_Screen)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        splash_loader = splash
        Worker1 = New Worker
        AddHandler Worker1.WorkerComplete, AddressOf WorkerHandler
        AddHandler Worker1.WorkerFolderCount, AddressOf WorkerFolderCountHandler
        AddHandler Worker1.WorkerFileCount, AddressOf WorkerFileCountHandler
        AddHandler Worker1.WorkerMessageExtracted, AddressOf WorkerMessageExtractedHandler
        AddHandler Worker1.WorkerErrorEncountered, AddressOf WorkerErrorEncounteredHandler
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
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents Timer2 As System.Windows.Forms.Timer
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents NotifyIcon1 As System.Windows.Forms.NotifyIcon
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents PictureBox5 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox4 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents folderscreated As System.Windows.Forms.Label
    Friend WithEvents filescreated As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents basefolder As System.Windows.Forms.TextBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents currenttime As System.Windows.Forms.Label
    Friend WithEvents scheduledtime As System.Windows.Forms.Label
    Friend WithEvents shour As System.Windows.Forms.NumericUpDown
    Friend WithEvents sminute As System.Windows.Forms.NumericUpDown
    Friend WithEvents ssecond As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents processstart As System.Windows.Forms.Label
    Friend WithEvents processend As System.Windows.Forms.Label
    Friend WithEvents showwindow As System.Windows.Forms.CheckBox
    Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Main_Screen))
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        Me.Label8 = New System.Windows.Forms.Label
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem4 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.shour = New System.Windows.Forms.NumericUpDown
        Me.sminute = New System.Windows.Forms.NumericUpDown
        Me.ssecond = New System.Windows.Forms.NumericUpDown
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
        Me.Label9 = New System.Windows.Forms.Label
        Me.Button3 = New System.Windows.Forms.Button
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.PictureBox5 = New System.Windows.Forms.PictureBox
        Me.PictureBox4 = New System.Windows.Forms.PictureBox
        Me.PictureBox3 = New System.Windows.Forms.PictureBox
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.folderscreated = New System.Windows.Forms.Label
        Me.filescreated = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.basefolder = New System.Windows.Forms.TextBox
        Me.Button2 = New System.Windows.Forms.Button
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.showwindow = New System.Windows.Forms.CheckBox
        Me.currenttime = New System.Windows.Forms.Label
        Me.scheduledtime = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.processstart = New System.Windows.Forms.Label
        Me.processend = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox
        CType(Me.shour, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.sminute, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ssecond, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 1000
        '
        'ImageList1
        '
        Me.ImageList1.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        '
        'Timer2
        '
        Me.Timer2.Interval = 1000
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.Transparent
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.DimGray
        Me.Label8.Location = New System.Drawing.Point(208, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(152, 16)
        Me.Label8.TabIndex = 33
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.ToolTip1.SetToolTip(Me.Label8, "Current System Time")
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.ContextMenu = Me.ContextMenu1
        Me.NotifyIcon1.Icon = CType(resources.GetObject("NotifyIcon1.Icon"), System.Drawing.Icon)
        Me.NotifyIcon1.Text = "Resting..."
        Me.NotifyIcon1.Visible = True
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem4, Me.MenuItem2, Me.MenuItem3})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.Text = "Display Main Screen"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 1
        Me.MenuItem4.Text = "Hide Main Screen"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 2
        Me.MenuItem2.Text = "-"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 3
        Me.MenuItem3.Text = "Exit Application"
        '
        'shour
        '
        Me.shour.BackColor = System.Drawing.Color.YellowGreen
        Me.shour.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.shour.ForeColor = System.Drawing.Color.DimGray
        Me.shour.Location = New System.Drawing.Point(24, 80)
        Me.shour.Maximum = New Decimal(New Integer() {24, 0, 0, 0})
        Me.shour.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.shour.Name = "shour"
        Me.shour.Size = New System.Drawing.Size(40, 20)
        Me.shour.TabIndex = 94
        Me.ToolTip1.SetToolTip(Me.shour, "Hours")
        Me.shour.Value = New Decimal(New Integer() {24, 0, 0, 0})
        '
        'sminute
        '
        Me.sminute.BackColor = System.Drawing.Color.YellowGreen
        Me.sminute.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.sminute.ForeColor = System.Drawing.Color.DimGray
        Me.sminute.Location = New System.Drawing.Point(72, 80)
        Me.sminute.Maximum = New Decimal(New Integer() {59, 0, 0, 0})
        Me.sminute.Name = "sminute"
        Me.sminute.Size = New System.Drawing.Size(40, 20)
        Me.sminute.TabIndex = 99
        Me.ToolTip1.SetToolTip(Me.sminute, "Minutes")
        '
        'ssecond
        '
        Me.ssecond.BackColor = System.Drawing.Color.YellowGreen
        Me.ssecond.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ssecond.ForeColor = System.Drawing.Color.DimGray
        Me.ssecond.Location = New System.Drawing.Point(120, 80)
        Me.ssecond.Maximum = New Decimal(New Integer() {59, 0, 0, 0})
        Me.ssecond.Name = "ssecond"
        Me.ssecond.Size = New System.Drawing.Size(40, 20)
        Me.ssecond.TabIndex = 100
        Me.ToolTip1.SetToolTip(Me.ssecond, "Seconds")
        '
        'FolderBrowserDialog1
        '
        Me.FolderBrowserDialog1.Description = "Select the base folder to archive"
        Me.FolderBrowserDialog1.ShowNewFolderButton = False
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.Transparent
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.DimGray
        Me.Label9.Location = New System.Drawing.Point(336, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(109, 16)
        Me.Label9.TabIndex = 54
        Me.Label9.Text = "BUILD 20060208.3"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(192, Byte), CType(0, Byte))
        Me.Button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button3.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.ForeColor = System.Drawing.Color.Black
        Me.Button3.Location = New System.Drawing.Point(328, 160)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(104, 16)
        Me.Button3.TabIndex = 70
        Me.Button3.Text = "FORCE ARCHIVAL"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.White
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Gray
        Me.Label7.Location = New System.Drawing.Point(24, 153)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(112, 26)
        Me.Label7.TabIndex = 69
        Me.Label7.Text = "Resting..."
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.White
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.LimeGreen
        Me.Label1.Location = New System.Drawing.Point(232, 160)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(66, 16)
        Me.Label1.TabIndex = 67
        Me.Label1.Text = "Waiting"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'PictureBox5
        '
        Me.PictureBox5.BackgroundImage = CType(resources.GetObject("PictureBox5.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox5.Location = New System.Drawing.Point(208, 160)
        Me.PictureBox5.Name = "PictureBox5"
        Me.PictureBox5.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox5.TabIndex = 66
        Me.PictureBox5.TabStop = False
        '
        'PictureBox4
        '
        Me.PictureBox4.BackgroundImage = CType(resources.GetObject("PictureBox4.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox4.Location = New System.Drawing.Point(192, 160)
        Me.PictureBox4.Name = "PictureBox4"
        Me.PictureBox4.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox4.TabIndex = 65
        Me.PictureBox4.TabStop = False
        '
        'PictureBox3
        '
        Me.PictureBox3.BackgroundImage = CType(resources.GetObject("PictureBox3.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox3.Location = New System.Drawing.Point(176, 160)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox3.TabIndex = 64
        Me.PictureBox3.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.BackgroundImage = CType(resources.GetObject("PictureBox2.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(160, 160)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox2.TabIndex = 63
        Me.PictureBox2.TabStop = False
        '
        'PictureBox1
        '
        Me.PictureBox1.BackgroundImage = CType(resources.GetObject("PictureBox1.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(144, 160)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox1.TabIndex = 62
        Me.PictureBox1.TabStop = False
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.White
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label2.Location = New System.Drawing.Point(16, 152)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(288, 32)
        Me.Label2.TabIndex = 68
        '
        'folderscreated
        '
        Me.folderscreated.BackColor = System.Drawing.Color.Transparent
        Me.folderscreated.ForeColor = System.Drawing.Color.Black
        Me.folderscreated.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.folderscreated.Location = New System.Drawing.Point(264, 24)
        Me.folderscreated.Name = "folderscreated"
        Me.folderscreated.Size = New System.Drawing.Size(80, 16)
        Me.folderscreated.TabIndex = 73
        Me.folderscreated.Text = "0"
        Me.folderscreated.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'filescreated
        '
        Me.filescreated.BackColor = System.Drawing.Color.Transparent
        Me.filescreated.ForeColor = System.Drawing.Color.Black
        Me.filescreated.Location = New System.Drawing.Point(264, 40)
        Me.filescreated.Name = "filescreated"
        Me.filescreated.Size = New System.Drawing.Size(80, 16)
        Me.filescreated.TabIndex = 75
        Me.filescreated.Text = "0"
        Me.filescreated.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.ForeColor = System.Drawing.Color.Black
        Me.Label5.Location = New System.Drawing.Point(16, 104)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(144, 16)
        Me.Label5.TabIndex = 88
        Me.Label5.Text = "Base Folder to Archive:"
        '
        'basefolder
        '
        Me.basefolder.BackColor = System.Drawing.Color.YellowGreen
        Me.basefolder.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.basefolder.ForeColor = System.Drawing.Color.DimGray
        Me.basefolder.Location = New System.Drawing.Point(24, 120)
        Me.basefolder.Name = "basefolder"
        Me.basefolder.ReadOnly = True
        Me.basefolder.Size = New System.Drawing.Size(344, 20)
        Me.basefolder.TabIndex = 87
        Me.basefolder.Text = ""
        '
        'Button2
        '
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.ForeColor = System.Drawing.Color.Black
        Me.Button2.Location = New System.Drawing.Point(376, 120)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(56, 20)
        Me.Button2.TabIndex = 89
        Me.Button2.Text = "BROWSE"
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.Color.Transparent
        Me.Label11.ForeColor = System.Drawing.Color.Black
        Me.Label11.Location = New System.Drawing.Point(16, 64)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(176, 16)
        Me.Label11.TabIndex = 95
        Me.Label11.Text = "Scheduled Archival: (HH:MM:SS)"
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.Color.Transparent
        Me.Label12.ForeColor = System.Drawing.Color.Black
        Me.Label12.Location = New System.Drawing.Point(352, 24)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(88, 16)
        Me.Label12.TabIndex = 96
        Me.Label12.Text = "Folders Effected"
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.Color.Transparent
        Me.Label13.ForeColor = System.Drawing.Color.Black
        Me.Label13.Location = New System.Drawing.Point(352, 40)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(80, 16)
        Me.Label13.TabIndex = 97
        Me.Label13.Text = "Files Effected"
        '
        'showwindow
        '
        Me.showwindow.BackColor = System.Drawing.Color.Transparent
        Me.showwindow.Checked = True
        Me.showwindow.CheckState = System.Windows.Forms.CheckState.Checked
        Me.showwindow.Enabled = False
        Me.showwindow.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.showwindow.ForeColor = System.Drawing.Color.Black
        Me.showwindow.Location = New System.Drawing.Point(176, 96)
        Me.showwindow.Name = "showwindow"
        Me.showwindow.Size = New System.Drawing.Size(128, 24)
        Me.showwindow.TabIndex = 101
        Me.showwindow.Text = "Silent Operation"
        Me.showwindow.Visible = False
        '
        'currenttime
        '
        Me.currenttime.BackColor = System.Drawing.Color.Transparent
        Me.currenttime.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.currenttime.ForeColor = System.Drawing.Color.DimGray
        Me.currenttime.Location = New System.Drawing.Point(136, 16)
        Me.currenttime.Name = "currenttime"
        Me.currenttime.Size = New System.Drawing.Size(112, 23)
        Me.currenttime.TabIndex = 102
        Me.currenttime.Text = "24:00:00"
        '
        'scheduledtime
        '
        Me.scheduledtime.BackColor = System.Drawing.Color.Transparent
        Me.scheduledtime.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.scheduledtime.ForeColor = System.Drawing.Color.Gainsboro
        Me.scheduledtime.Location = New System.Drawing.Point(16, 16)
        Me.scheduledtime.Name = "scheduledtime"
        Me.scheduledtime.Size = New System.Drawing.Size(112, 23)
        Me.scheduledtime.TabIndex = 103
        Me.scheduledtime.Text = "24:00:00"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Gainsboro
        Me.Label3.Location = New System.Drawing.Point(24, 40)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(88, 16)
        Me.Label3.TabIndex = 104
        Me.Label3.Text = "SCHEDULED TIME"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.DimGray
        Me.Label4.Location = New System.Drawing.Point(144, 40)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(88, 16)
        Me.Label4.TabIndex = 105
        Me.Label4.Text = "CURRENT TIME"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.Transparent
        Me.Label6.ForeColor = System.Drawing.Color.Black
        Me.Label6.Location = New System.Drawing.Point(352, 56)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(88, 16)
        Me.Label6.TabIndex = 107
        Me.Label6.Text = "Last Execute (S)"
        '
        'processstart
        '
        Me.processstart.BackColor = System.Drawing.Color.Transparent
        Me.processstart.ForeColor = System.Drawing.Color.Black
        Me.processstart.Location = New System.Drawing.Point(200, 56)
        Me.processstart.Name = "processstart"
        Me.processstart.Size = New System.Drawing.Size(144, 16)
        Me.processstart.TabIndex = 106
        Me.processstart.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'processend
        '
        Me.processend.BackColor = System.Drawing.Color.Transparent
        Me.processend.ForeColor = System.Drawing.Color.Black
        Me.processend.Location = New System.Drawing.Point(200, 72)
        Me.processend.Name = "processend"
        Me.processend.Size = New System.Drawing.Size(144, 16)
        Me.processend.TabIndex = 108
        Me.processend.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label14
        '
        Me.Label14.BackColor = System.Drawing.Color.Transparent
        Me.Label14.ForeColor = System.Drawing.Color.Black
        Me.Label14.Location = New System.Drawing.Point(352, 72)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(88, 16)
        Me.Label14.TabIndex = 109
        Me.Label14.Text = "Last Execute (E)"
        '
        'RichTextBox1
        '
        Me.RichTextBox1.AcceptsTab = True
        Me.RichTextBox1.BackColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(192, Byte), CType(0, Byte))
        Me.RichTextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.RichTextBox1.DetectUrls = False
        Me.RichTextBox1.ForeColor = System.Drawing.Color.DimGray
        Me.RichTextBox1.Location = New System.Drawing.Point(16, 232)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.ReadOnly = True
        Me.RichTextBox1.Size = New System.Drawing.Size(416, 64)
        Me.RichTextBox1.TabIndex = 110
        Me.RichTextBox1.Text = ""
        Me.RichTextBox1.WordWrap = False
        '
        'Main_Screen
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(192, Byte), CType(0, Byte))
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.ClientSize = New System.Drawing.Size(450, 216)
        Me.Controls.Add(Me.RichTextBox1)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.processend)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.processstart)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.scheduledtime)
        Me.Controls.Add(Me.currenttime)
        Me.Controls.Add(Me.showwindow)
        Me.Controls.Add(Me.ssecond)
        Me.Controls.Add(Me.sminute)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.shour)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.basefolder)
        Me.Controls.Add(Me.filescreated)
        Me.Controls.Add(Me.folderscreated)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.PictureBox5)
        Me.Controls.Add(Me.PictureBox4)
        Me.Controls.Add(Me.PictureBox3)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label9)
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(456, 240)
        Me.MinimumSize = New System.Drawing.Size(456, 240)
        Me.Name = "Main_Screen"
        Me.ShowInTaskbar = False
        Me.Text = "Automated RAR Archiver"
        CType(Me.shour, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.sminute, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ssecond, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private current_light As Integer = 0
    Private current_colour As Integer = 0
    Private currently_working As Boolean = False




    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then
                Dim Display_Message1 As New Display_Message("The Application encountered the following problem: " & vbCrLf & ex.ToString)
                Display_Message1.ShowDialog()
                Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
                If dir.Exists = False Then
                    dir.Create()
                End If
                Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & ex.ToString)
                filewriter.Flush()
                filewriter.Close()
            End If
        Catch exc As Exception

            MsgBox("An error occurred in Automated RAR Archiver's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Private Sub run_green_lights()
        Try
            Label1.ForeColor = Color.LimeGreen
            Label1.Text = "Waiting"
            Label7.Text = "Resting..."

            current_light = current_light - 1
            If current_light < 1 Then
                current_light = 5
            End If
            current_colour = 0
            PictureBox1.Image = ImageList1.Images(1)
            PictureBox2.Image = ImageList1.Images(1)
            PictureBox3.Image = ImageList1.Images(1)
            PictureBox4.Image = ImageList1.Images(1)
            PictureBox5.Image = ImageList1.Images(1)

            Select Case current_light
                Case 0

                    PictureBox1.Image = ImageList1.Images(0)
                Case 1

                    PictureBox2.Image = ImageList1.Images(0)
                Case 2

                    PictureBox3.Image = ImageList1.Images(0)
                Case 3

                    PictureBox4.Image = ImageList1.Images(0)
                Case 4

                    PictureBox5.Image = ImageList1.Images(0)
                Case 5

                    PictureBox1.Image = ImageList1.Images(0)
            End Select

            current_light = current_light + 1
            If current_light > 5 Then
                current_light = 1
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub run_orange_lights()
        Try
            Label1.ForeColor = Color.DarkOrange
            Label1.Text = "Working"

            current_light = current_light - 1
            If current_light < 1 Then
                current_light = 5
            End If
            current_colour = 1
            PictureBox1.Image = ImageList1.Images(3)
            PictureBox2.Image = ImageList1.Images(3)
            PictureBox3.Image = ImageList1.Images(3)
            PictureBox4.Image = ImageList1.Images(3)
            PictureBox5.Image = ImageList1.Images(3)
            Select Case current_light
                Case 0
                    PictureBox1.Image = ImageList1.Images(2)
                Case 1
                    PictureBox2.Image = ImageList1.Images(2)
                Case 2
                    PictureBox3.Image = ImageList1.Images(2)
                Case 3
                    PictureBox4.Image = ImageList1.Images(2)
                Case 4
                    PictureBox5.Image = ImageList1.Images(2)
                Case 5
                    PictureBox1.Image = ImageList1.Images(2)
            End Select

            current_light = current_light + 1
            If current_light > 5 Then
                current_light = 1
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub run_lights()
        Try
            If current_colour = 1 Then
                Select Case current_light
                    Case 0
                        PictureBox5.Image = ImageList1.Images(3)
                        PictureBox1.Image = ImageList1.Images(2)
                    Case 1
                        PictureBox1.Image = ImageList1.Images(3)
                        PictureBox2.Image = ImageList1.Images(2)
                    Case 2
                        PictureBox2.Image = ImageList1.Images(3)
                        PictureBox3.Image = ImageList1.Images(2)
                    Case 3
                        PictureBox3.Image = ImageList1.Images(3)
                        PictureBox4.Image = ImageList1.Images(2)
                    Case 4
                        PictureBox4.Image = ImageList1.Images(3)
                        PictureBox5.Image = ImageList1.Images(2)
                    Case 5
                        PictureBox5.Image = ImageList1.Images(3)
                        PictureBox1.Image = ImageList1.Images(2)
                End Select
            Else
                Select Case current_light
                    Case 0
                        PictureBox5.Image = ImageList1.Images(1)
                        PictureBox1.Image = ImageList1.Images(0)
                    Case 1
                        PictureBox1.Image = ImageList1.Images(1)
                        PictureBox2.Image = ImageList1.Images(0)
                    Case 2
                        PictureBox2.Image = ImageList1.Images(1)
                        PictureBox3.Image = ImageList1.Images(0)
                    Case 3
                        PictureBox3.Image = ImageList1.Images(1)
                        PictureBox4.Image = ImageList1.Images(0)
                    Case 4
                        PictureBox4.Image = ImageList1.Images(1)
                        PictureBox5.Image = ImageList1.Images(0)
                    Case 5
                        PictureBox5.Image = ImageList1.Images(1)
                        PictureBox1.Image = ImageList1.Images(0)
                End Select

            End If

            current_light = current_light + 1
            If current_light > 5 Then
                current_light = 1
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Try
            run_lights()
            Label8.Text = Format(Now(), "dd/MM/yyyy hh:mm:ss tt")
            currenttime.Text = Format(Now(), "HH:mm:ss")
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Main_Screen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try

            Load_Registry_Values()
            Label8.Text = Format(Now(), "dd/MM/yyyy hh:mm:ss tt")
            currenttime.Text = Format(Now(), "hh:mm:ss tt")
            Timer1.Start()
            Timer2.Start()
            Label5.Select()
            dataloaded = True
            splash_loader.Visible = False
            hide_application()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub exit_application()
        Try
            Save_Registry_Values()
            If Worker1.WorkerThread Is Nothing = False Then
                Worker1.WorkerThread.Abort()
                Worker1.Dispose()
            End If
            NotifyIcon1.Dispose()
            Application.Exit()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Main_Screen_closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        Try
            exit_application()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub




    Public Sub WorkerHandler(ByVal Result As String)
        Try
            currently_working = False
            processend.Text = Format(Now(), "dd/MM/yyyy hh:mm:ss tt")
            NotifyIcon1.Text = "Resting... "
            run_green_lights()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub WorkerFileCountHandler()
        Try
            filescreated.Text = CLng(filescreated.Text) + 1
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub WorkerMessageExtractedHandler(ByVal str As String)
        Try
            RichTextBox1.AppendText(str)

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub WorkerFolderCountHandler()
        Try
            folderscreated.Text = CLng(folderscreated.Text) + 1
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub
    Public Sub WorkerErrorEncounteredHandler(ByVal ex As Exception)
        Try
            Error_Handler(ex, "Worker Error Encountered")
        Catch exc As Exception
            Error_Handler(exc)
        End Try
    End Sub

    Private Sub run_worker(Optional ByVal threadselect As Integer = 1)
        run_orange_lights()
        folderscreated.Text = 0
        filescreated.Text = 0
        processend.Text = ""
        processstart.Text = Format(Now(), "dd/MM/yyyy hh:mm:ss tt")

        Worker1.foldername = basefolder.Text
        Worker1.showwindow = showwindow.Checked


        Select Case threadselect
            Case 1
                Worker1.ChooseThreads(1)
            Case 2
                Worker1.ChooseThreads(2)
        End Select

        currently_working = True
    End Sub



    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click
        Try
            Me.Close()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub show_application()
        Try
            Me.Opacity = 1

            Me.BringToFront()
            Me.Refresh()
            Me.WindowState = FormWindowState.Normal

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub NotifyIcon1_dblclick(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon1.DoubleClick
        show_application()
    End Sub
    Private Sub NotifyIcon1_snglclick(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon1.Click
        show_application()
    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click
        show_application()
    End Sub

    Private Sub Main_Screen_resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        Try

            If Me.WindowState = FormWindowState.Minimized Then
                NotifyIcon1.Visible = True
                Me.Opacity = 0
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub force_check(Optional ByVal threadselect As Integer = 1)
        Try
            Label7.Text = "Busy Working..."
            NotifyIcon1.Text = "Archiving Base Folder..."
            If currently_working = False Then
                run_worker(threadselect)
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        force_check()
    End Sub




    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            Dim tinfo As DirectoryInfo = New DirectoryInfo(basefolder.Text)
            If tinfo.Exists = True Then
                FolderBrowserDialog1.SelectedPath = basefolder.Text
            End If
            Dim result As DialogResult = FolderBrowserDialog1.ShowDialog
            If result = DialogResult.OK Then
                basefolder.Text = FolderBrowserDialog1.SelectedPath
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub Load_Registry_Values()
        Try
            Dim configflag As Boolean
            configflag = False
            Dim str As String
            Dim keyflag1 As Boolean = False
            Dim oReg As RegistryKey = Registry.LocalMachine
            Dim keys() As String = oReg.GetSubKeyNames()
            System.Array.Sort(keys)

            For Each str In keys
                If str.Equals("Software\Automated RAR Archiver") = True Then
                    keyflag1 = True
                    Exit For
                End If
            Next str

            If keyflag1 = False Then
                oReg.CreateSubKey("Software\Automated RAR Archiver")
            End If

            keyflag1 = False

            Dim oKey As RegistryKey = oReg.OpenSubKey("Software\Automated RAR Archiver", True)


            str = oKey.GetValue("basefolder")
            If Not IsNothing(str) And Not (str = "") Then
                basefolder.Text = str
            Else
                oKey.SetValue("basefolder", Application.StartupPath)
                basefolder.Text = Application.StartupPath
            End If

            str = oKey.GetValue("shour")
            If Not IsNothing(str) And Not (str = "") Then
                shour.Value = CInt(str)
            Else
                oKey.SetValue("shour", 23)
                shour.Value = 23
            End If


            str = oKey.GetValue("sminute")
            If Not IsNothing(str) And Not (str = "") Then
                sminute.Value = CInt(str)
            Else
                oKey.SetValue("sminute", 0)
                sminute.Value = 0
            End If

            str = oKey.GetValue("ssecond")
            If Not IsNothing(str) And Not (str = "") Then
                ssecond.Value = CInt(str)
            Else
                oKey.SetValue("ssecond", 0)
                ssecond.Value = 0
            End If

            oKey.Close()
            oReg.Close()

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Save_Registry_Values()
        Try
            Dim oReg As RegistryKey = Registry.LocalMachine
            Dim oKey As RegistryKey = oReg.OpenSubKey("Software\Automated RAR Archiver", True)



            oKey.SetValue("ssecond", ssecond.Value)
            oKey.SetValue("sminute", sminute.Value)
            oKey.SetValue("shour", shour.Value)
            oKey.SetValue("basefolder", basefolder.Text)

            oKey.Close()
            oReg.Close()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub


    Private Sub update_scheduled_time()
        Try

            Dim padhour, padminute, padsecond As String
            If shour.Value.ToString.Length < 2 Then
                padhour = "0" & shour.Value.ToString
            Else
                padhour = shour.Value.ToString
            End If
            If sminute.Value.ToString.Length < 2 Then
                padminute = "0" & sminute.Value.ToString
            Else
                padminute = sminute.Value.ToString
            End If
            If ssecond.Value.ToString.Length < 2 Then
                padsecond = "0" & ssecond.Value.ToString
            Else
                padsecond = ssecond.Value.ToString
            End If
            scheduledtime.Text = padhour & ":" & padminute & ":" & padsecond
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub shour_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles shour.ValueChanged, shour.TextChanged
        update_scheduled_time()
    End Sub

    Private Sub sminute_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sminute.ValueChanged, sminute.TextChanged
        update_scheduled_time()
    End Sub

    Private Sub ssecond_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ssecond.ValueChanged, ssecond.TextChanged
        update_scheduled_time()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If currenttime.Text.Substring(0, 8) = scheduledtime.Text.Substring(0, 8) Then
            force_check()
        End If
    End Sub

    Private Sub hide_application()
        Try
            Me.Opacity = 0
            Me.Refresh()
            Me.WindowState = FormWindowState.Minimized

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub MenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem4.Click
        hide_application()
    End Sub
End Class
