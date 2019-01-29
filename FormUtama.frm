VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Begin VB.Form FormUtama 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RikySoft - Simple Password Manager  - v1.0"
   ClientHeight    =   9600
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   17040
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormUtama.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   17040
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmAbout 
      Caption         =   "&Tentang"
      Height          =   855
      Left            =   6720
      Picture         =   "FormUtama.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton cmHirarkiView 
      Caption         =   "&Hirarki View"
      Height          =   855
      Left            =   2760
      Picture         =   "FormUtama.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Timer TimerWaktu 
      Interval        =   10
      Left            =   10080
      Top             =   8640
   End
   Begin MSComctlLib.StatusBar StatusBawah 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   9225
      Width           =   17040
      _ExtentX        =   30057
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   21167
            MinWidth        =   21167
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmPengaturan 
      Caption         =   "&Pengaturan"
      Height          =   855
      Left            =   5400
      Picture         =   "FormUtama.frx":06D6
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8280
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc AdodcUtama 
      Height          =   330
      Left            =   6120
      Top             =   8760
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmKeluar 
      Caption         =   "&Keluar"
      Height          =   855
      Left            =   15720
      Picture         =   "FormUtama.frx":0B18
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton cmRefresh 
      Caption         =   "&Refresh"
      Height          =   855
      Left            =   4080
      Picture         =   "FormUtama.frx":0F5A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton cmManage 
      Caption         =   "&Manage"
      Height          =   855
      Left            =   1440
      Picture         =   "FormUtama.frx":10A4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8280
      Width           =   1215
   End
   Begin MSComctlLib.ListView LV 
      Height          =   8055
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   16815
      _ExtentX        =   29660
      _ExtentY        =   14208
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmBaru 
      Caption         =   "&Baru"
      Height          =   855
      Left            =   120
      Picture         =   "FormUtama.frx":13AE
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8280
      Width           =   1215
   End
   Begin XPEngine.XPControl MesinXP 
      Left            =   0
      Top             =   0
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
End
Attribute VB_Name = "FormUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Program : Simple Password Manager v1.0
'Source Code by Rizky Khafitsyah
'kunjungi http://rikymetalist.blogspot.com


Sub AturKontrol()
    NyambungUtama
    With AdodcUtama
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * from TabelPasswordManager order by Nama_Akun asc;"
        .Refresh
    End With
    PENGATURAN_FORM_UTAMA
    With StatusBawah.Panels
        .Item(1).ToolTipText = "Tanggal Saat Ini"
        .Item(2).ToolTipText = "Waktu Saat Ini"
        .Item(3).ToolTipText = "Jumlah Record (Data)"
        .Item(4).ToolTipText = "Jumlah Cell Asli"
        .Item(3).Text = AdodcUtama.Recordset.RecordCount & " Record(s) Data"
        .Item(4).Text = Val(AdodcUtama.Recordset.RecordCount) * Val(AdodcUtama.Recordset.Fields.Count) & " Cell(s)"
        .Item(1).Alignment = sbrCenter
        .Item(2).Alignment = sbrCenter
        .Item(3).Alignment = sbrCenter
        .Item(4).Alignment = sbrLeft
    End With
        Call CekProgram(FormUtama)
        MesinXP.StartEngine
End Sub
Sub PENGATURAN_FORM_UTAMA()
    With LV
        .ColumnHeaders.Clear
        .ListItems.Clear
        .View = lvwReport
        .Sorted = True
        If FormPengaturan.cekGridlines.Value = Checked Then
            .Gridlines = True
        ElseIf FormPengaturan.cekGridlines.Value = Unchecked Then
            .Gridlines = False
        End If
    End With
    If FormPengaturan.cekTampilkanPassword.Value = Checked Then
        With LV
            .ColumnHeaders.Clear
            .ColumnHeaders.Add , , "Pemilik Akun", 1500
            .ColumnHeaders.Add , , "Jenis Akun", 2000, vbCenter
            .ColumnHeaders.Add , , "Nama Akun", 2000, vbCenter
            .ColumnHeaders.Add , , "User Name", 2000, vbCenter
            .ColumnHeaders.Add , , "E-Mail", 2000, vbCenter
            .ColumnHeaders.Add , , "Password", 2000, vbCenter
            .ColumnHeaders.Add , , "Alamat Web", 2000, vbCenter
            .ColumnHeaders.Add , , "Tanggal Simpan", 2000, vbCenter
            .ColumnHeaders.Add , , "Waktu Simpan", 1200, vbCenter
            .ColumnHeaders.Add , , "Nama Komputer", 2000, vbCenter
        End With
        'MASUKKAN DATABASE KE LISTVIEW
            Do Until AdodcUtama.Recordset.EOF
                Set LI = LV.ListItems.Add(, , AdodcUtama.Recordset.Fields(0).Value)
                    LI.SubItems(1) = AdodcUtama.Recordset.Fields(1).Value
                    LI.SubItems(2) = AdodcUtama.Recordset.Fields(2).Value
                    LI.SubItems(3) = AdodcUtama.Recordset.Fields(3).Value
                    LI.SubItems(4) = AdodcUtama.Recordset.Fields(4).Value
                    LI.SubItems(5) = AdodcUtama.Recordset.Fields(5).Value
                    LI.SubItems(6) = AdodcUtama.Recordset.Fields(6).Value
                    LI.SubItems(7) = AdodcUtama.Recordset.Fields(7).Value & ", " & AdodcUtama.Recordset.Fields(8).Value & " - " & AdodcUtama.Recordset.Fields(9).Value & " - " & AdodcUtama.Recordset.Fields(10).Value
                    LI.SubItems(8) = AdodcUtama.Recordset.Fields(11).Value & ":" & AdodcUtama.Recordset.Fields(12).Value & ":" & AdodcUtama.Recordset.Fields(13).Value
                    LI.SubItems(9) = AdodcUtama.Recordset.Fields(14).Value
                    AdodcUtama.Recordset.MoveNext
                Loop
    ElseIf FormPengaturan.cekTampilkanPassword.Value = Unchecked Then
        With LV
            .ColumnHeaders.Clear
            .ColumnHeaders.Add , , "Pemilik Akun", 1500
            .ColumnHeaders.Add , , "Jenis Akun", 2000, vbCenter
            .ColumnHeaders.Add , , "Nama Akun", 2000, vbCenter
            .ColumnHeaders.Add , , "User Name", 2000, vbCenter
            .ColumnHeaders.Add , , "E-Mail", 2000, vbCenter
            .ColumnHeaders.Add , , "Alamat Web", 2000, vbCenter
            .ColumnHeaders.Add , , "Tanggal Simpan", 2000, vbCenter
            .ColumnHeaders.Add , , "Waktu Simpan", 1200, vbCenter
            .ColumnHeaders.Add , , "Nama Komputer", 2000, vbCenter
        End With
        'MASUKKAN DATABASE KE LISTVIEW
            Do Until AdodcUtama.Recordset.EOF
                Set LI = LV.ListItems.Add(, , AdodcUtama.Recordset.Fields(0).Value)
                    LI.SubItems(1) = AdodcUtama.Recordset.Fields(1).Value
                    LI.SubItems(2) = AdodcUtama.Recordset.Fields(2).Value
                    LI.SubItems(3) = AdodcUtama.Recordset.Fields(3).Value
                    LI.SubItems(4) = AdodcUtama.Recordset.Fields(4).Value
                    LI.SubItems(5) = AdodcUtama.Recordset.Fields(6).Value
                    LI.SubItems(6) = AdodcUtama.Recordset.Fields(7).Value & ", " & AdodcUtama.Recordset.Fields(8).Value & " - " & AdodcUtama.Recordset.Fields(9).Value & " - " & AdodcUtama.Recordset.Fields(10).Value
                    LI.SubItems(7) = AdodcUtama.Recordset.Fields(11).Value & ":" & AdodcUtama.Recordset.Fields(12).Value & ":" & AdodcUtama.Recordset.Fields(13).Value
                    LI.SubItems(8) = AdodcUtama.Recordset.Fields(14).Value
                    AdodcUtama.Recordset.MoveNext
                Loop
    End If
End Sub
'bagian untuk mendeteksi jika aplikasi dijalankan 2x (aplikasi seharusnya tidak boleh berjalan 2 kali, okey!)
Public Sub CekProgram(X As Form)
On Error GoTo Pesan
    Dim SaveTitle$
    If App.PrevInstance Then
        SaveTitle$ = App.Title
        MsgBox "Program ini sedang dijalankan!", _
               vbCritical, "Sedang Dijalankan"
        App.Title = ""
        X.Caption = ""
        AppActivate SaveTitle$
        SendKeys "%{ENTER}", True
        End
    End If
    Exit Sub
Pesan:
    End
    Exit Sub
End Sub

Private Sub cmAbout_Click()
MsgBox "RikySoft Simple Password Manager v1.0" & vbCrLf & _
        "by Rizky Khafitsyah." & vbCrLf & _
        "Copyright(c)_2010 by RikySoft Software House Foundation" & vbCrLf & vbCrLf & _
        "Kunjungi : http://rikymetalist.blogspot.com", vbInformation + vbOKOnly, "Tentang..."
End Sub

Private Sub cmBaru_Click()
    With FormDataBaru
        .Icon = LoadPicture(App.Path & "\NEW16.ico")
        .Show vbModal, Me
    End With
End Sub

Private Sub cmHirarkiView_Click()
If AdodcUtama.Recordset.RecordCount = 0 Then
    MsgBox "Akses dibatasi!", vbCritical + vbOKOnly, ""
Else
    FormHirarkiView.Show vbModal, Me
End If
End Sub

Private Sub cmKeluar_Click()
    Z = MsgBox("Apakah Anda yakin ingin keluar?", vbQuestion + vbYesNo, "Keluar?")
    If Z = vbYes Then
        End
    End If
End Sub

Private Sub cmManage_Click()
If AdodcUtama.Recordset.RecordCount = 0 Then
    Z = MsgBox("Data Belum ada, Anda yakin ingin memanage?", vbQuestion + vbYesNo, "")
    If Z = vbYes Then FormManage.Show vbModal, Me
Else
    FormManage.Show vbModal, Me
End If
End Sub

Private Sub cmPengaturan_Click()
    With FormPengaturan
        .Show vbModal, Me
    End With
End Sub

Private Sub cmRefresh_Click()
    AturKontrol
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Z = MsgBox("Apakah Anda yakin ingin keluar?", vbQuestion + vbYesNo, "Keluar?")
    If Z = vbYes Then
        End
    Else
        Cancel = 1
    End If
End Sub

Private Sub TimerWaktu_Timer()
    With StatusBawah.Panels
        .Item(1).Text = Day(Date) & " - " & Month(Date) & " - " & Year(Date)
        .Item(2).Text = Time
    End With
End Sub
