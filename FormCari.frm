VERSION 5.00
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Begin VB.Form FormCari 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cari Data"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6105
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormCari.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmCari 
      Caption         =   "&Cari"
      Height          =   735
      Left            =   4680
      Picture         =   "FormCari.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox textKriteria 
      Height          =   390
      Left            =   2280
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   600
      Width           =   2295
   End
   Begin VB.ComboBox cmbBerdasarkan 
      Height          =   390
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   3735
   End
   Begin XPEngine.XPControl MesinXP 
      Left            =   0
      Top             =   0
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   2160
      TabIndex        =   4
      Top             =   600
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dengan Kriteria"
      Height          =   270
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cari Data Berdasarkan "
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1530
   End
End
Attribute VB_Name = "FormCari"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Program : Simple Password Manager v1.0
'Source Code by Rizky Khafitsyah
'kunjungi http://rikymetalist.blogspot.com


Sub AturKontrol()
    cmbBerdasarkan.Clear
    With cmbBerdasarkan
        .AddItem FormManage.AdodcUtama.Recordset.Fields(0).Name
        .AddItem FormManage.AdodcUtama.Recordset.Fields(1).Name
        .AddItem FormManage.AdodcUtama.Recordset.Fields(2).Name
        .AddItem FormManage.AdodcUtama.Recordset.Fields(3).Name
        .AddItem FormManage.AdodcUtama.Recordset.Fields(4).Name
        .AddItem FormManage.AdodcUtama.Recordset.Fields(5).Name
        .AddItem FormManage.AdodcUtama.Recordset.Fields(6).Name
        .AddItem FormManage.AdodcUtama.Recordset.Fields(7).Name
        .AddItem FormManage.AdodcUtama.Recordset.Fields(8).Name
        .AddItem FormManage.AdodcUtama.Recordset.Fields(9).Name
        .AddItem FormManage.AdodcUtama.Recordset.Fields(10).Name
        .AddItem FormManage.AdodcUtama.Recordset.Fields(11).Name
        .AddItem FormManage.AdodcUtama.Recordset.Fields(12).Name
        .AddItem FormManage.AdodcUtama.Recordset.Fields(13).Name
        .AddItem FormManage.AdodcUtama.Recordset.Fields(14).Name
        .ListIndex = 0
    End With
    textKriteria.Text = ""
        MesinXP.StartEngine
End Sub

Private Sub cmCari_Click()
    If textKriteria.Text = "" Then
        MsgBox "Silahkan isi kriteria yang ingin dicari.", vbExclamation + vbOKOnly, ""
        textKriteria.SetFocus
    Else
    FormManage.AdodcUtama.Refresh
    With FormManage.AdodcUtama.Recordset
        Select Case cmbBerdasarkan.ListIndex
        Case Is = 0
            .Find "Pemilik_Akun = '" & textKriteria.Text & "'"
        Case Is = 1
            .Find "Jenis_Akun = '" & textKriteria.Text & "'"
        Case Is = 2
            .Find "Nama_Akun = '" & textKriteria.Text & "'"
        Case Is = 3
            .Find "User_Name = '" & textKriteria.Text & "'"
        Case Is = 4
            .Find "Email = '" & textKriteria.Text & "'"
        Case Is = 5
            .Find "Password = '" & textKriteria.Text & "'"
        Case Is = 6
            .Find "Alamat_Web = '" & textKriteria.Text & "'"
        Case Is = 7
            .Find "Hari_Simpan = '" & textKriteria.Text & "'"
        Case Is = 8
            .Find "Tanggal_Simpan = '" & textKriteria.Text & "'"
        Case Is = 9
            .Find "Bulan_Simpan = '" & textKriteria.Text & "'"
        Case Is = 10
            .Find "Tahun_Simpan = '" & textKriteria.Text & "'"
        Case Is = 11
            .Find "Jam_Simpan = '" & textKriteria.Text & "'"
        Case Is = 12
            .Find "Menit_Simpan = '" & textKriteria.Text & "'"
        Case Is = 13
            .Find "Detik_Simpan = '" & textKriteria.Text & "'"
        Case Is = 14
            .Find "Nama_Komputer = '" & textKriteria.Text & "'"
        End Select
        If .EOF Then
            MsgBox ("Data tidak ditemukan"), vbExclamation + vbOKOnly, ""
        Else
            Set FormManage.DataGrid1.DataSource = FormManage.AdodcUtama.Recordset
            If FormPengaturan.cekTutupFormCari.Value = Checked Then Me.Hide
        End If
    End With
    End If
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub
