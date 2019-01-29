VERSION 5.00
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Begin VB.Form FormTambahKategori 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tambah Kategori Baru"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormTambahKategori.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmBatal 
      Caption         =   "&Batal"
      Height          =   975
      Left            =   1680
      Picture         =   "FormTambahKategori.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmSimpan 
      Caption         =   "&OK/Simpan"
      Height          =   975
      Left            =   3120
      Picture         =   "FormTambahKategori.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox textTambahKategoriBaru 
      Height          =   390
      Left            =   1680
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   2775
   End
   Begin XPEngine.XPControl MesinXP 
      Left            =   0
      Top             =   0
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kategori Baru :"
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   990
   End
End
Attribute VB_Name = "FormTambahKategori"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Program : Simple Password Manager v1.0
'Source Code by Rizky Khafitsyah
'kunjungi http://rikymetalist.blogspot.com


Private Sub cmBatal_Click()
    FormDataBaru.MasukkanDatabaseKeCMBJenisAkun
    Unload Me
End Sub

Private Sub cmSimpan_Click()
    If textTambahKategoriBaru.Text = "" Then
        MsgBox "Silahkan isi nama kategori baru yang ingin ditambahkan.", vbExclamation + vbOKOnly, ""
        textTambahKategoriBaru.SetFocus
    Else
        If cmSimpan.Caption = "&OK/Simpan" Then
            Z = MsgBox("Apakah anda ingin menambahkan '" & textTambahKategoriBaru.Text & "' sebagai kategori baru?" & vbCrLf & _
                        "Peringatan : Data Tidak dapat dihapus secara langsung", vbQuestion + vbYesNo, "Konfirmasi")
            If Z = vbYes Then
                With FormDataBaru
                    .AdodcJenisAkun.Recordset.AddNew
                    .AdodcJenisAkun.Recordset.Fields(0).Value = textTambahKategoriBaru.Text
                    .AdodcJenisAkun.Recordset.Update
                    .AdodcJenisAkun.Refresh
                    .MasukkanDatabaseKeCMBJenisAkun
                    .cmbJenisAkun.Text = textTambahKategoriBaru.Text
                End With
                With FormAturKategoriAkun
                    .AdodcUtama.Recordset.AddNew
                    .AdodcUtama.Recordset.Fields(0).Value = textTambahKategoriBaru.Text
                    .AdodcUtama.Recordset.Update
                    .AdodcUtama.Refresh
                    .AturKontrol
                End With
                Unload Me
            End If
        ElseIf cmSimpan.Caption = "&OK/Update" Then
            Z = MsgBox("Apakah anda ingin menyimpan hasil editan?", vbQuestion + vbYesNo, "Konfirmasi")
            If Z = vbYes Then
                With FormAturKategoriAkun
                    .AdodcUtama.Recordset.Delete
                    .AdodcUtama.Recordset.AddNew
                    .AdodcUtama.Recordset.Fields(0).Value = textTambahKategoriBaru.Text
                    .AdodcUtama.Recordset.Update
                    .AdodcUtama.Refresh
                    .AturKontrol
                End With
                Unload Me
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    textTambahKategoriBaru.Text = ""
    MesinXP.StartEngine
End Sub
