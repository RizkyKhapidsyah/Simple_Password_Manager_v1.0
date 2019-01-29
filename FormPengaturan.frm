VERSION 5.00
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Begin VB.Form FormPengaturan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pengaturan"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormPengaturan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Default Dimpan"
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   2775
      Begin VB.ComboBox cmbDefaultSimpan 
         Height          =   390
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5175
      Begin VB.CheckBox cekIzinkanPenambahanKategoriAkun 
         Caption         =   "Izinkan Penambahan Kategori Akun"
         Height          =   270
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CommandButton cmAturKategoriAKun 
         Caption         =   "&Atur Kategori"
         Height          =   375
         Left            =   3600
         TabIndex        =   10
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CheckBox cekTampilkanPassword 
         Caption         =   "Tampilkan Kolom Password"
         Height          =   270
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3135
      End
      Begin VB.CheckBox cekGridlines 
         Caption         =   "Tampilkan Gridlines Tabel"
         Height          =   270
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Value           =   1  'Checked
         Width           =   3135
      End
      Begin VB.CheckBox cekTutupFormINput 
         Caption         =   "Tutup Form Input Setelah Data Disimpan (Ditambah)"
         Height          =   270
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Value           =   1  'Checked
         Width           =   4695
      End
      Begin VB.CheckBox CekKonfirmasiSimpanEditData 
         Caption         =   "Konfirmasi untuk simpan/edit data"
         Height          =   270
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Value           =   1  'Checked
         Width           =   4695
      End
      Begin VB.CheckBox cekTutupFormCari 
         Caption         =   "Tutup Form Cari setelah data berhasil dicari"
         Height          =   270
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   4695
      End
   End
   Begin VB.CommandButton cmBatal 
      Caption         =   "&Batal"
      Height          =   975
      Left            =   3000
      Picture         =   "FormPengaturan.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmOK 
      Caption         =   "&OK"
      Height          =   975
      Left            =   4200
      Picture         =   "FormPengaturan.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2760
      Width           =   1095
   End
   Begin XPEngine.XPControl MesinXP 
      Left            =   0
      Top             =   0
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
End
Attribute VB_Name = "FormPengaturan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Program : Simple Password Manager v1.0
'Source Code by Rizky Khafitsyah
'kunjungi http://rikymetalist.blogspot.com


Sub AturKontrol()
    With cmbDefaultSimpan
        .Clear
        .AddItem "RikySoft Catatan Files (*.rcf)", 0
        .AddItem "RikySoft Email Files (*.rmd)", 1
        .AddItem "Text Files (*.txt)", 2
        .AddItem "Rich Text Format (*.rtf)", 3
        .AddItem "Word 2007 Files (*.docx)", 4
        .AddItem "Word 2003 Files (*.doc)", 5
        .AddItem "Excel 2003 Files (*.xls)", 6
        .ListIndex = 0
    End With
    AmbilPengaturan
    MesinXP.StartEngine
End Sub
'dalam program ini, kita menyimpan settingannya ke registry aja yaa, hehe
Sub SimpanPengaturan()
    SaveSetting App.Title, Me.Name, Me.cekTampilkanPassword.Name, Me.cekTampilkanPassword.Value
    SaveSetting App.Title, Me.Name, Me.cekGridlines.Name, Me.cekGridlines.Value
    SaveSetting App.Title, Me.Name, Me.CekKonfirmasiSimpanEditData.Name, Me.CekKonfirmasiSimpanEditData.Value
    SaveSetting App.Title, Me.Name, Me.cekTutupFormINput.Name, Me.cekTutupFormINput.Value
    SaveSetting App.Title, Me.Name, Me.cekTutupFormCari.Name, Me.cekTutupFormCari.Value
    SaveSetting App.Title, Me.Name, Me.cmbDefaultSimpan.Name, Me.cmbDefaultSimpan.ListIndex
    SaveSetting App.Title, Me.Name, Me.cekIzinkanPenambahanKategoriAkun.Name, Me.cekIzinkanPenambahanKategoriAkun.Value
End Sub
Sub AmbilPengaturan()
    Me.cekTampilkanPassword.Value = GetSetting(App.Title, Me.Name, Me.cekTampilkanPassword.Name, Me.cekTampilkanPassword.Value)
    Me.cekGridlines.Value = GetSetting(App.Title, Me.Name, Me.cekGridlines.Name, Me.cekGridlines.Value)
    Me.CekKonfirmasiSimpanEditData.Value = GetSetting(App.Title, Me.Name, Me.CekKonfirmasiSimpanEditData.Name, Me.CekKonfirmasiSimpanEditData.Value)
    Me.cekTutupFormINput.Value = GetSetting(App.Title, Me.Name, Me.cekTutupFormINput.Name, Me.cekTutupFormINput.Value)
    Me.cekTutupFormCari.Value = GetSetting(App.Title, Me.Name, Me.cekTutupFormCari.Name, Me.cekTutupFormCari.Value)
    Me.cmbDefaultSimpan.ListIndex = GetSetting(App.Title, Me.Name, Me.cmbDefaultSimpan.Name, Me.cmbDefaultSimpan.ListIndex)
    Me.cekIzinkanPenambahanKategoriAkun.Value = GetSetting(App.Title, Me.Name, Me.cekIzinkanPenambahanKategoriAkun.Name, Me.cekIzinkanPenambahanKategoriAkun.Value)
End Sub

Private Sub cekIzinkanPenambahanKategoriAkun_Click()
    If Me.cekIzinkanPenambahanKategoriAkun.Value = Checked Then
        cmAturKategoriAKun.Enabled = True
    Else
        cmAturKategoriAKun.Enabled = False
    End If
End Sub

Private Sub cmAturKategoriAKun_Click()
    FormAturKategoriAkun.Show vbModal, Me
End Sub

Private Sub cmBatal_Click()
    Unload Me
End Sub

Private Sub cmOK_Click()
    SimpanPengaturan
    FormUtama.AturKontrol
    Unload Me
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub
