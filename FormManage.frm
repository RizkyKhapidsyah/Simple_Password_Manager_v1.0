VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Begin VB.Form FormManage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Data"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   13815
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormManage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   13815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   8655
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   13575
      Begin VB.CommandButton cmCari 
         Caption         =   "&Cari"
         Height          =   975
         Left            =   2760
         Picture         =   "FormManage.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   7560
         Width           =   1215
      End
      Begin VB.CommandButton cmTutup 
         Caption         =   "&Tutup"
         Height          =   975
         Left            =   6720
         Picture         =   "FormManage.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   7560
         Width           =   1215
      End
      Begin MSAdodcLib.Adodc AdodcUtama 
         Height          =   330
         Left            =   9720
         Top             =   7560
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
      Begin VB.CommandButton cmAkhir 
         Caption         =   ">>>"
         Height          =   495
         Left            =   12840
         TabIndex        =   9
         Top             =   7560
         Width           =   615
      End
      Begin VB.CommandButton cmSelanjutnya 
         Caption         =   ">"
         Height          =   495
         Left            =   12240
         TabIndex        =   8
         Top             =   7560
         Width           =   615
      End
      Begin VB.CommandButton cmSebelumnya 
         Caption         =   "<"
         Height          =   495
         Left            =   11640
         TabIndex        =   7
         Top             =   7560
         Width           =   615
      End
      Begin VB.CommandButton cmAwal 
         Caption         =   "<<<"
         Height          =   495
         Left            =   11040
         TabIndex        =   6
         Top             =   7560
         Width           =   615
      End
      Begin VB.CommandButton cmFilter 
         Caption         =   "&Filter"
         Height          =   975
         Left            =   5400
         Picture         =   "FormManage.frx":0896
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   7560
         Width           =   1215
      End
      Begin VB.CommandButton cmSorot 
         Caption         =   "&Sorot"
         Height          =   975
         Left            =   4080
         Picture         =   "FormManage.frx":09E0
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   7560
         Width           =   1215
      End
      Begin VB.CommandButton cmHapus 
         Caption         =   "&Hapus"
         Height          =   975
         Left            =   1440
         Picture         =   "FormManage.frx":0B2A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   7560
         Width           =   1215
      End
      Begin VB.CommandButton cmEdit 
         Caption         =   "&Edit"
         Height          =   975
         Left            =   120
         Picture         =   "FormManage.frx":0C74
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   7560
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   7215
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   13335
         _ExtentX        =   23521
         _ExtentY        =   12726
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   21
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Agency FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Agency FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   8520
         Top             =   7680
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin XPEngine.XPControl MesinXP 
      Left            =   0
      Top             =   0
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin VB.Menu menuLainnya 
      Caption         =   "Lainnya"
      Begin VB.Menu menuRefresh 
         Caption         =   "Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuExport 
         Caption         =   "Export"
      End
      Begin VB.Menu menuBF 
         Caption         =   "Bersihkan Format"
      End
      Begin VB.Menu menuTP 
         Caption         =   "Tabel Properties"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu menuSSZ 
         Caption         =   "Set Size Datagrid"
      End
   End
End
Attribute VB_Name = "FormManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Program : Simple Password Manager v1.0
'Source Code by Rizky Khafitsyah
'kunjungi http://rikymetalist.blogspot.com


Sub AturKontrol()
On Error GoTo HancurkanError
NyambungUtama
    With AdodcUtama
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * From TabelPasswordManager"
        Set DataGrid1.DataSource = AdodcUtama
        .Refresh
    End With
    With DataGrid1
        .AllowUpdate = False
        .RowHeight = FormSizeDatagrid.LabelValueBaris.Caption
    End With
    With Me
        .menuLainnya.Visible = False
    End With
    MesinXP.StartEngine
Exit Sub
HancurkanError:
    PusatError
End Sub

Private Sub cmAkhir_Click()
    AdodcUtama.Recordset.MoveLast
End Sub

Private Sub cmAwal_Click()
    AdodcUtama.Recordset.MoveFirst
End Sub

Private Sub cmCari_Click()
If AdodcUtama.Recordset.RecordCount = 0 Then
    MsgBox "Tidak bisa mencari, karena data belum ada!", vbExclamation + vbOKOnly, ""
Else
    FormCari.Show vbModal, Me
End If
End Sub

Private Sub cmEdit_Click()
On Error GoTo HancurkanError
If AdodcUtama.Recordset.RecordCount = 0 Then
    MsgBox "Apa yang hendak diedit? sedangkan datanya aja belum ada -_-", vbExclamation + vbOKOnly, ""
Else
    With FormDataBaru
        .Caption = "Edit Data"
        .cmSimpan.Caption = "&Update"
        .cmExport.Enabled = True
        
        .textPemilikAkun.Text = AdodcUtama.Recordset.Fields(0).Value
        .cmbJenisAkun.Text = AdodcUtama.Recordset.Fields(1).Value
        .textNamaAkun.Text = AdodcUtama.Recordset.Fields(2).Value
        .textUserName.Text = AdodcUtama.Recordset.Fields(3).Value
        .textEmail.Text = AdodcUtama.Recordset.Fields(4).Value
        .textPassword.Text = AdodcUtama.Recordset.Fields(5).Value
        .textAlamatWeb.Text = AdodcUtama.Recordset.Fields(6).Value
        .cmbHari.Text = AdodcUtama.Recordset.Fields(7).Value
        .cmbTanggal.Text = AdodcUtama.Recordset.Fields(8).Value
        .cmbBulan.Text = AdodcUtama.Recordset.Fields(9).Value
        .cmbTahun.Text = AdodcUtama.Recordset.Fields(10).Value
        .textJam.Text = AdodcUtama.Recordset.Fields(11).Value
        .textMenit.Text = AdodcUtama.Recordset.Fields(12).Value
        .textDetik.Text = AdodcUtama.Recordset.Fields(13).Value
        .textNamaKomputer.Text = AdodcUtama.Recordset.Fields(14).Value
        
        .Icon = LoadPicture(App.Path & "\EDIT.ICO")
        
        
        .Show vbModal, Me
    End With
End If
Exit Sub
HancurkanError:
    PusatError
End Sub

Private Sub cmFilter_Click()
If AdodcUtama.Recordset.RecordCount = 0 Then
    MsgBox "Tidak bisa memfilter, karena data belum ada!", vbExclamation + vbOKOnly, ""
Else
    menuBF_Click
    FormFilter.Show vbModal, Me
End If
End Sub

Private Sub cmHapus_Click()
If AdodcUtama.Recordset.RecordCount = 0 Then
    MsgBox "Tidak bisa dihapus, karena data belum ada!", vbExclamation + vbOKOnly, ""
Else
    Z = MsgBox("Anda yakin ingin menghapus data ini?", vbQuestion + vbYesNo, "Hapus?")
    If Z = vbYes Then
        AdodcUtama.Recordset.Delete
    End If
        FormUtama.AturKontrol
        FormUtama.AturKontrol
End If
End Sub

Private Sub cmSebelumnya_Click()
    AdodcUtama.Recordset.MovePrevious
    If AdodcUtama.Recordset.BOF = True Then AdodcUtama.Recordset.MoveLast
End Sub

Private Sub cmSelanjutnya_Click()
    AdodcUtama.Recordset.MoveNext
    If AdodcUtama.Recordset.EOF = True Then AdodcUtama.Recordset.MoveFirst
End Sub

Private Sub cmSorot_Click()
If AdodcUtama.Recordset.RecordCount = 0 Then
    MsgBox "Tidak bisa menyorot, karena data belum ada!", vbExclamation + vbOKOnly, ""
Else
    FormSorot.Show vbModal, Me
End If
End Sub

Private Sub cmTutup_Click()
    FormUtama.AturKontrol
    Unload Me
End Sub

Private Sub DataGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu menuLainnya
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub


Private Sub menuBF_Click()
    AturKontrol
    With FormManage
        .cmEdit.Enabled = True
        .cmHapus.Enabled = True
        .cmCari.Enabled = True
        .cmSorot.Enabled = True
    End With
End Sub

Private Sub MenuExport_Click()
On Error GoTo ErrorHandler
    CommonDialog1.Filter = "RikySoft Catatan Files (*.rcf)|*.rcf|RikySoft Email Files (*.rmd)|*.rmd|Text Files (*.txt)|*.txt|Rich Text Format (*.rtf)|*.rtf|Word 2007 Files (*.docx)|*.docx|Word 2003 Files (*.doc)|*.doc|Excel 2003 Files (*.xls)|*.xls|All Files (*.*)|*.*"
    AturDefaultFormat
    CommonDialog1.ShowSave
    CommonDialog1.FileName = CommonDialog1.FileName
    
    Dim iFile As Integer
    Dim SaveFileFromTB As Boolean
    Dim TxtBox As Object
    Dim FilePath As String
    Dim Append As Boolean
    iFile = FreeFile
    If Append Then
    Open CommonDialog1.FileName For Append As #iFile
    Else
    Open CommonDialog1.FileName For Output As #iFile
    End If
    
    Print #iFile, AdodcUtama.Recordset.Fields(0).Name & " : " & AdodcUtama.Recordset.Fields(0).Value
    Print #iFile, AdodcUtama.Recordset.Fields(1).Name & " : " & AdodcUtama.Recordset.Fields(1).Value
    Print #iFile, AdodcUtama.Recordset.Fields(2).Name & " : " & AdodcUtama.Recordset.Fields(2).Value
    Print #iFile, AdodcUtama.Recordset.Fields(3).Name & " : " & AdodcUtama.Recordset.Fields(3).Value
    Print #iFile, AdodcUtama.Recordset.Fields(4).Name & " : " & AdodcUtama.Recordset.Fields(4).Value
    Print #iFile, AdodcUtama.Recordset.Fields(5).Name & " : " & AdodcUtama.Recordset.Fields(5).Value
    Print #iFile, AdodcUtama.Recordset.Fields(6).Name & " : " & AdodcUtama.Recordset.Fields(6).Value
    Print #iFile, AdodcUtama.Recordset.Fields(7).Name & " : " & AdodcUtama.Recordset.Fields(7).Value
    Print #iFile, AdodcUtama.Recordset.Fields(8).Name & " : " & AdodcUtama.Recordset.Fields(8).Value
    Print #iFile, AdodcUtama.Recordset.Fields(9).Name & " : " & AdodcUtama.Recordset.Fields(9).Value
    Print #iFile, AdodcUtama.Recordset.Fields(10).Name & " : " & AdodcUtama.Recordset.Fields(10).Value
    Print #iFile, AdodcUtama.Recordset.Fields(11).Name & " : " & AdodcUtama.Recordset.Fields(11).Value
    Print #iFile, AdodcUtama.Recordset.Fields(12).Name & " : " & AdodcUtama.Recordset.Fields(12).Value
    Print #iFile, AdodcUtama.Recordset.Fields(13).Name & " : " & AdodcUtama.Recordset.Fields(13).Value
    Print #iFile, AdodcUtama.Recordset.Fields(14).Name & " : " & AdodcUtama.Recordset.Fields(14).Value
    
    SaveFileFromTB = True
ErrorHandler:
    Close #iFile
End Sub

Private Sub menuRefresh_Click()
    AturKontrol
End Sub

Private Sub menuSSZ_Click()
    FormSizeDatagrid.Show vbModal, Me
End Sub

Private Sub menuTP_Click()
    FormTabelProperties.Show vbModal, Me
End Sub
