VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Begin VB.Form FormAturKategoriAkun 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atur Kategori Akun"
   ClientHeight    =   3135
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
   Icon            =   "FormAturKategoriAkun.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XPEngine.XPControl MesinXP 
      Left            =   3600
      Top             =   1920
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin VB.CommandButton cmHapus 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmEdit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmTambah 
      Caption         =   "&Tambah"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmAkhir 
      Caption         =   ">>"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton cmSelanjutnya 
      Caption         =   ">"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton cmSebelumnya 
      Caption         =   "<"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton cmAwal 
      Caption         =   "<<"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   375
   End
   Begin MSAdodcLib.Adodc AdodcUtama 
      Height          =   375
      Left            =   240
      Top             =   4320
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4260
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
End
Attribute VB_Name = "FormAturKategoriAkun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Sub AturKontrol()
    NyambungUtama
    With AdodcUtama
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * from tbJenisAkun"
        Set DataGrid1.DataSource = AdodcUtama
        .Refresh
    End With
    With Me
        .DataGrid1.RowHeight = 315
        .DataGrid1.AllowUpdate = False
        .DataGrid1.Columns(0).Width = 5000
    End With
    MesinXP.StartEngine
End Sub

Private Sub cmAkhir_Click()
    AdodcUtama.Recordset.MoveLast
End Sub

Private Sub cmAwal_Click()
    AdodcUtama.Recordset.MoveFirst
End Sub

Private Sub cmEdit_Click()
    If AdodcUtama.Recordset.RecordCount = 0 Then
        MsgBox "Tidak ada data yang dapat diedit", vbExclamation + vbOKOnly, ""
    Else
        With FormTambahKategori
            .textTambahKategoriBaru.Text = AdodcUtama.Recordset.Fields(0).Value
            .cmSimpan.Caption = "&OK/Update"
            .Caption = "Edit Kategori"
            .Show vbModal, Me
        End With
    End If
End Sub

Private Sub cmHapus_Click()
    If AdodcUtama.Recordset.RecordCount = 0 Then
        MsgBox "Tidak ada data yang dapat dihapus", vbExclamation + vbOKOnly, ""
    Else
        Z = MsgBox("Anda yakin ingin menghapus kategori ini?", vbQuestion + vbYesNo, "Hapus?")
        If Z = vbYes Then
            AdodcUtama.Recordset.Delete
        End If
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

Private Sub cmTambah_Click()
    With FormTambahKategori
        .Show vbModal, Me
    End With
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub
