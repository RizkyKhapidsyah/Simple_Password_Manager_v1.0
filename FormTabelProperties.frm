VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Begin VB.Form FormTabelProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabel Properties"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4305
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormTabelProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc AdodcUpdateTerakhir 
      Height          =   330
      Left            =   240
      Top             =   2640
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
   Begin VB.CommandButton cmTutup 
      Caption         =   "&Tutup"
      Height          =   975
      Left            =   3360
      Picture         =   "FormTabelProperties.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2400
      Width           =   855
   End
   Begin XPEngine.XPControl MesinXP 
      Left            =   0
      Top             =   0
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin VB.Label labelUPdateTerakhir 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      Height          =   270
      Left            =   1680
      TabIndex        =   17
      Top             =   1920
      Width           =   420
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1560
      TabIndex        =   16
      Top             =   1920
      Width           =   45
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Update Terakhir"
      Height          =   270
      Left            =   120
      TabIndex        =   15
      Top             =   1920
      Width           =   1050
   End
   Begin VB.Label labelJumlahCell 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      Height          =   270
      Left            =   1680
      TabIndex        =   14
      Top             =   1560
      Width           =   420
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1560
      TabIndex        =   13
      Top             =   1560
      Width           =   45
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Cell"
      Height          =   270
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label labelJumlahData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      Height          =   270
      Left            =   1680
      TabIndex        =   11
      Top             =   1200
      Width           =   420
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1560
      TabIndex        =   10
      Top             =   1200
      Width           =   45
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Data"
      Height          =   270
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   795
   End
   Begin VB.Label labelConnection 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      Height          =   270
      Left            =   1680
      TabIndex        =   8
      Top             =   840
      Width           =   420
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1560
      TabIndex        =   7
      Top             =   840
      Width           =   45
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Connection"
      Height          =   270
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   690
   End
   Begin VB.Label LabelJenis 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      Height          =   270
      Left            =   1680
      TabIndex        =   5
      Top             =   480
      Width           =   420
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1560
      TabIndex        =   4
      Top             =   480
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jenis"
      Height          =   270
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   345
   End
   Begin VB.Label labelNamaTabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      Height          =   270
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Tabel"
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   750
   End
End
Attribute VB_Name = "FormTabelProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Program : Simple Password Manager v1.0
'Source Code by Rizky Khafitsyah
'kunjungi http://rikymetalist.blogspot.com


Private Sub cmTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    NyambungUtama
    With AdodcUpdateTerakhir
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * from TabelUpdateTerakhir"
        .Refresh
    End With
    labelNamaTabel.Caption = "TabelPasswordManager"
    LabelJenis.Caption = "MyISAM"
    labelConnection.Caption = "ADO"
    labelJumlahData.Caption = FormManage.AdodcUtama.Recordset.RecordCount
    labelJumlahCell.Caption = Val(FormManage.AdodcUtama.Recordset.RecordCount) * Val(FormManage.AdodcUtama.Recordset.Fields.Count)
    labelUPdateTerakhir.Caption = AdodcUpdateTerakhir.Recordset.Fields(0).Value
    MesinXP.StartEngine
End Sub
