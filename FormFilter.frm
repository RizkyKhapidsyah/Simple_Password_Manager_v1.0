VERSION 5.00
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Begin VB.Form FormFilter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filter Data"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5025
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormFilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbFilterBerdasarkan 
      Height          =   390
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
   Begin VB.ComboBox cmbMode 
      Height          =   390
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   2655
   End
   Begin VB.CommandButton cmFilter 
      Caption         =   "&Filter"
      Height          =   615
      Left            =   3840
      Picture         =   "FormFilter.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
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
      Caption         =   "Filter Data Berdasarkan"
      Height          =   270
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1530
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   2160
      TabIndex        =   5
      Top             =   120
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mode"
      Height          =   270
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   315
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   2160
      TabIndex        =   3
      Top             =   600
      Width           =   45
   End
End
Attribute VB_Name = "FormFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Program : Simple Password Manager v1.0
'Source Code by Rizky Khafitsyah
'kunjungi http://rikymetalist.blogspot.com

Sub AturKontrol()
    With cmbFilterBerdasarkan
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
    With cmbMode
        .Clear
        .AddItem "Asc", 0
        .AddItem "Desc", 1
        .ListIndex = 0
    End With
    MesinXP.StartEngine
End Sub

Private Sub cmFilter_Click()
    With FormManage.AdodcUtama
        .Refresh
        .RecordSource = "Select " & cmbFilterBerdasarkan.Text & " from TabelPasswordManager order by " & cmbFilterBerdasarkan.Text & " " & cmbMode.Text & "; "
        .Refresh
    End With
    With FormManage
        .cmEdit.Enabled = False
        .cmHapus.Enabled = False
        .cmCari.Enabled = False
        .cmSorot.Enabled = False
    End With
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub
