VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Begin VB.Form FormDataBaru 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tambah Data Baru"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6945
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormDataBaru.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox textJenisAkun 
      Height          =   390
      Left            =   7920
      TabIndex        =   64
      Text            =   "Text1"
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox textWaktu 
      Height          =   390
      Left            =   7440
      TabIndex        =   63
      Text            =   "Text1"
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox textTanggal 
      Height          =   390
      Left            =   7440
      TabIndex        =   62
      Text            =   "Text1"
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmHapus 
      Caption         =   "X"
      Height          =   375
      Index           =   6
      Left            =   7440
      TabIndex        =   61
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton cmHapus 
      Caption         =   "X"
      Height          =   375
      Index           =   5
      Left            =   7440
      TabIndex        =   60
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton cmHapus 
      Caption         =   "X"
      Height          =   375
      Index           =   4
      Left            =   7440
      TabIndex        =   59
      Top             =   3840
      Width           =   375
   End
   Begin VB.CommandButton cmHapus 
      Caption         =   "X"
      Height          =   375
      Index           =   3
      Left            =   7440
      TabIndex        =   58
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton cmHapus 
      Caption         =   "X"
      Height          =   375
      Index           =   2
      Left            =   7440
      TabIndex        =   57
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton cmHapus 
      Caption         =   "X"
      Height          =   375
      Index           =   1
      Left            =   7440
      TabIndex        =   56
      Top             =   2400
      Width           =   375
   End
   Begin VB.CommandButton cmHapus 
      Caption         =   "X"
      Height          =   375
      Index           =   0
      Left            =   7440
      TabIndex        =   55
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cmCopy 
      Caption         =   "C"
      Height          =   375
      Index           =   9
      Left            =   6960
      TabIndex        =   54
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton cmCopy 
      Caption         =   "C"
      Height          =   375
      Index           =   8
      Left            =   6960
      TabIndex        =   53
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton cmCopy 
      Caption         =   "C"
      Height          =   375
      Index           =   7
      Left            =   6960
      TabIndex        =   52
      Top             =   3840
      Width           =   375
   End
   Begin VB.CommandButton cmCopy 
      Caption         =   "C"
      Height          =   375
      Index           =   6
      Left            =   6960
      TabIndex        =   51
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton cmCopy 
      Caption         =   "C"
      Height          =   375
      Index           =   5
      Left            =   6960
      TabIndex        =   50
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton cmCopy 
      Caption         =   "C"
      Height          =   375
      Index           =   4
      Left            =   6960
      TabIndex        =   49
      Top             =   2400
      Width           =   375
   End
   Begin VB.CommandButton cmCopy 
      Caption         =   "C"
      Height          =   375
      Index           =   3
      Left            =   6960
      TabIndex        =   48
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cmCopy 
      Caption         =   "C"
      Height          =   375
      Index           =   2
      Left            =   6960
      TabIndex        =   47
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton cmCopy 
      Caption         =   "C"
      Height          =   375
      Index           =   1
      Left            =   6960
      TabIndex        =   46
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton cmCopy 
      Caption         =   "C"
      Height          =   375
      Index           =   0
      Left            =   6960
      TabIndex        =   45
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton cmLebarkan 
      Caption         =   ">>>"
      Height          =   855
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   5400
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3000
      Top             =   6600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdodcJenisAkun 
      Height          =   375
      Left            =   360
      Top             =   6360
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.Timer TimerWaktu 
      Interval        =   10
      Left            =   4320
      Top             =   6960
   End
   Begin VB.CommandButton cmBatal 
      Caption         =   "&Batal"
      Height          =   855
      Left            =   4440
      Picture         =   "FormDataBaru.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton cmExport 
      Caption         =   "&Export"
      Height          =   855
      Left            =   3000
      Picture         =   "FormDataBaru.frx":044E
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton cmReset 
      Caption         =   "&Reset"
      Height          =   855
      Left            =   1560
      Picture         =   "FormDataBaru.frx":1010
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton cmSimpan 
      Caption         =   "&Simpan"
      Height          =   855
      Left            =   120
      Picture         =   "FormDataBaru.frx":115A
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Height          =   3615
      Left            =   120
      TabIndex        =   17
      Top             =   1680
      Width           =   6735
      Begin VB.CommandButton cmTambahKategori 
         Caption         =   "&Tambah"
         Height          =   375
         Left            =   5400
         TabIndex        =   43
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox textAlamatWeb 
         Height          =   390
         Left            =   2280
         TabIndex        =   38
         Text            =   "Text1"
         Top             =   3120
         Width           =   4335
      End
      Begin VB.TextBox textPassword 
         Height          =   390
         Left            =   2280
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   2640
         Width           =   4335
      End
      Begin VB.TextBox textEmail 
         Height          =   390
         Left            =   2280
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   2160
         Width           =   4335
      End
      Begin VB.TextBox textUserName 
         Height          =   390
         Left            =   2280
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   1680
         Width           =   4335
      End
      Begin VB.TextBox textNamaAkun 
         Height          =   390
         Left            =   2280
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   1200
         Width           =   4335
      End
      Begin VB.ComboBox cmbJenisAkun 
         Height          =   390
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox textPemilikAkun 
         Height          =   390
         Left            =   2280
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat Web"
         Height          =   270
         Left            =   120
         TabIndex        =   37
         Top             =   3120
         Width           =   780
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2160
         TabIndex        =   36
         Top             =   3120
         Width           =   45
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         Height          =   270
         Left            =   120
         TabIndex        =   34
         Top             =   2640
         Width           =   690
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2160
         TabIndex        =   33
         Top             =   2640
         Width           =   45
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail"
         Height          =   270
         Left            =   120
         TabIndex        =   31
         Top             =   2160
         Width           =   360
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2160
         TabIndex        =   30
         Top             =   2160
         Width           =   45
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         Height          =   270
         Left            =   120
         TabIndex        =   28
         Top             =   1680
         Width           =   720
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2160
         TabIndex        =   27
         Top             =   1680
         Width           =   45
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Akun"
         Height          =   270
         Left            =   120
         TabIndex        =   25
         Top             =   1200
         Width           =   720
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2160
         TabIndex        =   24
         Top             =   1200
         Width           =   45
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Akun"
         Height          =   270
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   705
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2160
         TabIndex        =   21
         Top             =   720
         Width           =   45
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pemilik Akun"
         Height          =   270
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   810
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2160
         TabIndex        =   18
         Top             =   240
         Width           =   45
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.TextBox textNamaKomputer 
         Height          =   390
         Left            =   2280
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   1200
         Width           =   4335
      End
      Begin VB.TextBox textDetik 
         Alignment       =   2  'Center
         Height          =   390
         Left            =   4200
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox textMenit 
         Alignment       =   2  'Center
         Height          =   390
         Left            =   3240
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox textJam 
         Alignment       =   2  'Center
         Height          =   390
         Left            =   2280
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   720
         Width           =   735
      End
      Begin VB.ComboBox cmbHari 
         Height          =   390
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmbTanggal 
         Height          =   390
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cmbBulan 
         Height          =   390
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cmbTahun 
         Height          =   390
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Komputer"
         Height          =   270
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   1035
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2160
         TabIndex        =   14
         Top             =   1200
         Width           =   45
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   4080
         TabIndex        =   13
         Top             =   720
         Width           =   45
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   3120
         TabIndex        =   12
         Top             =   720
         Width           =   45
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Waktu Simpan"
         Height          =   270
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   930
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2160
         TabIndex        =   7
         Top             =   720
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Simpan"
         Height          =   270
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2160
         TabIndex        =   5
         Top             =   240
         Width           =   45
      End
   End
   Begin MSAdodcLib.Adodc AdodcCMBDataLalu 
      Height          =   375
      Left            =   360
      Top             =   6720
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.ComboBox cmbDataLalu 
      Height          =   390
      Index           =   0
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   65
      Top             =   1920
      Width           =   4335
   End
   Begin VB.ComboBox cmbDataLalu 
      Height          =   390
      Index           =   1
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   66
      Top             =   2880
      Width           =   4335
   End
   Begin VB.ComboBox cmbDataLalu 
      Height          =   390
      Index           =   2
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   67
      Top             =   3360
      Width           =   4335
   End
   Begin VB.ComboBox cmbDataLalu 
      Height          =   390
      Index           =   3
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   68
      Top             =   3840
      Width           =   4335
   End
   Begin VB.ComboBox cmbDataLalu 
      Height          =   390
      Index           =   4
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   69
      Top             =   4320
      Width           =   4335
   End
   Begin VB.ComboBox cmbDataLalu 
      Height          =   390
      Index           =   5
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   70
      Top             =   4800
      Width           =   4335
   End
   Begin MSAdodcLib.Adodc AdodcUpdateTerakhir 
      Height          =   375
      Left            =   360
      Top             =   7080
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
   Begin XPEngine.XPControl MesinXP 
      Left            =   0
      Top             =   0
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
End
Attribute VB_Name = "FormDataBaru"
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
    With AdodcJenisAkun
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * from tbJenisAkun"
        .Refresh
    End With
    With AdodcCMBDataLalu
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * from TabelPasswordManager"
        .Refresh
    End With
    With AdodcUpdateTerakhir
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * from TabelUpdateTerakhir"
        .Refresh
    End With
    Reset
    With cmbHari
        .Clear
        .AddItem "Senin", 0
        .AddItem "Selasa", 1
        .AddItem "Rabu", 2
        .AddItem "Kamis", 3
        .AddItem "Jumat", 4
        .AddItem "Sabtu", 5
        .AddItem "Minggu", 6
        .ListIndex = 0
    End With
    With cmbBulan
        .Clear
        .AddItem "01", 0
        .AddItem "02", 1
        .AddItem "03", 2
        .AddItem "04", 3
        .AddItem "05", 4
        .AddItem "06", 5
        .AddItem "07", 6
        .AddItem "08", 7
        .AddItem "09", 8
        .AddItem "10", 9
        .AddItem "11", 10
        .AddItem "12", 11
        .ListIndex = 0
    End With
    cmbTahun.Clear
    For Z = 1800 To 3000
        cmbTahun.AddItem Z
    Next
    cmbTahun.Text = Year(Date)
    AturOtomatisWaktuSimpan
    MasukkanDatabaseKeCMBJenisAkun
    With textNamaKomputer
        .Alignment = vbCenter
        .Text = GetComputerName
        .Locked = True
    End With
    If cmSimpan.Caption = "&Simpan" Then
        cmExport.Enabled = False
    Else
        cmExport.Enabled = True
    End If
    MasukkanDataKeCMBDataLalu
    If FormPengaturan.cekIzinkanPenambahanKategoriAkun.Value = Checked Then
        cmTambahKategori.Visible = True
    ElseIf FormPengaturan.cekIzinkanPenambahanKategoriAkun.Value = Unchecked Then
        cmTambahKategori.Visible = False
    End If
    MesinXP.StartEngine
End Sub
Sub Reset()
    For Each Objek In Me
        If TypeName(Objek) = "TextBox" Then
            With Objek
                .Text = ""
                .MaxLength = 254
            End With
        End If
    Next
End Sub
Sub AturOtomatisWaktuSimpan()
    Select Case Month(Date)
        Case Is = 1
            cmbBulan.ListIndex = 0
        Case Is = 2
            cmbBulan.ListIndex = 1
        Case Is = 3
            cmbBulan.ListIndex = 2
        Case Is = 4
            cmbBulan.ListIndex = 3
        Case Is = 5
            cmbBulan.ListIndex = 4
        Case Is = 6
            cmbBulan.ListIndex = 5
        Case Is = 7
            cmbBulan.ListIndex = 6
        Case Is = 8
            cmbBulan.ListIndex = 7
        Case Is = 9
            cmbBulan.ListIndex = 8
        Case Is = 10
            cmbBulan.ListIndex = 9
        Case Is = 11
            cmbBulan.ListIndex = 10
        Case Is = 12
            cmbBulan.ListIndex = 11
    End Select
    cmbTahun.Text = Year(Date)
    cmbTanggal.Text = Day(Date)
End Sub
Sub MasukkanDatabaseKeCMBJenisAkun()
    With cmbJenisAkun
        .Clear
        Do Until AdodcJenisAkun.Recordset.EOF
            cmbJenisAkun.AddItem AdodcJenisAkun.Recordset.Fields(0).Value
            AdodcJenisAkun.Recordset.MoveNext
        Loop
        AdodcJenisAkun.Refresh
        If FormPengaturan.cekIzinkanPenambahanKategoriAkun.Value = Checked Then
            .AddItem "------------------------------------------------------------------------------------------------------------------------------------------"
            .AddItem "[Tambah...]"
        End If
        .Text = AdodcJenisAkun.Recordset.Fields(0).Value
    End With
End Sub
Sub SimpanDataKeDatabase()
On Error GoTo HancurkanError
    Select Case cmSimpan.Caption
    Case Is = "&Simpan"
        With FormUtama
            .AdodcUtama.Recordset.AddNew
            .AdodcUtama.Recordset.Fields(0).Value = textPemilikAkun.Text
            .AdodcUtama.Recordset.Fields(1).Value = cmbJenisAkun.Text
            .AdodcUtama.Recordset.Fields(2).Value = textNamaAkun.Text
            .AdodcUtama.Recordset.Fields(3).Value = textUserName.Text
            .AdodcUtama.Recordset.Fields(4).Value = textEmail.Text
            .AdodcUtama.Recordset.Fields(5).Value = textPassword.Text
            .AdodcUtama.Recordset.Fields(6).Value = textAlamatWeb.Text
            .AdodcUtama.Recordset.Fields(7).Value = cmbHari.Text
            .AdodcUtama.Recordset.Fields(8).Value = cmbTanggal.Text
            .AdodcUtama.Recordset.Fields(9).Value = cmbBulan.Text
            .AdodcUtama.Recordset.Fields(10).Value = cmbTahun.Text
            .AdodcUtama.Recordset.Fields(11).Value = textJam.Text
            .AdodcUtama.Recordset.Fields(12).Value = textMenit.Text
            .AdodcUtama.Recordset.Fields(13).Value = textDetik.Text
            .AdodcUtama.Recordset.Fields(14).Value = textNamaKomputer.Text
            .AdodcUtama.Recordset.Update
            .AdodcUtama.Refresh
            .AturKontrol
            UpdateTekahirTabel
        End With
    Case Is = "&Update"
        With FormManage
            .AdodcUtama.Recordset.Delete
            .AdodcUtama.Recordset.AddNew
            .AdodcUtama.Recordset.Fields(0).Value = textPemilikAkun.Text
            .AdodcUtama.Recordset.Fields(1).Value = cmbJenisAkun.Text
            .AdodcUtama.Recordset.Fields(2).Value = textNamaAkun.Text
            .AdodcUtama.Recordset.Fields(3).Value = textUserName.Text
            .AdodcUtama.Recordset.Fields(4).Value = textEmail.Text
            .AdodcUtama.Recordset.Fields(5).Value = textPassword.Text
            .AdodcUtama.Recordset.Fields(6).Value = textAlamatWeb.Text
            .AdodcUtama.Recordset.Fields(7).Value = cmbHari.Text
            .AdodcUtama.Recordset.Fields(8).Value = cmbTanggal.Text
            .AdodcUtama.Recordset.Fields(9).Value = cmbBulan.Text
            .AdodcUtama.Recordset.Fields(10).Value = cmbTahun.Text
            .AdodcUtama.Recordset.Fields(11).Value = textJam.Text
            .AdodcUtama.Recordset.Fields(12).Value = textMenit.Text
            .AdodcUtama.Recordset.Fields(13).Value = textDetik.Text
            .AdodcUtama.Recordset.Fields(14).Value = textNamaKomputer.Text
            .AdodcUtama.Recordset.Update
            .AdodcUtama.Refresh
            .AturKontrol
            UpdateTekahirTabel
        End With
        FormUtama.AturKontrol
    End Select
    Reset
    If FormPengaturan.cekTutupFormINput.Value = Checked Then Unload Me
    cmBatal.Caption = "&Tutup"
Exit Sub
HancurkanError:
    PusatError
End Sub
Sub MasukkanDataKeCMBDataLalu()
    For Z = 0 To 5
        cmbDataLalu.Item(Z).Clear
    Next
    With cmbDataLalu
        Do Until AdodcCMBDataLalu.Recordset.EOF
            .Item(0).AddItem AdodcCMBDataLalu.Recordset.Fields(0).Value
            .Item(1).AddItem AdodcCMBDataLalu.Recordset.Fields(2).Value
            .Item(2).AddItem AdodcCMBDataLalu.Recordset.Fields(3).Value
            .Item(3).AddItem AdodcCMBDataLalu.Recordset.Fields(4).Value
            .Item(4).AddItem AdodcCMBDataLalu.Recordset.Fields(5).Value
            .Item(5).AddItem AdodcCMBDataLalu.Recordset.Fields(6).Value
            AdodcCMBDataLalu.Recordset.MoveNext
        Loop
    End With
End Sub
Sub UpdateTekahirTabel()
    If AdodcUpdateTerakhir.Recordset.RecordCount = 0 Then
        With AdodcUpdateTerakhir
            .Recordset.AddNew
            .Recordset.Fields(0).Value = "[" & Day(Date) & "/" & Month(Date) & "/" & Year(Date) & "] - [" & Time & "]"
            .Recordset.Update
            .Refresh
        End With
    Else
        With AdodcUpdateTerakhir
            .Recordset.Delete
            .Recordset.AddNew
            .Recordset.Fields(0).Value = "[" & Day(Date) & "/" & Month(Date) & "/" & Year(Date) & "] - [" & Time & "]"
            .Recordset.Update
            .Refresh
        End With
    End If
End Sub


Private Sub cmBatal_Click()
    Unload Me
End Sub

Private Sub cmbBulan_Click()
    cmbTanggal.Clear
    Select Case cmbBulan.ListIndex
        Case Is = 0, 2, 4, 6, 7, 9, 11
            For Z = 1 To 31
                With cmbTanggal
                    .AddItem Z
                End With
            Next
        Case Is = 1
            If Val(cmbTahun.Text) Mod 4 Then
                For Z = 1 To 29
                    With cmbTanggal
                        .AddItem Z
                    End With
                Next
            Else
                For Z = 1 To 28
                    With cmbTanggal
                        .AddItem Z
                    End With
                Next
            End If
        Case Is = 3, 5, 8, 10
            For Z = 1 To 30
                With cmbTanggal
                    .AddItem Z
                End With
            Next
    End Select
    cmbTanggal.Text = "1"
End Sub

Private Sub cmbDataLalu_Click(Index As Integer)
    Select Case Index
        Case Is = 0
            With textPemilikAkun
                .Text = cmbDataLalu.Item(0).Text
                .SetFocus
            End With
        Case Is = 1
            With textNamaAkun
                .Text = cmbDataLalu.Item(1).Text
                .SetFocus
            End With
        Case Is = 2
            With textUserName
                .Text = cmbDataLalu.Item(2).Text
                .SetFocus
            End With
        Case Is = 3
            With textEmail
                .Text = cmbDataLalu.Item(3).Text
                .SetFocus
            End With
        Case Is = 4
            With textPassword
                .Text = cmbDataLalu.Item(4).Text
                .SetFocus
            End With
        Case Is = 5
            With textAlamatWeb
                .Text = cmbDataLalu.Item(5).Text
                .SetFocus
            End With
    End Select
End Sub

Private Sub cmbJenisAkun_Click()
    If cmbJenisAkun.Text = "------------------------------------------------------------------------------------------------------------------------------------------" Then
        MasukkanDatabaseKeCMBJenisAkun
    ElseIf cmbJenisAkun.Text = "[Tambah...]" Then
        FormTambahKategori.Show vbModal, Me
    End If
End Sub

Private Sub cmbTahun_Click()
    cmbBulan_Click
End Sub

Private Sub cmCopy_Click(Index As Integer)
Clipboard.Clear
Select Case Index
    Case Is = 0
        With textTanggal
            .Text = cmbHari.Text & ", " & cmbTanggal.Text & " - " & cmbBulan.Text & " - " & cmbTahun.Text
            .SelStart = 0
            .SelLength = Len(textTanggal.Text)
            Clipboard.SetText .Text
        End With
    Case Is = 1
        With textWaktu
            .Text = textJam.Text & " : " & textMenit.Text & " : " & textDetik.Text
            .SelStart = 0
            .SelLength = Len(textWaktu.Text)
            Clipboard.SetText .Text
        End With
    Case Is = 2
        With textNamaKomputer
            .SelStart = 0
            .SelLength = Len(textNamaKomputer.Text)
            Clipboard.SetText .Text
        End With
    Case Is = 3
        If textPemilikAkun.Text = "" Then
            MsgBox "Tidak bisa dicopy, karena input masih kosong!", vbExclamation + vbOKOnly, ""
        Else
            With textPemilikAkun
                .SelStart = 0
                .SelLength = Len(textPemilikAkun.Text)
                Clipboard.SetText .Text
            End With
        End If
    Case Is = 4
        With textJenisAkun
            .Text = cmbJenisAkun.Text
            .SelStart = 0
            .SelLength = Len(textJenisAkun.Text)
            Clipboard.SetText .Text
        End With
    Case Is = 5
        If textNamaAkun.Text = "" Then
            MsgBox "Tidak bisa dicopy, karena input masih kosong!", vbExclamation + vbOKOnly, ""
        Else
            With textNamaAkun
                .SelStart = 0
                .SelLength = Len(textNamaAkun.Text)
                Clipboard.SetText .Text
            End With
        End If
    Case Is = 6
        If textUserName.Text = "" Then
            MsgBox "Tidak bisa dicopy, karena input masih kosong!", vbExclamation + vbOKOnly, ""
        Else
            With textUserName
                .SelStart = 0
                .SelLength = Len(textUserName.Text)
                Clipboard.SetText .Text
            End With
        End If
    Case Is = 7
        If textEmail.Text = "" Then
            MsgBox "Tidak bisa dicopy, karena input masih kosong!", vbExclamation + vbOKOnly, ""
        Else
            With textEmail
                .SelStart = 0
                .SelLength = Len(textEmail.Text)
                Clipboard.SetText .Text
            End With
        End If
    Case Is = 8
        If textPassword.Text = "" Then
            MsgBox "Tidak bisa dicopy, karena input masih kosong!", vbExclamation + vbOKOnly, ""
        Else
            With textPassword
                .SelStart = 0
                .SelLength = Len(textPassword.Text)
                Clipboard.SetText .Text
            End With
        End If
    Case Is = 9
        If textAlamatWeb.Text = "" Then
            MsgBox "Tidak bisa dicopy, karena input masih kosong!", vbExclamation + vbOKOnly, ""
        Else
            With textAlamatWeb
                .SelStart = 0
                .SelLength = Len(textAlamatWeb.Text)
                Clipboard.SetText .Text
            End With
        End If
    End Select
End Sub

Private Sub cmExport_Click()
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
    Print #iFile, "======================================================================================================"
    Print #iFile, "Tanggal Simpan       : " & cmbHari.Text & ", " & cmbTanggal & " - " & cmbBulan.Text & " - " & cmbTahun
    Print #iFile, "Waktu Simpan         : " & textJam.Text & " : " & textMenit.Text & " : " & textDetik.Text
    Print #iFile, "Nama Komputer        : " & textNamaKomputer.Text
    Print #iFile, "======================================================================================================"
    Print #iFile, "Pemilik Akun         : " & textPemilikAkun.Text
    Print #iFile, "Jenis Akun           : " & cmbJenisAkun.Text
    Print #iFile, "Nama Akun            : " & textNamaAkun.Text
    Print #iFile, "User Name            : " & textUserName.Text
    Print #iFile, "E-Mail               : " & textEmail.Text
    Print #iFile, "Password             : " & textPassword.Text
    Print #iFile, "Alamat Web           : " & textAlamatWeb.Text
    Print #iFile, "======================================================================================================"
    
    
    SaveFileFromTB = True
ErrorHandler:
    Close #iFile
End Sub

Private Sub cmHapus_Click(Index As Integer)
Select Case Index
    Case Is = 0
        With textPemilikAkun
            .Text = ""
            .SetFocus
        End With
    Case Is = 1
        MasukkanDatabaseKeCMBJenisAkun
    Case Is = 2
        With textNamaAkun
            .Text = ""
            .SetFocus
        End With
    Case Is = 3
        With textUserName
            .Text = ""
            .SetFocus
        End With
    Case Is = 4
        With textEmail
            .Text = ""
            .SetFocus
        End With
    Case Is = 5
        With textPassword
            .Text = ""
            .SetFocus
        End With
    Case Is = 6
        With textAlamatWeb
            .Text = ""
            .SetFocus
        End With
End Select
End Sub

Private Sub cmLebarkan_Click()
    Select Case cmLebarkan.Caption
    Case ">>>"
        Me.Width = 8010
        cmLebarkan.Caption = "<<<"
    Case "<<<"
        Me.Width = 7035
        cmLebarkan.Caption = ">>>"
    End Select
End Sub

Private Sub cmReset_Click()
    Reset
    AturKontrol
End Sub

Private Sub cmSetTanggaldanWaktu_Click()

End Sub

Private Sub cmSimpan_Click()
    If textPemilikAkun.Text = "" Then
        MsgBox "Silahkan isi nama pemilik akun!", vbExclamation + vbOKOnly, ""
        textPemilikAkun.SetFocus
    ElseIf textNamaAkun.Text = "" Then
        MsgBox "Silahkan isi nama akun!", vbExclamation + vbOKOnly, ""
        textNamaAkun.SetFocus
    ElseIf textUserName.Text = "" Then
        MsgBox "Silahkan isi username yang digunakan pada akun anda!", vbExclamation + vbOKOnly, ""
        textUserName.SetFocus
    ElseIf textEmail.Text = "" Then
        MsgBox "Silahkan isi alamat email yang digunakan pada akun anda!", vbExclamation + vbOKOnly, ""
        textEmail.SetFocus
    ElseIf textPassword.Text = "" Then
        MsgBox "Silahkan isi password yang anda gunakan untuk login di akun anda", vbExclamation + vbOKOnly, ""
        textPassword.SetFocus
    ElseIf textAlamatWeb.Text = "" Then
        MsgBox "Silahkan ini alamat web pada akun anda!", vbExclamation + vbOKOnly, ""
        textAlamatWeb.SetFocus
    Else
        If FormPengaturan.CekKonfirmasiSimpanEditData.Value = Checked Then
            Z = MsgBox("Apakah Anda yakin ingin menyimpan data ini?", vbQuestion + vbYesNo, "Konfirmasi")
            If Z = vbYes Then
                SimpanDataKeDatabase
            End If
        ElseIf FormPengaturan.CekKonfirmasiSimpanEditData.Value = Unchecked Then
            SimpanDataKeDatabase
        End If
    End If
End Sub

Private Sub cmTambahKategori_Click()
    FormTambahKategori.Show vbModal, Me
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub


Private Sub textAlamatWeb_DblClick()
    R = SendMessageLong(cmbDataLalu.Item(5).hWnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textEmail_DblClick()
    R = SendMessageLong(cmbDataLalu.Item(3).hWnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textNamaAkun_DblClick()
    R = SendMessageLong(cmbDataLalu.Item(1).hWnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textPassword_DblClick()
    R = SendMessageLong(cmbDataLalu.Item(4).hWnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textPemilikAkun_DblClick()
    R = SendMessageLong(cmbDataLalu.Item(0).hWnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textUserName_DblClick()
    R = SendMessageLong(cmbDataLalu.Item(2).hWnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub TimerWaktu_Timer()
    textJam.Text = Hour(Time)
    textMenit.Text = Minute(Time)
    textDetik.Text = Second(Time)
End Sub
