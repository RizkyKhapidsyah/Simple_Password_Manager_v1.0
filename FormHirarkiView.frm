VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Begin VB.Form FormHirarkiView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hirarki View"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7290
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormHirarkiView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   840
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmTutup 
      Caption         =   "&Tutup"
      Height          =   975
      Left            =   6000
      Picture         =   "FormHirarkiView.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmCopySemua 
      Caption         =   "&Copy Semua"
      Height          =   975
      Left            =   1440
      Picture         =   "FormHirarkiView.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmExport 
      Caption         =   "&Export"
      Height          =   975
      Left            =   120
      Picture         =   "FormHirarkiView.frx":06D6
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox textTersembunyi 
      Height          =   1335
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   52
      Text            =   "FormHirarkiView.frx":1298
      Top             =   8400
      Width           =   5415
   End
   Begin VB.CommandButton cmCopy 
      Caption         =   "Copy"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   6600
      TabIndex        =   51
      Top             =   5400
      Width           =   555
   End
   Begin VB.CommandButton cmCopy 
      Caption         =   "Copy"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   6600
      TabIndex        =   50
      Top             =   4920
      Width           =   555
   End
   Begin VB.CommandButton cmCopy 
      Caption         =   "Copy"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   6600
      TabIndex        =   49
      Top             =   4440
      Width           =   555
   End
   Begin VB.CommandButton cmCopy 
      Caption         =   "Copy"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   6600
      TabIndex        =   48
      Top             =   3960
      Width           =   555
   End
   Begin VB.CommandButton cmCopy 
      Caption         =   "Copy"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   6600
      TabIndex        =   47
      Top             =   3480
      Width           =   555
   End
   Begin VB.CommandButton cmCopy 
      Caption         =   "Copy"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   6600
      TabIndex        =   46
      Top             =   3000
      Width           =   555
   End
   Begin VB.CommandButton cmCopy 
      Caption         =   "Copy"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   6600
      TabIndex        =   45
      Top             =   2520
      Width           =   555
   End
   Begin VB.CommandButton cmCopy 
      Caption         =   "Copy"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   6600
      TabIndex        =   44
      Top             =   2040
      Width           =   555
   End
   Begin VB.CommandButton cmCopy 
      Caption         =   "Copy"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   6600
      TabIndex        =   43
      Top             =   1560
      Width           =   555
   End
   Begin VB.CommandButton cmCopy 
      Caption         =   "Copy"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   6600
      TabIndex        =   42
      Top             =   1080
      Width           =   555
   End
   Begin MSAdodcLib.Adodc AdodcUtama 
      Height          =   330
      Left            =   120
      Top             =   8160
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
   Begin VB.Frame Frame2 
      Height          =   5055
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   6375
      Begin VB.TextBox textNamaKomputer 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   1920
         TabIndex        =   41
         Text            =   "Text1"
         Top             =   4560
         Width           =   4335
      End
      Begin VB.TextBox textDetikSimpan 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   3840
         TabIndex        =   38
         Text            =   "Text1"
         Top             =   4080
         Width           =   855
      End
      Begin VB.TextBox textMenitSimpan 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   2880
         TabIndex        =   37
         Text            =   "Text1"
         Top             =   4080
         Width           =   855
      End
      Begin VB.TextBox textJamSimpan 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   1920
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   4080
         Width           =   855
      End
      Begin VB.TextBox textTahunSimpan 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   4920
         TabIndex        =   33
         Text            =   "Text1"
         Top             =   3600
         Width           =   1335
      End
      Begin VB.TextBox textBulanSimpan 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   4080
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   3600
         Width           =   735
      End
      Begin VB.TextBox textTanggalSimpan 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   3240
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   3600
         Width           =   735
      End
      Begin VB.TextBox textHariSimpan 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   1920
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   3600
         Width           =   1215
      End
      Begin VB.TextBox textAlamatWeb 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   1920
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   3120
         Width           =   4335
      End
      Begin VB.TextBox textPassword 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   1920
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   2640
         Width           =   4335
      End
      Begin VB.TextBox textEmail 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   1920
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   2160
         Width           =   4335
      End
      Begin VB.TextBox textUserName 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   1920
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   1680
         Width           =   4335
      End
      Begin VB.TextBox textNamaAkun 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   1920
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1200
         Width           =   4335
      End
      Begin VB.TextBox textJenisAkun 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   1920
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   720
         Width           =   4335
      End
      Begin VB.TextBox textPemilikAkun 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   1920
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1800
         TabIndex        =   40
         Top             =   4560
         Width           =   45
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Komputer"
         Height          =   270
         Left            =   -120
         TabIndex        =   39
         Top             =   4560
         Width           =   1770
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1800
         TabIndex        =   35
         Top             =   4080
         Width           =   45
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Waktu Simpan"
         Height          =   270
         Left            =   -120
         TabIndex        =   34
         Top             =   4080
         Width           =   1770
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1800
         TabIndex        =   29
         Top             =   3600
         Width           =   45
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Simpan"
         Height          =   270
         Left            =   -120
         TabIndex        =   28
         Top             =   3600
         Width           =   1770
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1800
         TabIndex        =   26
         Top             =   3120
         Width           =   45
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat Web"
         Height          =   270
         Left            =   -120
         TabIndex        =   25
         Top             =   3120
         Width           =   1770
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1800
         TabIndex        =   23
         Top             =   2640
         Width           =   45
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         Height          =   270
         Left            =   -120
         TabIndex        =   22
         Top             =   2640
         Width           =   1770
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1800
         TabIndex        =   20
         Top             =   2160
         Width           =   45
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         Height          =   270
         Left            =   -120
         TabIndex        =   19
         Top             =   2160
         Width           =   1770
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1800
         TabIndex        =   17
         Top             =   1680
         Width           =   45
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         Height          =   270
         Left            =   -120
         TabIndex        =   16
         Top             =   1680
         Width           =   1770
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1800
         TabIndex        =   14
         Top             =   1200
         Width           =   45
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Akun"
         Height          =   270
         Left            =   -120
         TabIndex        =   13
         Top             =   1200
         Width           =   1770
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1800
         TabIndex        =   11
         Top             =   720
         Width           =   45
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Akun"
         Height          =   270
         Left            =   -120
         TabIndex        =   10
         Top             =   720
         Width           =   1770
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1800
         TabIndex        =   8
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pemilik Akun"
         Height          =   270
         Left            =   -120
         TabIndex        =   7
         Top             =   240
         Width           =   1770
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin VB.CommandButton cmAkhir 
         Caption         =   ">>"
         Height          =   495
         Left            =   5640
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmSelanjutnya 
         Caption         =   ">"
         Height          =   495
         Left            =   5040
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox TextData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   1320
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   240
         Width           =   3720
      End
      Begin VB.CommandButton cmSebelumnya 
         Caption         =   "<"
         Height          =   495
         Left            =   720
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmAwal 
         Caption         =   "<<"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
   End
   Begin XPEngine.XPControl MesinXP 
      Left            =   0
      Top             =   0
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
End
Attribute VB_Name = "FormHirarkiView"
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
        .RecordSource = "Select * from TabelPasswordManager"
        .Refresh
    End With
    For Each Objek In Me
        If TypeName(Objek) = "TextBox" Then
            With Objek
                .Text = ""
                .Locked = True
            End With
        End If
    Next
    MasukkanDatabaseKeTexBox
    MesinXP.StartEngine
End Sub
Sub MasukkanDatabaseKeTexBox()
On Error GoTo HancurkanError
    With Me
        .TextData.Text = "Nama Akun : " & AdodcUtama.Recordset.Fields(2).Value
        .textPemilikAkun.Text = AdodcUtama.Recordset.Fields(0).Value
        .textJenisAkun.Text = AdodcUtama.Recordset.Fields(1).Value
        .textNamaAkun.Text = AdodcUtama.Recordset.Fields(2).Value
        .textUserName.Text = AdodcUtama.Recordset.Fields(3).Value
        .textEmail.Text = AdodcUtama.Recordset.Fields(4).Value
        .textPassword.Text = AdodcUtama.Recordset.Fields(5).Value
        .textAlamatWeb.Text = AdodcUtama.Recordset.Fields(6).Value
        .textHariSimpan.Text = AdodcUtama.Recordset.Fields(7).Value
        .textTanggalSimpan.Text = AdodcUtama.Recordset.Fields(8).Value
        .textBulanSimpan.Text = AdodcUtama.Recordset.Fields(9).Value
        .textTahunSimpan.Text = AdodcUtama.Recordset.Fields(10).Value
        .textJamSimpan.Text = AdodcUtama.Recordset.Fields(11).Value
        .textMenitSimpan.Text = AdodcUtama.Recordset.Fields(12).Value & "'"
        .textDetikSimpan.Text = AdodcUtama.Recordset.Fields(13).Value & "''"
        .textNamaKomputer.Text = AdodcUtama.Recordset.Fields(14).Value
    End With
Exit Sub
HancurkanError:
    PusatError
End Sub

Private Sub cmAkhir_Click()
    AdodcUtama.Recordset.MoveLast
    MasukkanDatabaseKeTexBox
End Sub

Private Sub cmAwal_Click()
    AdodcUtama.Recordset.MoveFirst
    MasukkanDatabaseKeTexBox
End Sub

Private Sub cmCopy_Click(Index As Integer)
Clipboard.Clear
Select Case Index
    Case Is = 0
        With textPemilikAkun
            .SetFocus
            .SelStart = 0
            .SelLength = Len(textPemilikAkun.Text)
            Clipboard.SetText .Text
        End With
    Case Is = 1
        With textJenisAkun
            .SetFocus
            .SelStart = 0
            .SelLength = Len(textJenisAkun.Text)
            Clipboard.SetText .Text
        End With
    Case Is = 2
        With textNamaAkun
            .SetFocus
            .SelStart = 0
            .SelLength = Len(textNamaAkun.Text)
            Clipboard.SetText .Text
        End With
    Case Is = 3
        With textUserName
            .SetFocus
            .SelStart = 0
            .SelLength = Len(textUserName.Text)
            Clipboard.SetText .Text
        End With
    Case Is = 4
        With textEmail
            .SetFocus
            .SelStart = 0
            .SelLength = Len(textEmail.Text)
            Clipboard.SetText .Text
        End With
    Case Is = 5
        With textPassword
            .SetFocus
            .SelStart = 0
            .SelLength = Len(textPassword.Text)
            Clipboard.SetText .Text
        End With
    Case Is = 6
        With textAlamatWeb
            .SetFocus
            .SelStart = 0
            .SelLength = Len(textAlamatWeb.Text)
            Clipboard.SetText .Text
        End With
    Case Is = 7
        With textTersembunyi
            .Text = textHariSimpan.Text & ", " & textTanggalSimpan.Text & " - " & textBulanSimpan.Text & " - " & textTahunSimpan.Text
            .SelStart = 0
            .SelLength = Len(textTersembunyi.Text)
            Clipboard.SetText .Text
        End With
        textHariSimpan.SetFocus
    Case Is = 8
        With textTersembunyi
            .Text = textJamSimpan.Text & " : " & textMenitSimpan.Text & " : " & textDetikSimpan.Text
            .SelStart = 0
            .SelLength = Len(textTersembunyi.Text)
            Clipboard.SetText .Text
        End With
        textJamSimpan.SetFocus
    Case Is = 9
        With textNamaKomputer
            .SetFocus
            .SelStart = 0
            .SelLength = Len(textNamaKomputer.Text)
            Clipboard.SetText .Text
        End With
End Select
End Sub

Private Sub cmCopySemua_Click()
With textTersembunyi
    .Text = "Pemilik Akun     : " & textPemilikAkun.Text & vbCrLf & _
            "Jenis Akun       : " & textJenisAkun.Text & vbCrLf & _
            "Nama Akun        : " & textNamaAkun.Text & vbCrLf & _
            "User Name        : " & textUserName.Text & vbCrLf & _
            "Email            : " & textEmail.Text & vbCrLf & _
            "Password         : " & textPassword.Text & vbCrLf & _
            "Alamat Web       : " & textAlamatWeb.Text & vbCrLf & _
            "Tanggal Simpan   : " & textHariSimpan.Text & ", " & textTanggalSimpan.Text & " - " & textBulanSimpan.Text & " - " & textTahunSimpan.Text & vbCrLf & _
            "Waktu Simpan     : " & textJamSimpan.Text & " : " & textMenitSimpan.Text & " : " & textDetikSimpan.Text & vbCrLf & _
            "Nama Komputer    : " & textPemilikAkun.Text
    .SelStart = 0
    .SelLength = Len(.Text)
    Clipboard.SetText .Text
End With
MsgBox "Semua data berhasil disalin!", vbInformation + vbOKOnly, ""
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
    Print #iFile, "Pemilik Akun     : " & textPemilikAkun.Text
    Print #iFile, "Jenis Akun       : " & textJenisAkun.Text
    Print #iFile, "Nama Akun        : " & textNamaAkun.Text
    Print #iFile, "User Name        : " & textUserName.Text
    Print #iFile, "Email            : " & textEmail.Text
    Print #iFile, "Password         : " & textPassword.Text
    Print #iFile, "Alamat Web       : " & textAlamatWeb.Text
    Print #iFile, "Tanggal Simpan   : " & textHariSimpan.Text & ", " & textTanggalSimpan.Text & " - " & textBulanSimpan.Text & " - " & textTahunSimpan.Text
    Print #iFile, "Waktu Simpan     : " & textJamSimpan.Text & " : " & textMenitSimpan.Text & " : " & textDetikSimpan.Text
    Print #iFile, "Nama Komputer    : " & textPemilikAkun.Text
    Print #iFile, "======================================================================================================"
    
    
    SaveFileFromTB = True
ErrorHandler:
    Close #iFile
End Sub

Private Sub cmSebelumnya_Click()
    AdodcUtama.Recordset.MovePrevious
    If AdodcUtama.Recordset.BOF = True Then AdodcUtama.Recordset.MoveLast
    MasukkanDatabaseKeTexBox
End Sub

Private Sub cmSelanjutnya_Click()
    AdodcUtama.Recordset.MoveNext
    If AdodcUtama.Recordset.EOF = True Then AdodcUtama.Recordset.MoveFirst
    MasukkanDatabaseKeTexBox
End Sub

Private Sub cmTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    AturKontrol
    MasukkanDatabaseKeTexBox
End Sub
