VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Begin VB.Form FormLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3945
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   3945
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.TextBox textNamaPengguna 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   1680
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox textPassword 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   1680
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton cmLogin 
         Caption         =   "&Login"
         Default         =   -1  'True
         Height          =   375
         Left            =   2640
         TabIndex        =   1
         Top             =   1200
         Width           =   975
      End
      Begin MSAdodcLib.Adodc AdodcLogin 
         Height          =   330
         Left            =   120
         Top             =   1320
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pengguna"
         Height          =   270
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1560
         TabIndex        =   6
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         Height          =   270
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1560
         TabIndex        =   4
         Top             =   720
         Width           =   45
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
Attribute VB_Name = "FormLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    NyambungUtama
    With AdodcLogin
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * from tbLogin"
        .Refresh
    End With
    For Each Objek In Me
        If TypeName(Objek) = "TextBox" Then
            With Objek
                .Text = ""
                .MaxLength = 254
            End With
        End If
    Next
    textPassword.PasswordChar = "*"
    Call CekProgram(FormLogin)
    MesinXP.StartEngine
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

Private Sub cmLogin_Click()
    If textNamaPengguna.Text = "" Then
        MsgBox "Silahkan isi nama pengguna Anda!", vbExclamation + vbOKOnly, ""
        textNamaPengguna.SetFocus
    ElseIf textPassword.Text = "" Then
        MsgBox "Silahkan isi password Anda!", vbExclamation + vbOKOnly, ""
        textPassword.SetFocus
    Else
        If RS.State = 1 Then RS.Close
            Kalimat = "select * from tbLogin where NamaPengguna= '" & textNamaPengguna.Text & "' And Password = '" & textPassword.Text & "'"
            RS.Open Kalimat, CN, 3, 3
            If Not RS.EOF Then
                Me.Hide
                FormUtama.Show vbModal, Me
            Else
                MsgBox "Maaf, input login yang Anda masukkan tidak benar!" & vbCrLf & _
                        "Silahkan diperiksa kembali!", vbCritical + vbOKOnly, "Error"
                textNamaPengguna.SetFocus
            End If
    End If
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

Private Sub Frame1_Click()
    textNamaPengguna.SetFocus
End Sub

