VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Begin VB.Form FormSizeDatagrid 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Size Datagrid"
   ClientHeight    =   1125
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3345
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormSizeDatagrid.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   3345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.TextBox TextValueBaris 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   2400
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   360
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSComctlLib.Slider SliderBAris 
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
      End
      Begin VB.Label LabelValueBaris 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   270
         Left            =   2520
         TabIndex        =   2
         Top             =   420
         Width           =   375
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
Attribute VB_Name = "FormSizeDatagrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Program : Simple Password Manager v1.0
'Source Code by Rizky Khafitsyah
'kunjungi http://rikymetalist.blogspot.com


Sub AturKontrol()
    For Each Objek In Me
        If TypeName(Objek) = "Slider" Then
            Objek.TickStyle = 3
        ElseIf TypeName(Objek) = "TextBox" Then
            Objek.MaxLength = 4
        End If
    Next
    With Me
        .SliderBAris.Min = 5
        .SliderBAris.Max = 2000
        .SliderBAris.Value = FormManage.DataGrid1.RowHeight
        .LabelValueBaris.Caption = SliderBAris.Value
    End With
    AmbilPengaturan
    MesinXP.StartEngine
End Sub
Sub SimpanPengaturan()
    SaveSetting App.Title, Me.Caption, Me.SliderBAris.Name, Me.SliderBAris.Value
    SaveSetting App.Title, Me.Caption, Me.LabelValueBaris.Name, Me.LabelValueBaris.Caption
    SaveSetting App.Title, Me.Caption, Me.TextValueBaris.Name, Me.TextValueBaris.Text
End Sub
Sub AmbilPengaturan()
    Me.LabelValueBaris.Caption = GetSetting(App.Title, Me.Caption, Me.LabelValueBaris.Name, Me.LabelValueBaris.Caption)
    Me.TextValueBaris.Text = GetSetting(App.Title, Me.Caption, Me.TextValueBaris.Name, Me.TextValueBaris.Text)
    Me.SliderBAris.Value = GetSetting(App.Title, Me.Caption, Me.SliderBAris.Name, Me.SliderBAris.Value)
End Sub


Private Sub Form_Load()
    AturKontrol
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SimpanPengaturan
End Sub

Private Sub LabelValueBaris_DblClick()
    With TextValueBaris
        .Text = LabelValueBaris.Caption
        .Visible = True
        .SetFocus
    End With
    LabelValueBaris.Visible = False
    FormManage.DataGrid1.RowHeight = SliderBAris.Value
End Sub

Private Sub SliderBAris_Scroll()
    LabelValueBaris.Caption = SliderBAris.Value
    TextValueBaris.Text = SliderBAris.Value
    If SliderBAris.Value > 15 Then
        FormManage.DataGrid1.RowHeight = SliderBAris.Value
    End If
End Sub

Private Sub TextValueBaris_Change()
    If Val(TextValueBaris.Text) > 2000 Then
        With TextValueBaris
            .Text = 2000
            .SetFocus
        End With
    Else
        SliderBAris.Value = Val(TextValueBaris.Text)
        LabelValueBaris.Caption = Val(TextValueBaris.Text)
        If SliderBAris.Value > 15 Then
            FormManage.DataGrid1.RowHeight = SliderBAris.Value
        End If
    End If
End Sub

Private Sub TextValueBaris_DblClick()
    If Val(TextValueBaris.Text) > 2000 Then
        MsgBox "Maaf, batas ukuran hanya sampai 2000", vbExclamation + vbOKOnly, ""
        TextValueBaris.SetFocus
    ElseIf TextValueBaris.Text = "" Then
        MsgBox "Nilai tidak boleh kosong", vbExclamation + vbOKOnly, ""
        TextValueBaris.SetFocus
    ElseIf Val(TextValueBaris.Text) = 0 Then
        With TextValueBaris
            .Text = 315
            .SetFocus
        End With
    Else
        SliderBAris.Value = Val(TextValueBaris.Text)
        LabelValueBaris.Caption = Val(TextValueBaris.Text)
        With LabelValueBaris
            .Caption = TextValueBaris.Text
            .Visible = True
        End With
        TextValueBaris.Visible = False
        If SliderBAris.Value > 15 Then
            FormManage.DataGrid1.RowHeight = SliderBAris.Value
        End If
    End If
End Sub

Private Sub TextValueBaris_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(TextValueBaris.Text) > 2000 Then
            MsgBox "Maaf, batas ukuran hanya sampai 2000", vbExclamation + vbOKOnly, ""
            TextValueBaris.SetFocus
        ElseIf Val(TextValueBaris.Text) = 0 Then
            With TextValueBaris
                .Text = 315
                .SetFocus
            End With
        ElseIf TextValueBaris.Text = "" Then
            MsgBox "Nilai tidak boleh kosong", vbExclamation + vbOKOnly, ""
            TextValueBaris.SetFocus
        Else
            SliderBAris.Value = Val(TextValueBaris.Text)
            LabelValueBaris.Caption = Val(TextValueBaris.Text)
            With LabelValueBaris
                .Caption = TextValueBaris.Text
                .Visible = True
            End With
            TextValueBaris.Visible = False
        If SliderBAris.Value > 15 Then
            FormManage.DataGrid1.RowHeight = SliderBAris.Value
        End If
        End If
    ElseIf Not ((KeyAscii >= 48) And (KeyAscii <= 57) Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub
