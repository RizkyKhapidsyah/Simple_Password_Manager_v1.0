Attribute VB_Name = "Module1"
Option Explicit

'Program : Simple Password Manager v1.0
'Source Code by Rizky Khafitsyah
'kunjungi http://rikymetalist.blogspot.com

'fungsi API untuk mendeteksi nama komputer
Public Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
'fungsi api yang dipakai untuk membuka combobox tanpa mengkliknya
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg _
As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'nah, ini adalah konstanta (sebuah variabel yg telah ditetapkan nilainya)
Public Const CB_SHOWDROPDOWN = &H14F


'disini semua variabel yg kita gunakan di progam, kita jadikan public atau global agar bisa dibaca oleh semua program
Public CN As New ADODB.Connection
Public RS As New ADODB.Recordset
Global LI As ListItem
Public Kalimat As String
Public Objek As Control
Public Z As Integer
Public R As Long

'nah ini bagian untuk mengkoneksikan database utama
Public Sub NyambungUtama()
If CN.State = adStateOpen Then CN.Close
    CN.CursorLocation = adUseClient
    CN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\RikyPasswordManager.rdb;Persist Security Info=False"
End Sub
 
 'fungsi buatan untuk mengambil nama komputer
Public Function GetComputerName() As String
Dim sResult As String * 255
    GetComputerNameA sResult, 255
    GetComputerName = Left$(sResult, InStr(sResult, Chr$(0)) - 1)
End Function

'bagian untuk mengatur default format penyimpanan file
Public Sub AturDefaultFormat()
Select Case FormPengaturan.cmbDefaultSimpan.ListIndex
    Case Is = 0
        FormDataBaru.CommonDialog1.FilterIndex = 1
        FormHirarkiView.CommonDialog1.FilterIndex = 1
        FormManage.CommonDialog1.FilterIndex = 1
    Case Is = 1
        FormDataBaru.CommonDialog1.FilterIndex = 2
        FormHirarkiView.CommonDialog1.FilterIndex = 2
        FormManage.CommonDialog1.FilterIndex = 2
    Case Is = 2
        FormDataBaru.CommonDialog1.FilterIndex = 3
        FormHirarkiView.CommonDialog1.FilterIndex = 3
        FormManage.CommonDialog1.FilterIndex = 3
    Case Is = 3
        FormDataBaru.CommonDialog1.FilterIndex = 4
        FormHirarkiView.CommonDialog1.FilterIndex = 4
        FormManage.CommonDialog1.FilterIndex = 4
    Case Is = 4
        FormDataBaru.CommonDialog1.FilterIndex = 5
        FormHirarkiView.CommonDialog1.FilterIndex = 5
        FormManage.CommonDialog1.FilterIndex = 5
    Case Is = 5
        FormDataBaru.CommonDialog1.FilterIndex = 6
        FormHirarkiView.CommonDialog1.FilterIndex = 6
        FormManage.CommonDialog1.FilterIndex = 6
    Case Is = 6
        FormDataBaru.CommonDialog1.FilterIndex = 7
        FormHirarkiView.CommonDialog1.FilterIndex = 7
        FormManage.CommonDialog1.FilterIndex = 7
End Select
End Sub

'penanganan error
Public Sub PusatError()
    MsgBox "Maaf, terdapat kesalahan pada system", vbCritical + vbOKOnly, "Error"
End Sub
'penanganan error yg saya buat masih terlalu sederhana karena buatnya buru.buru, silahkan dikembangkan lagi
