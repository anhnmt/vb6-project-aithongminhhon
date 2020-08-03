Attribute VB_Name = "modConfig"
'17103282 - NguyenManhTuanAnh - TH22.12
Option Explicit
Public db As New ADODB.Connection
Public rs As New ADODB.Recordset
Public sql As String

Public Sub CSDL()
    If db.State = 1 Then db.Close
    Set db = New ADODB.Connection
    
    db.CursorLocation = adUseClient
    db.Provider = "Microsoft.Jet.Oledb.4.0"
    db.ConnectionString = "Data Source=" & App.Path & "\DB_GAME - Copy.MDB"
    db.Open
End Sub

Sub Quit()
    Dim frm As Form
    
    For Each frm In Forms
        Unload frm
        Set frm = Nothing
    Next
    
    If rs.State = 1 Then rs.Close
    If db.State = 1 Then db.Close
End Sub

Sub Logout()
    Call Quit
    frmDangnhap.Show
End Sub
