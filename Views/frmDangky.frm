VERSION 5.00
Begin VB.Form frmDangky 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AI THONG MINH HON - DANG KY"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdThoat 
      Caption         =   "THOAT"
      Height          =   615
      Left            =   3240
      TabIndex        =   8
      Top             =   4200
      Width           =   2775
   End
   Begin VB.CommandButton cmdHanhDong 
      Caption         =   "DANG NHAP"
      Height          =   615
      Left            =   360
      TabIndex        =   7
      Top             =   4200
      Width           =   2775
   End
   Begin VB.CommandButton cmdBatDau 
      Caption         =   "DANG KY NGAY"
      Height          =   615
      Left            =   360
      TabIndex        =   6
      Top             =   3360
      Width           =   5655
   End
   Begin VB.TextBox txtMatKhau 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Text            =   "Nhap mat khau"
      Top             =   2280
      Width           =   5655
   End
   Begin VB.TextBox txtTaiKhoan 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Text            =   "Nhap tai khoan cua ban"
      Top             =   1200
      Width           =   5655
   End
   Begin VB.Label lblMatKhau 
      BackColor       =   &H8000000D&
      Caption         =   "Mat khau :"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label lblTaiKhoan 
      BackColor       =   &H8000000D&
      Caption         =   "Ten dang nhap :"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "TUAN ANH @ TH22-12"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   5040
      Width           =   6375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "AI THONG MINH HON"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "frmDangky"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdHanhDong_Click()
    frmDangnhap.Show
    Me.Hide
End Sub

Private Sub cmdThoat_Click()
    Call Quit
End Sub

Private Sub Form_Load()
    Call CSDL
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Quit
End Sub

Private Sub cmdBatDau_Click()
    'Kiem tra dau vao trong
    If ((txtTaiKhoan.Text = "" Or txtTaiKhoan.Text = "Nhap tai khoan cua ban") _
        Or (txtMatKhau.Text = "" Or txtMatKhau.Text = "Nhap mat khau")) _
    Then
        MsgBox "Ten dang nhap hoac mat khau khong duoc de trong!"
        Exit Sub
    End If
    
    'Them tai khoan vao DB
    sql = "INSERT INTO NguoiChoi (TenDangNhap, MatKhau) VALUES('" & txtTaiKhoan.Text & "', '" & txtMatKhau.Text & "');"
    Set rs = New ADODB.Recordset
'        rs.Open sql, db, 3, 3
    Set rs = db.Execute(sql)
    
    'Tim kiem trong DB
    sql = "SELECT TenDangNhap, isAdmin FROM NguoiChoi WHERE ((TenDangNhap Like '" & txtTaiKhoan.Text & "') AND (MatKhau Like '" & txtMatKhau.Text & "'));"
    Set rs = New ADODB.Recordset
    rs.Open sql, db, adOpenStatic, adLockOptimistic
        
    'Kiem tra trung khop du lieu
    If (rs.RecordCount > 0) Then
        MsgBox "Dang ky thanh cong, hay dang nhap!"
        frmDangnhap.Show
        Me.Hide
    'Nguoc lai in ra thong bao dang nhap ko thanh cong
    Else
        MsgBox "Dang ky khong thanh cong, vui long thu lai!"
        Exit Sub
    End If
End Sub

'Xu ly tai khoan
Private Sub txtTaiKhoan_GotFocus()
    If (txtTaiKhoan.Text) = "Nhap tai khoan cua ban" Then
        txtTaiKhoan.Text = ""
        txtTaiKhoan.ForeColor = vbBlack
    End If
End Sub

Private Sub txtTaiKhoan_LostFocus()
    If (txtTaiKhoan.Text) = "" Then
        txtTaiKhoan.Text = "Nhap tai khoan cua ban"
        txtTaiKhoan.ForeColor = &H8000000A
    End If
End Sub

Private Sub txtTaiKhoan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdBatDau_Click
    End If
End Sub

'Xu ly mat khau
Private Sub txtMatKhau_GotFocus()
    If (txtMatKhau.Text) = "Nhap mat khau" Then
        txtMatKhau.Text = ""
        txtMatKhau.ForeColor = vbBlack
        txtMatKhau.PasswordChar = "*"
    End If
End Sub

Private Sub txtMatKhau_LostFocus()
    If (txtMatKhau.Text) = "" Then
        txtMatKhau.Text = "Nhap mat khau"
        txtMatKhau.ForeColor = &H8000000A
        txtMatKhau.PasswordChar = ""
    End If
End Sub

Private Sub txtMatKhau_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdBatDau_Click
    End If
End Sub


