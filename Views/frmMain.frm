VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AI THONG MINH HON - MAIN"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9270
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
   ScaleHeight     =   6120
   ScaleWidth      =   9270
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDiem 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   405
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   930
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtDA 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   405
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   930
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdDoiCauHoi 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "Doi cau hoi"
      Height          =   615
      Left            =   6120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1680
      Width           =   2775
   End
   Begin VB.CommandButton cmdYKienKhanGia 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Hoi y kien khan gia"
      Height          =   615
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1680
      Width           =   2775
   End
   Begin VB.CommandButton cmdTroGiup50 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tro giup 50/50"
      Height          =   615
      Left            =   360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1680
      Width           =   2775
   End
   Begin VB.CommandButton cmdThoat 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Thoat chuong trinh"
      Height          =   615
      Left            =   360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label lblCauHoi 
      BackColor       =   &H8000000D&
      Caption         =   "Cau hoi :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lblDiem 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      Caption         =   "Diem: 00000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6840
      TabIndex        =   9
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblThoigian 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      Caption         =   "Thoi gian: 00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4440
      TabIndex        =   8
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label txtCauHoi 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Noi dung cau hoi"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   360
      TabIndex        =   6
      Top             =   3000
      Width           =   8415
   End
   Begin VB.Label txtD 
      Alignment       =   2  'Center
      Caption         =   "D"
      Height          =   735
      Left            =   4680
      TabIndex        =   5
      Top             =   4800
      Width           =   4095
   End
   Begin VB.Label txtC 
      Alignment       =   2  'Center
      Caption         =   "C"
      Height          =   735
      Left            =   360
      TabIndex        =   4
      Top             =   4800
      Width           =   3975
   End
   Begin VB.Label txtB 
      Alignment       =   2  'Center
      Caption         =   "B"
      Height          =   735
      Left            =   4680
      TabIndex        =   3
      Top             =   3840
      Width           =   4095
   End
   Begin VB.Label txtA 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "A"
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   3840
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "AI THONG MINH HON"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   9255
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
      TabIndex        =   0
      Top             =   5760
      Width           =   9255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Diem, Giay, I As Integer

Private Sub Form_Load()
    Diem = 0
    I = 1
    lblCauHoi.Caption = "Cau hoi " & I & " :"
    Call CSDL
    Call TimeIn
    
    'Tim kiem trong DB
    sql = "SELECT TOP 16 CauHoi.*, IIf([DoKho]='DE', 50, IIf([DoKho]='TB', 100, IIf([DoKho]='KHO', 150, 0))) AS Diem, * From CauHoi ORDER BY Rnd(Int(Now()*MaCauHoi)-Now()*MaCauHoi);"
    If rs.State = 1 Then rs.Close
'    rs.Open strSQL, DB, 3, 3
    Set rs = db.Execute(sql)
    
    Set txtCauHoi.DataSource = rs
    txtCauHoi.DataField = "NoiDung"
    
    Set txtA.DataSource = rs
    txtA.DataField = "A"
    
    Set txtB.DataSource = rs
    txtB.DataField = "B"
    
    Set txtC.DataSource = rs
    txtC.DataField = "C"
    
    Set txtD.DataSource = rs
    txtD.DataField = "D"
        
    Set txtDA.DataSource = rs
    txtDA.DataField = "DA"
            
    Set txtDiem.DataSource = rs
    txtDiem.DataField = "Diem"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Quit
End Sub

Private Sub cmdThoat_Click()
    If (MsgBox("Thoat chuong trinh?", vbYesNo) = vbYes) Then
        Call Logout
    End If
End Sub

Private Sub Timer_Timer()
    If Giay <= 10 Then
        lblThoigian.ForeColor = &HFF&
    End If
    
    If Giay > 0 Then
        Giay = Giay - 1
        lblThoigian.Caption = "Thoi gian: " & Right("00" & Giay, 2)
    Else
        Call TimeOut
        MsgBox ("Het thoi gian!")
        Call ErrOpt
    End If
End Sub

Private Sub TimeOut()
    Timer.Enabled = False
    txtA.Enabled = False
    txtB.Enabled = False
    txtC.Enabled = False
    txtD.Enabled = False
End Sub

Private Sub TimeIn()
    Giay = 30
    lblThoigian.Caption = "Thoi gian: " & Right("00" & Giay, 2)
    lblThoigian.ForeColor = &HFFFFFF
    Timer.Enabled = True
    txtA.Enabled = True
    txtB.Enabled = True
    txtC.Enabled = True
    txtD.Enabled = True
End Sub

Private Sub txtA_Click()
    If Trim(txtDA.Text) = "A" Then
        Diem = Diem + txtDiem.Text
        lblDiem.Caption = "Diem: " & Right("0000" & Diem, 4)
        Call TimeIn
        Call MoveNext
    Else
        Call ErrOpt
    End If
End Sub

Private Sub txtB_Click()
    If Trim(txtDA.Text) = "B" Then
        Diem = Diem + txtDiem.Text
        lblDiem.Caption = "Diem: " & Right("0000" & Diem, 4)
        Call TimeIn
        Call MoveNext
    Else
        Call ErrOpt
    End If
End Sub

Private Sub txtC_Click()
    If Trim(txtDA.Text) = "C" Then
        Diem = Diem + txtDiem.Text
        lblDiem.Caption = "Diem: " & Right("0000" & Diem, 4)
        Call TimeIn
        Call MoveNext
    Else
        Call ErrOpt
    End If
End Sub

Private Sub txtD_Click()
    If Trim(txtDA.Text) = "D" Then
        Diem = Diem + txtDiem.Text
        lblDiem.Caption = "Diem: " & Right("0000" & Diem, 4)
        Call TimeIn
        Call MoveNext
    Else
        Call ErrOpt
    End If
End Sub

Private Sub cmdDoiCauHoi_Click()
    cmdDoiCauHoi.Enabled = False
    Call MoveNext
End Sub

Private Sub cmdTroGiup50_Click()
    cmdTroGiup50.Enabled = False
End Sub

Private Sub cmdYKienKhanGia_Click()
    cmdYKienKhanGia.Enabled = False
    Dim S, A, b, C, d As Integer
    Randomize
    
    S = 100
    A = Int((S * Rnd) + 0)
    b = Int(((S - A) * Rnd) + 0)
    C = Int(((S - A - b) * Rnd) + 0)
    d = S - A - b - C

    MsgBox "Tro giup tu khan gia: " & vbCrLf & _
        "     A = " & A & "%" & vbCrLf & _
        "     B = " & b & "%" & vbCrLf & _
        "     C = " & C & "%" & vbCrLf & _
        "     D = " & d & "%"
End Sub

Private Sub MoveNext()
    If I < 15 Then
        I = I + 1
        lblCauHoi.Caption = "Cau hoi " & I & " :"
    End If

    If rs.AbsolutePosition < rs.RecordCount And I <= 15 Then
        rs.MoveNext
    Else
        Call TimeOut
    End If
End Sub

Private Sub ErrOpt()
    If (MsgBox("Diem cua ban la: " & Diem & vbCrLf & _
        "Ban co muon choi lai khong?", vbYesNo, "Ban da tra loi sai!") = vbYes) Then
        Call Form_Load
    Else
        Call Logout
    End If
End Sub

