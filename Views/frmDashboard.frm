VERSION 5.00
Begin VB.Form frmDashboard 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AI THONG MINH HON - DASHBOARD"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10455
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
   ScaleHeight     =   6000
   ScaleWidth      =   10455
   StartUpPosition =   2  'CenterScreen
   Tag             =   "4"
   Begin VB.ComboBox txtDoKho 
      Height          =   405
      ItemData        =   "frmDashboard.frx":0000
      Left            =   1800
      List            =   "frmDashboard.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Tag             =   "3"
      Top             =   1440
      Width           =   6495
   End
   Begin VB.CommandButton cmdSua 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sua"
      Height          =   495
      Left            =   8520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      Tag             =   "2"
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdXoa 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Xoa"
      Height          =   495
      Left            =   8520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      Tag             =   "2"
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdLuu 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Luu"
      Height          =   495
      Left            =   8520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      Tag             =   "3"
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdHuy 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Huy"
      Height          =   495
      Left            =   8520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      Tag             =   "3"
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmdDau 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dau"
      Height          =   495
      Left            =   1800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      Tag             =   "2"
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmdTruoc 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Truoc"
      Height          =   495
      Left            =   3480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      Tag             =   "2"
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmdSau 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sau"
      Height          =   495
      Left            =   5160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Tag             =   "2"
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmdCuoi 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cuoi"
      Height          =   495
      Left            =   6840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      Tag             =   "2"
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmdThoat 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Thoat"
      Height          =   495
      Left            =   8520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      Tag             =   "2"
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmdThem 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Them"
      Height          =   495
      Left            =   8520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   12
      Tag             =   "2"
      Top             =   720
      Width           =   1455
   End
   Begin VB.OptionButton optB 
      BackColor       =   &H8000000D&
      Caption         =   "B"
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
      Height          =   495
      Left            =   840
      TabIndex        =   11
      Tag             =   "4"
      Top             =   2760
      Width           =   735
   End
   Begin VB.OptionButton optC 
      BackColor       =   &H8000000D&
      Caption         =   "C"
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
      Left            =   840
      TabIndex        =   10
      Tag             =   "4"
      Top             =   3480
      Width           =   735
   End
   Begin VB.OptionButton optD 
      BackColor       =   &H8000000D&
      Caption         =   "D"
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
      Left            =   840
      TabIndex        =   9
      Tag             =   "4"
      Top             =   4200
      Width           =   735
   End
   Begin VB.OptionButton optA 
      BackColor       =   &H8000000D&
      Caption         =   "A"
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
      Left            =   840
      TabIndex        =   8
      Tag             =   "4"
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox txtA 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Tag             =   "1"
      Top             =   2040
      Width           =   6495
   End
   Begin VB.TextBox txtB 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Tag             =   "1"
      Top             =   2760
      Width           =   6495
   End
   Begin VB.TextBox txtC 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Tag             =   "1"
      Top             =   3480
      Width           =   6495
   End
   Begin VB.TextBox txtD 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Tag             =   "1"
      Top             =   4200
      Width           =   6495
   End
   Begin VB.TextBox txtDA 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Tag             =   "1"
      Top             =   4920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtCauHoi 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Tag             =   "1"
      Top             =   720
      Width           =   6495
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000D&
      Caption         =   "Do kho :"
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
      Height          =   495
      Left            =   480
      TabIndex        =   24
      Top             =   1440
      Width           =   1215
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
      Left            =   -120
      TabIndex        =   23
      Top             =   5640
      Width           =   10575
   End
   Begin VB.Label Label2 
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
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   1215
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
      TabIndex        =   0
      Top             =   120
      Width           =   10455
   End
End
Attribute VB_Name = "frmDashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CT As Control

Private Sub Form_Load()
    Call CSDL
    Call Khoa(True)
    
    'Tim kiem trong DB
    sql = "SELECT * FROM CauHoi"
    Set rs = New ADODB.Recordset
    rs.Open sql, db, adOpenStatic, adLockOptimistic, adCmdText
'    Set rs = db.Execute(sql)

'    MsgBox rs!DoKho
'    Exit Sub
    
    If (rs.RecordCount > 0) Then
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
        
        Set txtDoKho.DataSource = rs
        txtDoKho.DataField = "DoKho"
    End If
    
    If rs.AbsolutePosition = 1 Then
        cmdDau.Enabled = False
        cmdTruoc.Enabled = False
    End If
    
    Call CheckDA
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Quit
End Sub

Private Sub cmdCuoi_Click()
    Call CheckDA
    
    cmdDau.Enabled = True
    cmdTruoc.Enabled = True
    cmdCuoi.Enabled = True
    cmdSau.Enabled = True
    rs.MoveLast
    
    If rs.AbsolutePosition = rs.RecordCount Then
        cmdCuoi.Enabled = False
        cmdSau.Enabled = False
    End If
End Sub

Private Sub cmdDau_Click()
    cmdCuoi.Enabled = True
    cmdSau.Enabled = True
    cmdDau.Enabled = True
    cmdTruoc.Enabled = True
    rs.MoveFirst
    
    If rs.AbsolutePosition = 1 Then
        cmdDau.Enabled = False
        cmdTruoc.Enabled = False
    End If
    
    Call CheckDA
End Sub

Private Sub cmdHuy_Click()
    rs.CancelUpdate
    rs.Resync
    Call Khoa(True)
End Sub

Private Sub cmdLuu_Click()
    rs.Update
    Call Khoa(True)
    
    If txtCauHoi.Text = "" Or txtA.Text = "" Or txtB.Text = "" _
        Or txtC.Text = "" Or txtD.Text = "" Or txtDA.Text = "" Then
        MsgBox "Thong tin khong the de trong.", vbInfomation, "Loi!"
        txtCauHoi.SetFocus
        Exit Sub
    End If
        
    If rs.AbsolutePosition = 1 Then
        cmdDau.Enabled = False
        cmdTruoc.Enabled = False
    End If
    
    If rs.AbsolutePosition = rs.RecordCount Then
        cmdCuoi.Enabled = False
        cmdSau.Enabled = False
    End If
End Sub

Private Sub cmdSau_Click()
    cmdDau.Enabled = True
    cmdTruoc.Enabled = True

    If rs.AbsolutePosition < rs.RecordCount Then rs.MoveNext
    
    If rs.AbsolutePosition = rs.RecordCount Then
        cmdCuoi.Enabled = False
        cmdSau.Enabled = False
    End If
    
    Call CheckDA
End Sub

Private Sub cmdSua_Click()
    Call Khoa(False)
    txtCauHoi.SetFocus
End Sub

Private Sub cmdThem_Click()
    rs.AddNew
    Call Khoa(False)
    Call Reset
    txtCauHoi.SetFocus
End Sub

Private Sub cmdThoat_Click()
    Call Quit
    frmDangnhap.Show
End Sub

Private Sub cmdTruoc_Click()
    cmdCuoi.Enabled = True
    cmdSau.Enabled = True
    
    If rs.AbsolutePosition > 1 Then rs.MovePrevious
    
    If rs.AbsolutePosition = 1 Then
        cmdDau.Enabled = False
        cmdTruoc.Enabled = False
    End If
    
    Call CheckDA
End Sub

Private Sub cmdXoa_Click()
    If MsgBox("Ban co chac muon xoa?", vbYesNo + vbQuestion, "Xoa") = vbYes Then
        rs.Delete
        
        Dim VT As Long
        VT = rs.AbsolutePosition
        If rs.RecordCount > 0 Then
            If VT < rs.RecordCount Then
                rs.AbsolutePosition = VT
            Else
                rs.AbsolutePosition = VT - 1
            End If
        End If
    End If
End Sub

Private Sub optA_Click()
    txtDA.Text = "A"
End Sub

Private Sub optB_Click()
    txtDA.Text = "B"
End Sub

Private Sub optC_Click()
    txtDA.Text = "C"
End Sub

Private Sub optD_Click()
    txtDA.Text = "D"
End Sub

Private Sub Khoa(A As Boolean)
    For Each CT In Me.Controls
        If CT.Tag = 1 Then CT.Enabled = Not A
        If CT.Tag = 2 Then CT.Enabled = A
        If CT.Tag = 3 Then CT.Enabled = Not A
        If CT.Tag = 4 Then CT.Enabled = Not A
    Next
End Sub

Private Sub Reset()
    For Each CT In Me.Controls
        If CT.Tag = 1 Then CT.Text = ""
        If CT.Tag = 4 Then CT.Value = False
    Next
End Sub

Private Sub CheckDA()
    If txtDA.Text = "A" Then
        optA.Value = True
    End If
    
    If txtDA.Text = "B" Then
        optB.Value = True
    End If
    
    If txtDA.Text = "C" Then
        optC.Value = True
    End If
    
    If txtDA.Text = "D" Then
        optD.Value = True
    End If
End Sub
