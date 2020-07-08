VERSION 5.00
Begin VB.Form Login 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Login"
   ClientHeight    =   5115
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11115
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000017&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   19680
      Top             =   960
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      TabIndex        =   6
      Top             =   5280
      Width           =   1815
   End
   Begin VB.TextBox TPASSWORD 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   9720
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox TLOGIN 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   4
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      TabIndex        =   3
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   18120
      TabIndex        =   7
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   2
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   1
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PT. PAPYRUS SAKTI PAPER MILL"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6840
      TabIndex        =   0
      Top             =   1560
      Width           =   6735
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim rs As New ADODB.Recordset

If TLOGIN = "" Or TPASSWORD = "" Then
    MsgBox "Invalid Login", vbCritical + vbOKOnly, "Information"
    Exit Sub
End If

Set rs = JalankanSQL("select * from pegawai where nama_pegawai = '" & Trim(TLOGIN.Text & "'"))

If rs.RecordCount = 0 Then
    MsgBox "Username Not Found!", vbCritical + vbOKOnly, "Information"
    TLOGIN.SetFocus
    TLOGIN.Text = ""
    TPASSWORD.Text = ""
    Exit Sub
End If

Set rs = JalankanSQL("select * from pegawai where password = '" & Trim(TPASSWORD.Text & "'"))

If rs.RecordCount = 0 Then
    MsgBox "Incorrect Password!", vbCritical + vbOKOnly, "Information"
    TPASSWORD.SetFocus
    TLOGIN.Text = ""
    TPASSWORD.Text = ""
    Exit Sub
Else
    Me.Visible = False
    Menu.StatusBar1.Panels(1) = rs!id_pegawai
    Menu.StatusBar1.Panels(2) = rs!nama_pegawai
    Menu.StatusBar1.Panels(3) = rs!Status
    MenuStaff.StatusBar1.Panels(1) = rs!id_pegawai
    MenuStaff.StatusBar1.Panels(2) = rs!nama_pegawai
    MenuStaff.StatusBar1.Panels(3) = rs!Status
    
    If Menu.StatusBar1.Panels(3) <> "Admin" Then
        Menu.mnmaster.Enabled = False
        Me.Hide
        MenuStaff.Show
    Else
        Menu.Show
        Menu.mnmaster.Enabled = True
    End If
End If
End Sub
Private Sub Command2_Click()
    TLOGIN.SetFocus
    TLOGIN.Text = ""
    TPASSWORD.Text = ""
End Sub

Private Sub Timer1_Timer()
Label4.Caption = Format(Now, "dd MMM yyyy hh:mm:ss")
End Sub
