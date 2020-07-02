VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Menu 
   Caption         =   "Menu Utama"
   ClientHeight    =   5100
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   11775
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   4725
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command6 
      Caption         =   "LOGOUT"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   18360
      TabIndex        =   6
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "PERENCANAAN PRODUKSI"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14160
      TabIndex        =   5
      Top             =   3360
      Width           =   3375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "PEGAWAI"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11280
      TabIndex        =   4
      Top             =   3360
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "PRODUK"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8280
      TabIndex        =   3
      Top             =   3360
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "MESIN PRODUKSI"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   2
      Top             =   3360
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PEMESANAN"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   1
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Label Label1 
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
      Height          =   855
      Left            =   7200
      TabIndex        =   0
      Top             =   1320
      Width           =   6615
   End
   Begin VB.Menu mnmaster 
      Caption         =   "Master"
      Begin VB.Menu mnpemesanan 
         Caption         =   "Pemesanan"
      End
      Begin VB.Menu mnmesinproduksi 
         Caption         =   "Mesin Produksi"
      End
      Begin VB.Menu mnproduk 
         Caption         =   "Produk"
      End
      Begin VB.Menu mnpegawai 
         Caption         =   "Pegawai"
      End
      Begin VB.Menu mnperencanaan 
         Caption         =   "Perencanaan Produksi"
      End
   End
   Begin VB.Menu mnlaporan 
      Caption         =   "Laporan"
      Begin VB.Menu mnlappemesanan 
         Caption         =   "Laporan Pemesanan"
      End
   End
   Begin VB.Menu mnexit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Pemesanan.Show
    Me.Hide
End Sub
Private Sub Command2_Click()
    MesinProduksi.Show
    Me.Hide
End Sub
Private Sub Command3_Click()
    Produk.Show
    Me.Hide
End Sub
Private Sub Command4_Click()
    Pegawai.Show
    Me.Hide
End Sub
Private Sub Command5_Click()
    Perencanaan.Show
    Me.Hide
End Sub
Private Sub Command6_Click()
Menu.Hide
Login.Show
Login.TLOGIN.Text = ""
Login.TPASSWORD.Text = ""
Login.TLOGIN.SetFocus
End Sub
Private Sub mnexit_Click()
    Pesan = MsgBox("Close Application?", vbQuestion + vbYesNo, "Confirmation")
    If Pesan = vbYes Then End
End Sub
Private Sub mnmesinproduksi_Click()
    MesinProduksi.Show
    Me.Hide
End Sub

Private Sub mnpegawai_Click()
    Pegawai.Show
    Me.Hide
End Sub

Private Sub mnpemesanan_Click()
    Pemesanan.Show
    Me.Hide
End Sub

Private Sub mnperencanaan_Click()
    Perencanaan.Show
    Me.Hide
End Sub

Private Sub mnproduk_Click()
    Produk.Show
    Me.Hide
End Sub
