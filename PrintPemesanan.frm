VERSION 5.00
Begin VB.Form PrintPemesanan 
   Caption         =   "Order Report"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   11760
      TabIndex        =   4
      Top             =   2880
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8760
      TabIndex        =   3
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
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
      Left            =   10320
      TabIndex        =   1
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PRINT"
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
      Left            =   10320
      TabIndex        =   0
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "UNTIL"
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
      Left            =   10680
      TabIndex        =   6
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "PERIODE"
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
      Left            =   7320
      TabIndex        =   5
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "ORDER REPORT"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      TabIndex        =   2
      Top             =   1200
      Width           =   3255
   End
End
Attribute VB_Name = "PrintPemesanan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    CrystalReport1.SelectionFormula = "{pemesanan.order_date} in date (" & Combo1 & ") to date (" & Combo2 & ")"
    CrystalReport1.ReportFileName = App.Path & "\LaporanPemesanan.Rpt"
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.RetrieveDataFiles
    CrystalReport1.Action = 1
End Sub
Private Sub Command2_Click()
    Me.Hide
    Pemesanan.Show
End Sub

Private Sub Form_Load()
Dim RSPemesanan As New ADODB.Recordset
Dim conn As New ADODB.Connection

Set conn = New ADODB.Connection
    conn.ConnectionString = _
    "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=simproduksi;Data Source=DESKTOP-KQT6V0C"
    conn.Open
    
    RSPemesanan.Open "Select distinct date from pemesanan", conn
    Combo1.Clear
    Combo2.Clear
    
    Do While Not RSPemesanan.EOF
        Combo1.AddItem Format(RSPemesanan!order_date, "YYYY,MM,DD")
        Combo2.AddItem Format(RSPemesanan!order_date, "YYYY,MM,DD")
        RSPemesanan.MoveNext
    Loop
End Sub
