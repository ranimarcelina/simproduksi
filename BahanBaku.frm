VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form BahanBaku 
   Caption         =   "Raw Material"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   17040
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2655
      Left            =   4080
      TabIndex        =   23
      Top             =   7680
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   4683
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   24
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4080
      Top             =   7200
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=simproduksi;Data Source=DESKTOP-KQT6V0C"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=simproduksi;Data Source=DESKTOP-KQT6V0C"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "bahanbaku"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      Caption         =   "PRINT"
      Height          =   615
      Left            =   15480
      TabIndex        =   22
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "PREVIEW"
      Height          =   615
      Left            =   13800
      TabIndex        =   21
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "DELETE"
      Height          =   615
      Left            =   13800
      TabIndex        =   20
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "UPDATE"
      Height          =   615
      Left            =   13800
      TabIndex        =   19
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SAVE"
      Height          =   615
      Left            =   13800
      TabIndex        =   18
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox Text7 
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
      Left            =   11280
      TabIndex        =   17
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox Text6 
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
      Left            =   11280
      TabIndex        =   16
      Top             =   2880
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6240
      TabIndex        =   15
      Top             =   6480
      Width           =   9135
   End
   Begin VB.TextBox Text5 
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
      Left            =   6240
      TabIndex        =   14
      Top             =   5760
      Width           =   2895
   End
   Begin VB.TextBox Text4 
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
      Left            =   6240
      TabIndex        =   13
      Top             =   5040
      Width           =   2895
   End
   Begin VB.TextBox Text3 
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
      Left            =   6240
      TabIndex        =   12
      Top             =   4320
      Width           =   2895
   End
   Begin VB.TextBox Text2 
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
      Left            =   6240
      TabIndex        =   11
      Top             =   3600
      Width           =   2895
   End
   Begin VB.TextBox Text1 
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
      Left            =   6240
      TabIndex        =   10
      Top             =   2880
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BACK"
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
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "ENDING BALANCE"
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
      Left            =   9600
      TabIndex        =   9
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "RETUR"
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
      Left            =   9600
      TabIndex        =   8
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "SEARCH"
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
      Left            =   4080
      TabIndex        =   7
      Top             =   6480
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "USAGE"
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
      Left            =   4080
      TabIndex        =   6
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "RECEIPT"
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
      Left            =   4080
      TabIndex        =   5
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "BEGINNING BALANCE"
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
      Left            =   4080
      TabIndex        =   4
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ITEM"
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
      Left            =   4080
      TabIndex        =   3
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ITEM CODE"
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
      Left            =   4080
      TabIndex        =   2
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "RAW MATERIAL"
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
      Left            =   8880
      TabIndex        =   1
      Top             =   1440
      Width           =   3255
   End
End
Attribute VB_Name = "BahanBaku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
Dim RSBahanBaku As New ADODB.Recordset
Dim conn As New ADODB.Connection

Set conn = New ADODB.Connection
    conn.ConnectionString = _
    "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=simproduksi;Data Source=DESKTOP-KQT6V0C"
    conn.Open

RSBahanBaku.CursorLocation = adUseClient
RSBahanBaku.Open "Select * from bahanbaku where nama_barang like '%" & Combo1 & "%'", conn
If Not RSBahanBaku.EOF Then
    With RSBahanBaku
        With DataGrid1
            Set .DataSource = RSBahanBaku
                .Refresh
        End With
    End With
Else
    MsgBox "Data Not Found", vbCritical + vbOKOnly, "Information"
End If
End Sub

Private Sub Command1_Click()
    Menu.Show
    Me.Hide
End Sub

Private Sub Command2_Click()
Dim SQLTambah As String
Dim RSBahanBaku As New ADODB.Recordset
Dim conn As New ADODB.Connection

Set conn = New ADODB.Connection
    conn.ConnectionString = _
    "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=simproduksi;Data Source=DESKTOP-KQT6V0C"
    conn.Open

    If Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Then
        MsgBox "Data Not Completed", vbCritical + vbOKOnly, "Information"
    Exit Sub
    End If

    SQLTambah = "INSERT INTO dbo.bahanbaku(kode_barang, nama_barang, saldo_awal, penerimaan, pemakaian, retur, saldo_akhir)" & "values ('" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & Text5.Text & "','" & Text6.Text & "','" & Text7.Text & "')"
    conn.Execute SQLTambah
    MsgBox " Data Saved ", vbInformation, "Messages"
    Text1.SetFocus
    Text2.SetFocus
    Text3.SetFocus
    Text4.SetFocus
    Text5.SetFocus
    Text6.SetFocus
    Text7.SetFocus
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    
Call Form_Load
conn.Close
End Sub

Private Sub Command3_Click()
Dim SQLEdit As String
Dim conn As New ADODB.Connection

Set conn = New ADODB.Connection
    conn.ConnectionString = _
    "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=simproduksi;Data Source=DESKTOP-KQT6V0C"
    conn.Open
    
    nama_barang = Text2
    Text2 = Replace(Text2, "'", "''")
    
    If Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Then
        MsgBox "Data Not Completed", vbCritical + vbOKOnly, "Information"
    Exit Sub
    End If
    
    SQLEdit = "Update bahanbaku Set nama_barang = '" & Text2 & "', saldo_awal ='" & Text3 & "', penerimaan ='" & Text4 & "', pemakaian ='" & Text5 & "', retur = '" & Text6 & "', saldo_akhir = '" & "' where kode_barang ='" & Text1 & "'"
    conn.Execute SQLEdit
    MsgBox " Data Updated ", vbInformation, "Messages"
    Text1.SetFocus
    Text2.SetFocus
    Text3.SetFocus
    Text4.SetFocus
    Text5.SetFocus
    Text6.SetFocus
    Text7.SetFocus
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    
Call Form_Load
conn.Close
End Sub

Private Sub Command4_Click()
Dim SQLHapus As String
Dim conn As New ADODB.Connection

Set conn = New ADODB.Connection
    conn.ConnectionString = _
    "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=simproduksi;Data Source=DESKTOP-KQT6V0C"
    conn.Open
    
    If Text1 = "" Then
        MsgBox "Data Not Found", vbCritical + vbOKOnly, "Information"
    Else
        If MsgBox("Data Will be Deleted?", vbQuestion + vbOKCancel, "Confirmation") = vbOK Then
            conn.Execute "Delete From bahanbaku where kode_barang = '" & Text1 & "'"
            MsgBox " Data Deleted ", vbInformation, "Messages"
            Text1.SetFocus
            Text1.Text = ""
        End If
    End If
Call Form_Load
conn.Close
End Sub

Private Sub Command5_Click()
    CrystalReport1.ReportFileName = App.Path & "\ReportRawMaterial.Rpt"
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.RetrieveDataFiles
    CrystalReport1.Action = 0
End Sub

Private Sub Command6_Click()
    CrystalReport1.ReportFileName = App.Path + "\ReportRawMaterial.Rpt"
    CrystalReport1.Destination = crptToPrinter
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.PrintReport
End Sub

Sub Form_Load()
Dim RSBahanBaku As New ADODB.Recordset
Dim conn As New ADODB.Connection

Set conn = New ADODB.Connection
    conn.ConnectionString = _
    "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=simproduksi;Data Source=DESKTOP-KQT6V0C"
    conn.Open
    
RSBahanBaku.CursorLocation = adUseClient
RSBahanBaku.Open " Select * from bahanbaku", conn, 3, 1
 
With DataGrid1
 Set .DataSource = RSBahanBaku
 .Refresh
 
End With
End Sub

