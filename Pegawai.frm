VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Pegawai 
   Caption         =   "Employee"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7320
      Top             =   6600
      Width           =   3255
      _ExtentX        =   5741
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
      RecordSource    =   "pegawai"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2175
      Left            =   7320
      TabIndex        =   13
      Top             =   7080
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   3836
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
   Begin VB.CommandButton Command4 
      Caption         =   "DELETE"
      Height          =   615
      Left            =   12480
      TabIndex        =   12
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "UPDATE"
      Height          =   615
      Left            =   12480
      TabIndex        =   11
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SAVE"
      Height          =   615
      Left            =   12480
      TabIndex        =   10
      Top             =   3360
      Width           =   1695
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
      Left            =   8280
      TabIndex        =   9
      Top             =   5760
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
      Left            =   8280
      TabIndex        =   8
      Top             =   4920
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
      Left            =   8280
      TabIndex        =   7
      Top             =   4080
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
      Left            =   8280
      TabIndex        =   6
      Top             =   3240
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BACK"
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
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "STATUS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   5
      Top             =   5760
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   4
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "EMPLOYEE NAME"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   3
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "EMPLOYEE ID"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   2
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "EMPLOYEE"
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
      Left            =   8880
      TabIndex        =   1
      Top             =   1560
      Width           =   2295
   End
End
Attribute VB_Name = "Pegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Menu.Show
    Me.Hide
End Sub
Sub Form_Load()
Dim RSPegawai As New ADODB.Recordset
Dim conn As New ADODB.Connection

Set conn = New ADODB.Connection
    conn.ConnectionString = _
    "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=simproduksi;Data Source=DESKTOP-KQT6V0C"
    conn.Open
    
RSPegawai.CursorLocation = adUseClient
RSPegawai.Open " Select * from pegawai", conn, 3, 1
 
With DataGrid1
 Set .DataSource = RSPegawai
 .Refresh
 
End With
End Sub
Private Sub Command2_Click()
Dim SQLTambah As String
Dim conn As New ADODB.Connection

Set conn = New ADODB.Connection
    conn.ConnectionString = _
    "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=simproduksi;Data Source=DESKTOP-KQT6V0C"
    conn.Open

    If Text2 = "" Or Text3 = "" Or Text4 = "" Then
        MsgBox "Data Not Completed", vbCritical + vbOKOnly, "Information"
    Exit Sub
    End If
    
    SQLTambah = "INSERT INTO dbo.pegawai(id_pegawai, nama_pegawai, password, status)" & "values ('" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "')"
    conn.Execute SQLTambah
    MsgBox " Data Saved ", vbInformation, "Messages"
    Text1.SetFocus
    Text2.SetFocus
    Text3.SetFocus
    Text4.SetFocus
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    
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
    
    UserName = Text2
    Text2 = Replace(Text2, "'", "''")
    
    If Text2 = "" Or Text3 = "" Or Text4 = "" Then
        MsgBox "Data Not Completed", vbCritical + vbOKOnly, "Information"
    Exit Sub
    End If
    
    SQLEdit = "Update pegawai Set nama_pegawai = '" & Text2 & "',   password = '" & Text3 & "', status = '" & Text4 & "' where id_pegawai ='" & Text1 & "'"
    conn.Execute SQLEdit
    MsgBox " Data Updated ", vbInformation, "Messages"
    Text1.SetFocus
    Text2.SetFocus
    Text3.SetFocus
    Text4.SetFocus
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    
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
            conn.Execute "Delete From pegawai where id_pegawai = '" & Text1 & "'"
            MsgBox " Data Deleted ", vbInformation, "Messages"
            Text1.SetFocus
            Text1.Text = ""
        End If
    End If
Call Form_Load
conn.Close
End Sub

