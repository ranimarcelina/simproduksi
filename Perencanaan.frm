VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Perencanaan 
   Caption         =   "Schedule"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   8160
      TabIndex        =   21
      Top             =   3360
      Width           =   2535
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   8160
      TabIndex        =   20
      Top             =   6240
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   8160
      TabIndex        =   19
      Top             =   5520
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   8160
      TabIndex        =   18
      Top             =   4800
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   8160
      TabIndex        =   17
      Top             =   4080
      Width           =   2535
   End
   Begin VB.CommandButton Command6 
      Caption         =   "SAVE"
      Height          =   615
      Left            =   13080
      TabIndex        =   12
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "+"
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
      Left            =   10920
      TabIndex        =   11
      Top             =   2640
      Width           =   615
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4800
      Top             =   7680
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
      Left            =   4800
      TabIndex        =   10
      Top             =   8160
      Width           =   11175
      _ExtentX        =   19711
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
      Left            =   7320
      TabIndex        =   9
      Top             =   7080
      Width           =   7335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "PRINT"
      Height          =   615
      Left            =   13080
      TabIndex        =   8
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "DELETE"
      Height          =   615
      Left            =   13080
      TabIndex        =   7
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "COUNT"
      Height          =   615
      Left            =   13080
      TabIndex        =   6
      Top             =   3480
      Width           =   1575
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
      Left            =   8160
      TabIndex        =   4
      Top             =   2640
      Width           =   2535
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
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "LATENESS"
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
      Left            =   5640
      TabIndex        =   16
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "DUE DATE"
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
      Left            =   5640
      TabIndex        =   15
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "FLOWTIME"
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
      Left            =   5640
      TabIndex        =   14
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "PROCESSING TIME"
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
      Left            =   5640
      TabIndex        =   13
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
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
      Left            =   5640
      TabIndex        =   5
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "NO ORDER"
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
      Left            =   5640
      TabIndex        =   3
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ID SCHEDULE"
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
      Left            =   5640
      TabIndex        =   2
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SCHEDULE"
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
      Left            =   9360
      TabIndex        =   1
      Top             =   1440
      Width           =   2175
   End
End
Attribute VB_Name = "Perencanaan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Menu.Show
    Me.Hide
End Sub

Private Sub Command2_Click()
Call IdOtomatis
End Sub
Private Sub IdOtomatis()
Dim RSJadwal As New ADODB.Recordset
Dim conn As New ADODB.Connection

Set conn = New ADODB.Connection
    conn.ConnectionString = _
    "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=simproduksi;Data Source=DESKTOP-KQT6V0C"
    conn.Open
    
If RSJadwal.State = adStateOpen Then RSJadwal.Close
    RSJadwal.Open "select * from jadwalproduksi where id_jadwalproduksi in (Select max(id_jadwalproduksi) from jadwalproduksi) order by id_jadwalproduksi desc", conn, adOpenKeyset
    
Dim IDJadwal As String * 6
Dim Hitung As Long

If RSJadwal.EOF Then
    IDJadwal = "SCH" + "001"
    Text1.Text = IDJadwal
Else
    Hitung = Right(RSJadwal!id_jadwalproduksi, 3) + 1
    IDJadwal = "SC" + Right("000" & Hitung, 3)
End If
    Text1.Text = IDJadwal
End Sub

Private Sub Command6_Click()

Dim conn As New ADODB.Connection

Set conn = New ADODB.Connection
    conn.ConnectionString = _
    "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=simproduksi;Data Source=DESKTOP-KQT6V0C"
    conn.Open
    
End Sub
