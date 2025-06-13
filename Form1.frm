VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15615
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   15615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btn_agregar 
      Caption         =   "Agregar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   4
      Top             =   5400
      Width           =   1695
   End
   Begin MSComctlLib.ListView list_libros 
      Height          =   4935
      Left            =   3240
      TabIndex        =   3
      Top             =   240
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   8705
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton btn_leiste 
      Caption         =   "Ya leiste"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   2
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Height          =   6735
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      Begin VB.CommandButton btn_catalogo 
         Caption         =   "Catalogo MEGA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2415
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CargarLibros(filtroSQL As String)
    Dim rs As ADODB.Recordset
    Dim sql As String
    
    sql = "SELECT L.LibroID, L.Titulo, L.Autor, G.Nombre As Genero, L.Calificacion, L.PrestadoA" & _
        "FROM Libros L INNER JOIN Generos G ON L.GeneroID = G.GeneroID"
        
    If filtroSQL <> "" Then
        sql = sql & " WHERE " & filtroSQL
    End If
    
    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenStatic, adLockReadOnly
    
    list_libros.ListItems.Clear
    
    
    
End Sub


Private Sub btn_catalogo_Click()
    CargarLibros ""
End Sub

Private Sub Form_Load()
    Set conn = New ADODB.Connection
    conn.CursorLocation = adUseClient
    
    Dim connString As String
    connString = "Provider=SQLOLEDB.1;Data Source=Topete;Initial Catalog=LibreriaMega;Integrated Security=SSPI;"
    
    conn.Open connString
    
    With list_libros
        .View = lvwReport
        .GridLines = True
        .FullRowSelect = True
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Titulo", 1500
        .ColumnHeaders.Add , , "Autor", 1500
        .ColumnHeaders.Add , , "Genero", 1500
        .ColumnHeaders.Add , , "Calificacion", 1500
        .ColumnHeaders.Add , , "Prestado a", 1500
        
        End With
    
End Sub
