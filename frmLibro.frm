VERSION 5.00
Begin VB.Form frmLibro 
   Caption         =   "Agrega un Libro"
   ClientHeight    =   7080
   ClientLeft      =   1470
   ClientTop       =   3045
   ClientWidth     =   7275
   LinkTopic       =   "Form2"
   ScaleHeight     =   7080
   ScaleWidth      =   7275
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   15
      Top             =   5880
      Width           =   2895
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   14
      Top             =   5880
      Width           =   3495
   End
   Begin VB.TextBox txtPrestadoA 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   600
      TabIndex        =   13
      Text            =   "Prestado a"
      Top             =   4800
      Width           =   5775
   End
   Begin VB.CheckBox chkPrestado 
      Caption         =   "Prestado Actualmente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   4440
      Width           =   3615
   End
   Begin VB.CheckBox chkRecomendado 
      Caption         =   "Recomendado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   10
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CheckBox chkPorLeer 
      Caption         =   "Quiero Leer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   3240
      Width           =   1935
   End
   Begin VB.CheckBox chkLeido 
      Caption         =   "Ya leido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox txtCalificacion 
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   2160
      Width           =   615
   End
   Begin VB.ComboBox cboGenero 
      Height          =   315
      Left            =   1920
      TabIndex        =   4
      Top             =   1680
      Width           =   4695
   End
   Begin VB.TextBox txtAutor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   960
      Width           =   4695
   End
   Begin VB.TextBox txtTitulo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   4695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Prestamo"
      Height          =   1335
      Left            =   480
      TabIndex        =   11
      Top             =   4200
      Width           =   6495
   End
   Begin VB.Label Label4 
      Caption         =   "Calificacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Genero"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Autor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Titulo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmLibro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False









Private Sub Combo1_Change()

End Sub

Private Sub Command1_Click()

End Sub

Private Sub chkLeido_Click()
    If chkLeido.Value = 1 Then
        chkPorLeer.Value = 0
        txtCalificacion.Enabled = True
    Else
        txtCalificacion.Enabled = False
    End If
End Sub

Private Sub chkPorLeer_Click()
    If chkPorLeer.Value = 1 Then
        chkLeido.Value = 0
    End If
End Sub

Private Sub chkPrestado_Click()
    If chkPrestado.Value = 1 Then
        txtPrestadoA.Enabled = True
    Else
        txtPrestadoA.Enabled = False
        txtPrestadoA.Text = ""
    End If
    
End Sub

Private Sub Form_Load()
    Dim rsG As ADODB.Recordset
    Set rsG = New ADODB.Recordset
    rsG.Open "SELECT GeneroID, Nombre FROM Generos ORDER BY Nombre", conn, adOpenStatic, adLockReadOnly
    cboGenero.Clear
    Do Until rsG.EOF
        cboGenero.AddItem rsG!Nombre
        cboGenero.ItemData(cboGenero.NewIndex) = rsG!GeneroID
        rsG.MoveNext
    Loop
    
    rsG.Close: Set rsG = Nothing
    
    If EditandoID = 0 Then
        ' Modo agregar, limpiar campos
        txtTitulo.Text = ""
        txtAutor = ""
        cboGenero.ListIndex = -1 ' no hay nada seleccionado
        txtCalificacion = ""
        chkLeido.Value = 0
        txtPrestadoA.Enabled = False
        Me.Caption = "Agregar Libro"
        
    Else
        
    
    End If
    
End Sub
