VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAutor 
   BackColor       =   &H00EBEBED&
   Caption         =   "Autores"
   ClientHeight    =   8370
   ClientLeft      =   3045
   ClientTop       =   1920
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   Picture         =   "frmAutor.frx":0000
   ScaleHeight     =   8370
   ScaleWidth      =   10110
   Begin VB.CommandButton cmdButton 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   5
      Left            =   8760
      TabIndex        =   13
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Eliminar"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   7440
      TabIndex        =   12
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Editar"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   5880
      TabIndex        =   11
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4320
      TabIndex        =   10
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2760
      TabIndex        =   9
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Registrar"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1200
      TabIndex        =   8
      Top             =   4200
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid dgrAutores 
      Bindings        =   "frmAutor.frx":4BC5D
      Height          =   2175
      Left            =   600
      TabIndex        =   7
      Top             =   4920
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   3836
      _Version        =   393216
      AllowUpdate     =   -1  'True
      Enabled         =   0   'False
      HeadLines       =   1
      RowHeight       =   24
      FormatLocked    =   -1  'True
      AllowDelete     =   -1  'True
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
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Id"
         Caption         =   "Id"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Nombre"
         Caption         =   "Nombre"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "País de origen"
         Caption         =   "País de origen"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Fecha de nacimiento"
         Caption         =   "Fecha de nacimiento"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         ScrollBars      =   2
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   1275,024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2429,858
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoAutores 
      Height          =   375
      Left            =   3960
      Top             =   3480
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
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
      Connect         =   $"frmAutor.frx":4BC76
      OLEDBString     =   $"frmAutor.frx":4BDE1
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Autores"
      Caption         =   "Anterior     Siguiente"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtInfo 
      Alignment       =   2  'Center
      DataField       =   "Fecha de nacimiento"
      DataSource      =   "adoAutores"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   3960
      TabIndex        =   6
      Top             =   2760
      Width           =   3615
   End
   Begin VB.TextBox txtInfo 
      Alignment       =   2  'Center
      DataField       =   "País de origen"
      DataSource      =   "adoAutores"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3960
      TabIndex        =   5
      Top             =   2040
      Width           =   3615
   End
   Begin VB.TextBox txtInfo 
      Alignment       =   2  'Center
      DataField       =   "Nombre"
      DataSource      =   "adoAutores"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   3960
      TabIndex        =   4
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Autores"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "Fecha de nacimiento"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   600
      TabIndex        =   2
      Top             =   2760
      Width           =   3135
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "País de origen"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   1320
      Width           =   1695
   End
End
Attribute VB_Name = "frmAutor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdButton_Click(Index As Integer)

Select Case Index

Case 0 'Registrar
 
MsgBox "Ingrese la información y haga clic en guardar"
adoAutores.Recordset.AddNew
    For Index = 0 To 2
        txtInfo(Index).Enabled = True
        txtInfo(0).SetFocus
      Next

Case 1 'Guardar
adoAutores.Recordset.Update
dgrAutores.Enabled = False
For Index = 0 To 2
        txtInfo(Index).Enabled = False
      Next
 
Case 2 'Buscar
Dim Buscar As String, Criterio As String
 Buscar = InputBox("Escriba el nombre del autor que desea buscar", "Búsqueda por nombre", vbQuestion)
  If Buscar = "" Then Exit Sub
    Criterio = "Nombre like '*" & Buscar & "*'"
    adoAutores.Recordset.MoveNext
   If Not adoAutores.Recordset.EOF Then
      adoAutores.Recordset.Find Criterio
  End If
 If adoAutores.Recordset.EOF Then
    adoAutores.Recordset.MoveFirst
    adoAutores.Recordset.Find Criterio
        If adoAutores.Recordset.EOF Then
      adoAutores.Recordset.MoveLast
Respuesta = MsgBox("Autor no registrado", vbCritical)
        End If
   End If

Case 3 'Editar
      For Index = 0 To 2
        txtInfo(Index).Enabled = True
        txtInfo(0).SetFocus
      Next
    dgrAutores.DataChanged = False
    
Case 4 'Eliminar
Dim Confirmacion As Integer
Confirmacion = MsgBox("¿Está seguro que desea eliminar los datos?", vbYesNoCancel + vbQuestion + vbDefaultButton2, "Eliminar")
   If Confirmacion = vbYes Then
    adoAutores.Recordset.Delete
    MsgBox ("Datos eliminados")
     adoAutores.Recordset.MoveNext
If adoAutores.Recordset.EOF Then
   adoAutores.Recordset.MoveLast
   End If
     Else
       Exit Sub
End If

Case 5 'Cerrar
Unload Me
Load frmLibrary

 End Select
End Sub
