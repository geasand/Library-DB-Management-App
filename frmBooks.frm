VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBooks 
   BackColor       =   &H00004080&
   Caption         =   "Libros"
   ClientHeight    =   7755
   ClientLeft      =   3405
   ClientTop       =   1920
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   Picture         =   "frmBooks.frx":0000
   ScaleHeight     =   7755
   ScaleWidth      =   9510
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
      Height          =   495
      Index           =   5
      Left            =   8160
      TabIndex        =   19
      Top             =   7200
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
      Height          =   495
      Index           =   4
      Left            =   7560
      TabIndex        =   18
      Top             =   4680
      Width           =   1215
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
      Height          =   495
      Index           =   3
      Left            =   6120
      TabIndex        =   17
      Top             =   4680
      Width           =   1215
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
      Height          =   495
      Index           =   2
      Left            =   4560
      TabIndex        =   16
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Guardar cambios"
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
      Left            =   2160
      TabIndex        =   15
      Top             =   4680
      Width           =   2175
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
      Height          =   495
      Index           =   0
      Left            =   600
      TabIndex        =   14
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox txtInfo 
      Alignment       =   2  'Center
      DataField       =   "Estado"
      DataSource      =   "dbBiblioteca"
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
      Height          =   390
      Index           =   5
      Left            =   2880
      TabIndex        =   13
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox txtInfo 
      Alignment       =   2  'Center
      DataField       =   "Género"
      DataSource      =   "dbBiblioteca"
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
      Height          =   390
      Index           =   4
      Left            =   2880
      TabIndex        =   12
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox txtInfo 
      Alignment       =   2  'Center
      DataField       =   "Año"
      DataSource      =   "dbBiblioteca"
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
      Height          =   390
      Index           =   3
      Left            =   2880
      MaxLength       =   4
      TabIndex        =   11
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox txtInfo 
      DataField       =   "Editorial"
      DataSource      =   "dbBiblioteca"
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
      Height          =   390
      Index           =   2
      Left            =   2880
      TabIndex        =   10
      Top             =   2040
      Width           =   5295
   End
   Begin VB.TextBox txtInfo 
      DataField       =   "Autor"
      DataSource      =   "dbBiblioteca"
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
      Height          =   390
      Index           =   1
      Left            =   2880
      TabIndex        =   9
      Top             =   1560
      Width           =   5295
   End
   Begin VB.TextBox txtInfo 
      DataField       =   "Título"
      DataSource      =   "dbBiblioteca"
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
      Height          =   390
      Index           =   0
      Left            =   2880
      TabIndex        =   8
      Top             =   1080
      Width           =   5295
   End
   Begin MSDataGridLib.DataGrid dgrBiblioteca 
      Bindings        =   "frmBooks.frx":26229
      Height          =   1815
      Left            =   1680
      TabIndex        =   0
      Top             =   5400
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   3201
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   16777215
      Enabled         =   0   'False
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   21
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   7
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
         DataField       =   "Título"
         Caption         =   "Título"
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
         DataField       =   "Autor"
         Caption         =   "Autor"
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
         DataField       =   "Editorial"
         Caption         =   "Editorial"
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
      BeginProperty Column04 
         DataField       =   "Año"
         Caption         =   "Año"
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
      BeginProperty Column05 
         DataField       =   "Género"
         Caption         =   "Género"
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
      BeginProperty Column06 
         DataField       =   "Estado"
         Caption         =   "Estado"
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
         SizeMode        =   1
         BeginProperty Column00 
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2085,166
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc dbBiblioteca 
      Height          =   375
      Left            =   2880
      Top             =   4080
      Width           =   5280
      _ExtentX        =   9313
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
      Connect         =   $"frmBooks.frx":26244
      OLEDBString     =   $"frmBooks.frx":263AF
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Libros"
      Caption         =   "Anterior                                  Siguiente"
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
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00004080&
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   6
      Left            =   360
      TabIndex        =   7
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00004080&
      Caption         =   "Género"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   5
      Left            =   360
      TabIndex        =   6
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00004080&
      Caption         =   "Año"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   4
      Left            =   360
      TabIndex        =   5
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00004080&
      Caption         =   "Editorial"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   3
      Left            =   360
      TabIndex        =   4
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00004080&
      Caption         =   "Autor"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00004080&
      Caption         =   "Título"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "LIBROS"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   615
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "frmBooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdButton_Click(Index As Integer)

Select Case Index

Case 0 'Registrar
 MsgBox "Ingrese la información y haga clic en guardar", vbInformation
  dbBiblioteca.Recordset.AddNew
   For Index = 0 To 5
        txtInfo(Index).Enabled = True
        txtInfo(0).SetFocus
      Next

Case 1 'Guardar
dbBiblioteca.Recordset.Update
dgrBiblioteca.Enabled = False
For Index = 0 To 5
        txtInfo(Index).Enabled = False
      Next

Case 2 'Buscar
Dim Buscar As String, Criterio As String
 Buscar = InputBox("Escriba el nombre del libro que desea buscar", "Búsqueda por nombre", vbQuestion)
  If Buscar = "" Then Exit Sub
    Criterio = "Título like '*" & Buscar & "*'"
    dbBiblioteca.Recordset.MoveNext
   If Not dbBiblioteca.Recordset.EOF Then
      dbBiblioteca.Recordset.Find Criterio
  End If
 If dbBiblioteca.Recordset.EOF Then
    dbBiblioteca.Recordset.MoveFirst
    dbBiblioteca.Recordset.Find Criterio
        If dbBiblioteca.Recordset.EOF Then
      dbBiblioteca.Recordset.MoveLast
Respuesta = MsgBox("Autor no registrado", vbCritical)
        End If
   End If

Case 3 'Editar
      For Index = 0 To 5
        txtInfo(Index).Enabled = True
        txtInfo(Index).SetFocus
      Next
     dgrBiblioteca.DataChanged = False
      
Case 4 'Eliminar
Dim Confirmacion As Integer
Confirmacion = MsgBox("¿Está seguro que desea eliminar los datos?", vbYesNoCancel + vbQuestion + vbDefaultButton2, "Eliminar")
   If Confirmacion = vbYes Then
    dbBiblioteca.Recordset.Delete
    MsgBox "Datos eliminados"
     dbBiblioteca.Recordset.MoveNext
If dbBiblioteca.Recordset.EOF Then
   dbBiblioteca.Recordset.MoveLast
   End If
     Else
       Exit Sub
End If

Case 5 'Cerrar
Unload Me
Load frmLibrary

 End Select
End Sub
