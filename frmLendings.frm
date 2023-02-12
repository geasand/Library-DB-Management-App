VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmLendings 
   BackColor       =   &H00404080&
   Caption         =   "Libros prestados"
   ClientHeight    =   8250
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   Picture         =   "frmLendings.frx":0000
   ScaleHeight     =   8250
   ScaleWidth      =   11640
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid dgrPrestados 
      Bindings        =   "frmLendings.frx":A737
      Height          =   3495
      Left            =   5880
      TabIndex        =   17
      Top             =   1320
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   6165
      _Version        =   393216
      AllowUpdate     =   -1  'True
      Enabled         =   0   'False
      HeadLines       =   1
      RowHeight       =   24
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
         Size            =   11.25
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
            LCID            =   3082
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
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         ScrollBars      =   2
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoPrestados 
      Height          =   495
      Left            =   360
      Top             =   4800
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   873
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
      Connect         =   $"frmLendings.frx":A752
      OLEDBString     =   $"frmLendings.frx":A8BD
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Prestados"
      Caption         =   "Anterior                           Siguiente"
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
      Left            =   10080
      TabIndex        =   16
      Top             =   7680
      Width           =   1455
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
      Left            =   2040
      TabIndex        =   15
      Top             =   6720
      Width           =   1455
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
      Left            =   2880
      TabIndex        =   14
      Top             =   6120
      Width           =   1455
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
      Left            =   1200
      TabIndex        =   13
      Top             =   6120
      Width           =   1455
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
      Height          =   495
      Index           =   1
      Left            =   2880
      TabIndex        =   12
      Top             =   5520
      Width           =   1455
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
      Left            =   1200
      TabIndex        =   11
      Top             =   5520
      Width           =   1455
   End
   Begin VB.TextBox txtInfo 
      Alignment       =   2  'Center
      DataField       =   "Estado"
      DataSource      =   "adoPrestados"
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
      Index           =   4
      Left            =   2280
      TabIndex        =   10
      Top             =   4200
      Width           =   2895
   End
   Begin VB.TextBox txtInfo 
      Alignment       =   2  'Center
      DataField       =   "Devolución"
      DataSource      =   "adoPrestados"
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
      Index           =   3
      Left            =   2280
      TabIndex        =   9
      Top             =   3480
      Width           =   2895
   End
   Begin VB.TextBox txtInfo 
      Alignment       =   2  'Center
      DataField       =   "Préstamo"
      DataSource      =   "adoPrestados"
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
      Left            =   2280
      TabIndex        =   8
      Top             =   2760
      Width           =   2895
   End
   Begin VB.TextBox txtInfo 
      Alignment       =   2  'Center
      DataField       =   "Libro"
      DataSource      =   "adoPrestados"
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
      Left            =   2280
      TabIndex        =   7
      Top             =   2040
      Width           =   2895
   End
   Begin VB.TextBox txtInfo 
      Alignment       =   2  'Center
      DataField       =   "Lector"
      DataSource      =   "adoPrestados"
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
      Left            =   2280
      TabIndex        =   6
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00EBEBED&
      BackStyle       =   0  'Transparent
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
      Height          =   495
      Index           =   5
      Left            =   720
      TabIndex        =   5
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00EBEBED&
      BackStyle       =   0  'Transparent
      Caption         =   "Devolución"
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
      Index           =   4
      Left            =   360
      TabIndex        =   4
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00EBEBED&
      BackStyle       =   0  'Transparent
      Caption         =   "Préstamo"
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
      Index           =   3
      Left            =   480
      TabIndex        =   3
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00EBEBED&
      BackStyle       =   0  'Transparent
      Caption         =   "Libro"
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
      Left            =   720
      TabIndex        =   2
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00EBEBED&
      BackStyle       =   0  'Transparent
      Caption         =   "Lector"
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
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00EBEBED&
      BackStyle       =   0  'Transparent
      Caption         =   "Libros prestados"
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
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
End
Attribute VB_Name = "frmLendings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdButton_Click(Index As Integer)

Select Case Index

Case 0 'Registrar
MsgBox "Ingrese la información y haga clic en guardar"
adoPrestados.Recordset.AddNew
    For Index = 0 To 4
        txtInfo(Index).Enabled = True
        txtInfo(0).SetFocus
      Next

Case 1 'Guardar
adoPrestados.Recordset.Update
dgrPrestados.Enabled = False
For Index = 0 To 4
        txtInfo(Index).Enabled = False
      Next

Case 2 'Buscar
Dim Buscar As String, Criterio As String
 Buscar = InputBox("Escriba el nombre del usuario", "Búsqueda por nombre", vbQuestion)
  If Buscar = "" Then Exit Sub
    Criterio = "Lector like '*" & Buscar & "*'"
    adoPrestados.Recordset.MoveNext
   If Not adoPrestados.Recordset.EOF Then
      adoPrestados.Recordset.Find Criterio
  End If
 If adoPrestados.Recordset.EOF Then
    adoPrestados.Recordset.MoveFirst
    adoPrestados.Recordset.Find Criterio
        If adoPrestados.Recordset.EOF Then
      adoPrestados.Recordset.MoveLast
Respuesta = MsgBox("No registrado", vbCritical)
        End If
   End If
Case 3 'Editar
      For Index = 0 To 4
        txtInfo(Index).Enabled = True
        txtInfo(0).SetFocus
      Next
    dgrPrestados.DataChanged = False
      
Case 4 'Eliminar
Dim Confirmacion As Integer
Confirmacion = MsgBox("¿Está seguro que desea eliminar los datos?", vbYesNoCancel + vbQuestion + vbDefaultButton2, "Eliminar")
   If Confirmacion = vbYes Then
    adoPrestados.Recordset.Delete
    MsgBox "Datos eliminados"
     adoPrestados.Recordset.MoveNext
If adoPrestados.Recordset.EOF Then
   adoPrestados.Recordset.MoveLast
   End If
     Else
       Exit Sub
End If

Case 5 'Cerrar
Unload Me
Load frmLibrary

 End Select

End Sub
