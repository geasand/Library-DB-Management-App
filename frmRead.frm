VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRead 
   BackColor       =   &H00004080&
   Caption         =   "Lectores"
   ClientHeight    =   8250
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11655
   LinkTopic       =   "Form1"
   Picture         =   "frmRead.frx":0000
   ScaleHeight     =   8250
   ScaleWidth      =   11655
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid dgrLectores 
      Bindings        =   "frmRead.frx":A737
      Height          =   3375
      Left            =   5880
      TabIndex        =   15
      Top             =   1440
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5953
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
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
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
         DataField       =   "Cédula"
         Caption         =   "Cédula"
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
         DataField       =   "Teléfono"
         Caption         =   "Teléfono"
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
         DataField       =   "Estatus de cliente"
         Caption         =   "Estatus de cliente"
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
         BeginProperty Column04 
            ColumnWidth     =   2429,858
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoLectores 
      Height          =   495
      Left            =   360
      Top             =   4320
      Width           =   5055
      _ExtentX        =   8916
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
      Connect         =   $"frmRead.frx":A751
      OLEDBString     =   $"frmRead.frx":A8BC
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Lectores"
      Caption         =   "Anterior                             Siguiente"
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
      Left            =   10320
      TabIndex        =   14
      Top             =   7680
      Width           =   1095
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
      Left            =   2160
      TabIndex        =   13
      Top             =   6360
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
      Left            =   3000
      TabIndex        =   12
      Top             =   5760
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
      Left            =   1320
      TabIndex        =   11
      Top             =   5760
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
      Left            =   3000
      TabIndex        =   10
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Nuevo"
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
      Left            =   1320
      TabIndex        =   9
      Top             =   5160
      Width           =   1455
   End
   Begin VB.TextBox txtInfo 
      Alignment       =   2  'Center
      DataField       =   "Estatus de cliente"
      DataSource      =   "adoLectores"
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
      Left            =   2760
      TabIndex        =   8
      Top             =   3600
      Width           =   2655
   End
   Begin VB.TextBox txtInfo 
      Alignment       =   2  'Center
      DataField       =   "Teléfono"
      DataSource      =   "adoLectores"
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
      Left            =   2760
      MaxLength       =   11
      TabIndex        =   7
      Top             =   2880
      Width           =   2655
   End
   Begin VB.TextBox txtInfo 
      Alignment       =   2  'Center
      DataField       =   "Cédula"
      DataSource      =   "adoLectores"
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
      Left            =   2760
      MaxLength       =   10
      TabIndex        =   6
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox txtInfo 
      Alignment       =   2  'Center
      DataField       =   "Nombre"
      DataSource      =   "adoLectores"
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
      Left            =   2760
      TabIndex        =   5
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Estatus"
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
      Left            =   600
      TabIndex        =   4
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono"
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
      Left            =   600
      TabIndex        =   3
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cédula"
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
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Lectores"
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
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
End
Attribute VB_Name = "frmRead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdButton_Click(Index As Integer)

Select Case Index

Case 0 'Registrar
MsgBox "Ingrese la información y haga clic en guardar"
adoLectores.Recordset.AddNew
    For Index = 0 To 3
        txtInfo(Index).Enabled = True
        txtInfo(0).SetFocus
      Next

Case 1 'Guardar
adoLectores.Recordset.Update
dgrLectores.Enabled = False
For Index = 0 To 3
        txtInfo(Index).Enabled = False
      Next

Case 2 'Buscar
Dim Buscar As String, Criterio As String
 Buscar = InputBox("Escriba el nombre del usuario", "Búsqueda por nombre", vbQuestion)
  If Buscar = "" Then Exit Sub
    Criterio = "Nombre like '*" & Buscar & "*'"
    adoLectores.Recordset.MoveNext
   If Not adoLectores.Recordset.EOF Then
      adoLectores.Recordset.Find Criterio
  End If
 If adoLectores.Recordset.EOF Then
    adoLectores.Recordset.MoveFirst
    adoLectores.Recordset.Find Criterio
        If adoLectores.Recordset.EOF Then
      adoLectores.Recordset.MoveLast
Respuesta = MsgBox("No registrado", vbCritical)
        End If
   End If


Case 3 'Editar
      For Index = 0 To 3
        txtInfo(Index).Enabled = True
        txtInfo(0).SetFocus
      Next
    dgrLectores.DataChanged = False
      
Case 4 'Eliminar
Dim Confirmacion As Integer
Confirmacion = MsgBox("¿Está seguro que desea eliminar los datos?", vbYesNoCancel + vbQuestion + vbDefaultButton2, "Eliminar")
   If Confirmacion = vbYes Then
    adoLectores.Recordset.Delete
    MsgBox "Datos eliminados"
     adoLectores.Recordset.MoveNext
If adoLectores.Recordset.EOF Then
   adoLectores.Recordset.MoveLast
   End If
     Else
       Exit Sub
End If

Case 5 'Cerrar
Unload Me
Load frmLibrary

 End Select

End Sub
