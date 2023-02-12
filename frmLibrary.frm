VERSION 5.00
Begin VB.Form frmLibrary 
   BackColor       =   &H00004080&
   Caption         =   "Biblioteca"
   ClientHeight    =   4635
   ClientLeft      =   4875
   ClientTop       =   3495
   ClientWidth     =   7095
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   Picture         =   "frmLibrary.frx":0000
   ScaleHeight     =   4635
   ScaleWidth      =   7095
   Begin VB.Label lblInfo 
      BackColor       =   &H00404080&
      BackStyle       =   0  'Transparent
      Caption         =   "Biblioteca"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1095
      Left            =   1320
      TabIndex        =   0
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Menu registros 
      Caption         =   "Registros"
      Begin VB.Menu libros 
         Caption         =   "Libros"
         Begin VB.Menu ftecnica 
            Caption         =   "Ficha técnica"
         End
         Begin VB.Menu autores 
            Caption         =   "Autores"
         End
      End
      Begin VB.Menu lectores 
         Caption         =   "Lectores"
      End
      Begin VB.Menu prestamos 
         Caption         =   "Préstamos"
      End
   End
End
Attribute VB_Name = "frmLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub autores_Click()
frmAutor.Show
End Sub

Private Sub ftecnica_Click()
frmBooks.Show

End Sub

Private Sub lectores_Click()
frmRead.Show
End Sub

Private Sub prestamos_Click()
frmLendings.Show
End Sub
