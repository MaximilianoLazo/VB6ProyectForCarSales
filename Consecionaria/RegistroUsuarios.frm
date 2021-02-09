VERSION 5.00
Begin VB.Form RegistroUsuarios 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   0  'None
   Caption         =   "Registro"
   ClientHeight    =   3960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Volver al Login"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3240
      Picture         =   "RegistroUsuarios.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Agregar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Picture         =   "RegistroUsuarios.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   1920
      Picture         =   "RegistroUsuarios.frx":1194
      Top             =   1440
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   1920
      Picture         =   "RegistroUsuarios.frx":1A5E
      Top             =   480
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "RegistroUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
db.Execute ("Insert into login(usuario,contraseña) Values ('" & Text1.Text & " ' ,'" & Text2.Text & " ')")
rs.AddNew
rs("usuario") = Text1.Text
rs("contraseña") = Text2.Text
rs.Update
MsgBox "Nuevo registro creado", 32, "Carga Exitosa"

Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
End Sub

Private Sub Command2_Click()
respuesta = MsgBox("¿Desea volver al login?", vbYesNo + 32, " Continuar")
If respuesta = vbYes Then
RegistroUsuarios.Hide
login.Show
Else
Text1.SetFocus
End If
End Sub

Private Sub Form_Load()
Set db = OpenDatabase("C:\Users\Mona\Desktop\VisualBasic 2d año\Consecionaria\login.mdb")
Set rs = db.OpenRecordset("select * from login")
End Sub
