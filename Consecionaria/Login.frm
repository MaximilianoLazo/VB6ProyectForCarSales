VERSION 5.00
Begin VB.Form Login 
   Caption         =   "Login"
   ClientHeight    =   7590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11730
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   11730
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080FFFF&
      Caption         =   "Borrar"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6480
      Picture         =   "Login.frx":10CA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1680
      Top             =   1920
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FFFF&
      Caption         =   "Registro"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      Picture         =   "Login.frx":1994
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6480
      Picture         =   "Login.frx":225E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "INGRESAR"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      Picture         =   "Login.frx":2B28
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   6600
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   6600
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   240
      Width           =   3015
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   6000
      Picture         =   "Login.frx":33F2
      Top             =   2640
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   6000
      Picture         =   "Login.frx":3CBC
      Top             =   1920
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "CONTRASEÑA:"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   4560
      TabIndex        =   5
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   7575
      Left            =   0
      Picture         =   "Login.frx":4586
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command3_Click()
contra = InputBox("Ingrese la contraseña para poder ir a crear nuevo registro", "Permisos Administrador")
If contra = "admin" Then
MsgBox "Contraseña correcta bienvenido al panel de administrador ", 32, "Panel Adminitrador"
login.Hide
RegistroUsuarios.Show

Else
MsgBox "Contraseña incorrecta vuelva a intentarlo", 16, "Ha ocurrido un error"
End If


End Sub

Private Sub Command4_Click()
Text1.Text = ""
Text2.Text = ""

End Sub

Private Sub Form_Load()
Set db = OpenDatabase("D:\Usuarios\User\Desktop\Consecionaria\login.mdb")
Set rs = db.OpenRecordset("select * from login")
End Sub

Private Sub Form_Resize()
Rem redimensionar imagen

Image1.Width = login.Width
Image1.Height = login.Height

End Sub

Private Sub Command1_Click()
Dim i As Integer
i = 0
rs.MoveFirst
While Not rs.EOF = True
    If rs.Fields(0).Value = Text1.Text And rs.Fields(1).Value = Text2.Text Then
    i = i + 1
    MsgBox "Bienvenido", 32, "Welcome"
    
    login.Hide
    Cargando.Show
    End If
    rs.MoveNext
    
    
Wend
    If i = 0 Then
    MsgBox "Error vuelva a intentarlo", 16, "Error"
    End If
    

End Sub

Private Sub Command2_Click()
respuesta = MsgBox("Desea salir de la aplicacion", vbYesNo + 16, "Continuar")
If respuesta = vbYes Then
End
End If

End Sub


Private Sub Timer1_Timer()
Label3.Caption = Date
Label4.Caption = Time
End Sub
