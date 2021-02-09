VERSION 5.00
Begin VB.Form MenuPrincipal 
   Caption         =   "Agencia Particular solo 0km"
   ClientHeight    =   8160
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   14115
   Icon            =   "MenuPrincipal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   14115
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   1680
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
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3015
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
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   8175
      Left            =   0
      Picture         =   "MenuPrincipal.frx":10CA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14175
   End
   Begin VB.Menu autos 
      Caption         =   "Autos"
      Begin VB.Menu nuevo 
         Caption         =   "Alta-Baja-Modificacion-Borrar"
         Index           =   1
      End
      Begin VB.Menu buscar 
         Caption         =   "Listado y Búsqueda"
      End
   End
   Begin VB.Menu clientes 
      Caption         =   "Clientes"
      Begin VB.Menu nuevocliente 
         Caption         =   "Alta-Baja-Modificación-Borrar"
      End
      Begin VB.Menu buscarcliente 
         Caption         =   "Listado y Búsqueda"
      End
   End
   Begin VB.Menu venta 
      Caption         =   "Venta"
      Begin VB.Menu realizarventa 
         Caption         =   "Realizar Venta"
      End
      Begin VB.Menu ventalistado 
         Caption         =   "Listado y Búsqueda"
      End
   End
   Begin VB.Menu Reportes 
      Caption         =   "Reportes"
      Begin VB.Menu impvehiculos 
         Caption         =   "Imprimir Vehículos"
      End
      Begin VB.Menu impclientes 
         Caption         =   "Imprimir Clientes"
      End
      Begin VB.Menu impventas 
         Caption         =   "Imprimir Ventas"
      End
   End
End
Attribute VB_Name = "MenuPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub buscar_Click()


ListadoAutos.Show

End Sub

Private Sub buscarcliente_Click()
ListadoClientes.Show

End Sub

Private Sub Form_Resize()
Rem redimensionar imagen

Image1.Width = MenuPrincipal.Width
Image1.Height = MenuPrincipal.Height

End Sub

Private Sub impclientes_Click()
DataReport2.Show
End Sub

Private Sub impvehiculos_Click()
DataReport1.Show

End Sub

Private Sub impventas_Click()
DataReport3.Show

End Sub

Private Sub nuevo_Click(Index As Integer)

CargarNuevoAuto.Show
End Sub

Private Sub nuevocliente_Click()
CargarCliente.Show

End Sub

Private Sub realizarventa_Click()
Ventas.Show

End Sub

Private Sub Timer1_Timer()
Label3.Caption = Date
Label4.Caption = Time

End Sub

Private Sub ventalistado_Click()
ListadoVentas.Show

End Sub
