VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Ventas 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Ventas"
   ClientHeight    =   9270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16755
   LinkTopic       =   "Form1"
   ScaleHeight     =   9270
   ScaleWidth      =   16755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Detalle de la Venta"
      BeginProperty Font 
         Name            =   "Bodoni MT Poster Compressed"
         Size            =   15.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7695
      Left            =   9840
      TabIndex        =   34
      Top             =   240
      Width           =   6615
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FFFF&
         Caption         =   "Guardar venta"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   240
         Picture         =   "Ventas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   6240
         Width           =   1935
      End
      Begin VB.TextBox txtsaldoo 
         Enabled         =   0   'False
         Height          =   495
         Left            =   3720
         TabIndex        =   52
         Top             =   4440
         Width           =   2655
      End
      Begin VB.TextBox txtpago 
         Enabled         =   0   'False
         Height          =   495
         Left            =   3720
         TabIndex        =   50
         Top             =   5640
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker fecha 
         Height          =   495
         Left            =   3720
         TabIndex        =   48
         Top             =   840
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   873
         _Version        =   393216
         Format          =   148963329
         CurrentDate     =   43745
      End
      Begin VB.TextBox txtsaldo 
         Enabled         =   0   'False
         Height          =   495
         HideSelection   =   0   'False
         Left            =   3720
         TabIndex        =   47
         Top             =   5040
         Width           =   2655
      End
      Begin VB.TextBox txtañov 
         Enabled         =   0   'False
         Height          =   495
         Left            =   3720
         TabIndex        =   46
         Top             =   3840
         Width           =   2655
      End
      Begin VB.TextBox txtvendido 
         Enabled         =   0   'False
         Height          =   495
         Left            =   3720
         TabIndex        =   45
         Top             =   3240
         Width           =   2655
      End
      Begin VB.TextBox txtcontacto 
         Enabled         =   0   'False
         Height          =   495
         Left            =   3720
         TabIndex        =   44
         Top             =   2640
         Width           =   2655
      End
      Begin VB.TextBox txtapenom 
         Enabled         =   0   'False
         Height          =   495
         Left            =   3720
         TabIndex        =   43
         Top             =   2040
         Width           =   2655
      End
      Begin VB.TextBox txtcuitv 
         Enabled         =   0   'False
         Height          =   495
         Left            =   3720
         TabIndex        =   42
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Image Image13 
         Height          =   480
         Left            =   3000
         Picture         =   "Ventas.frx":08CA
         Top             =   3240
         Width           =   480
      End
      Begin VB.Image Image12 
         Height          =   480
         Left            =   3000
         Picture         =   "Ventas.frx":1994
         Top             =   2640
         Width           =   480
      End
      Begin VB.Image Image11 
         Height          =   480
         Left            =   3000
         Picture         =   "Ventas.frx":225E
         Top             =   2040
         Width           =   480
      End
      Begin VB.Image Image10 
         Height          =   480
         Left            =   3000
         Picture         =   "Ventas.frx":2B28
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080FFFF&
         Caption         =   "Valor del vehículo"
         BeginProperty Font 
            Name            =   "Bodoni MT Poster Compressed"
            Size            =   18
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   54
         Top             =   4440
         Width           =   2415
      End
      Begin VB.Image Image9 
         Height          =   480
         Left            =   3000
         Picture         =   "Ventas.frx":33F2
         Top             =   840
         Width           =   480
      End
      Begin VB.Image Image8 
         Height          =   480
         Left            =   3000
         Picture         =   "Ventas.frx":3CBC
         Top             =   3840
         Width           =   480
      End
      Begin VB.Image Image7 
         Height          =   480
         Left            =   3000
         Picture         =   "Ventas.frx":4586
         Top             =   4440
         Width           =   480
      End
      Begin VB.Image Image6 
         Height          =   480
         Left            =   3000
         Picture         =   "Ventas.frx":4E50
         Top             =   5040
         Width           =   480
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   3000
         Picture         =   "Ventas.frx":571A
         Top             =   5640
         Width           =   480
      End
      Begin VB.Label Label14 
         BackColor       =   &H0080FFFF&
         Caption         =   "Saldo"
         BeginProperty Font 
            Name            =   "Bodoni MT Poster Compressed"
            Size            =   18
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   51
         Top             =   5040
         Width           =   1455
      End
      Begin VB.Label Label13 
         BackColor       =   &H0080FFFF&
         Caption         =   "A pagar"
         BeginProperty Font 
            Name            =   "Bodoni MT Poster Compressed"
            Size            =   18
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   49
         Top             =   5640
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackColor       =   &H0080FFFF&
         Caption         =   "Precio del vehículo"
         BeginProperty Font 
            Name            =   "Bodoni MT Poster Compressed"
            Size            =   12
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   41
         Top             =   5160
         Width           =   1455
      End
      Begin VB.Label Label11 
         BackColor       =   &H0080FFFF&
         Caption         =   "Año"
         BeginProperty Font 
            Name            =   "Bodoni MT Poster Compressed"
            Size            =   18
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   40
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label Label10 
         BackColor       =   &H0080FFFF&
         Caption         =   "Vehículo vendido:"
         BeginProperty Font 
            Name            =   "Bodoni MT Poster Compressed"
            Size            =   18
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   39
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackColor       =   &H0080FFFF&
         Caption         =   "Contacto:"
         BeginProperty Font 
            Name            =   "Bodoni MT Poster Compressed"
            Size            =   18
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   38
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackColor       =   &H0080FFFF&
         Caption         =   "Apellido y Nombre:"
         BeginProperty Font 
            Name            =   "Bodoni MT Poster Compressed"
            Size            =   18
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   37
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080FFFF&
         Caption         =   "Número de CUIT:"
         BeginProperty Font 
            Name            =   "Bodoni MT Poster Compressed"
            Size            =   18
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   36
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackColor       =   &H0080FFFF&
         Caption         =   "Fecha de venta:"
         BeginProperty Font 
            Name            =   "Bodoni MT Poster Compressed"
            Size            =   18
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   35
         Top             =   840
         Width           =   1815
      End
   End
   Begin VB.TextBox txtdire 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6600
      TabIndex        =   33
      Top             =   4680
      Width           =   2655
   End
   Begin VB.TextBox txtfnac 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6600
      TabIndex        =   32
      Top             =   4200
      Width           =   2655
   End
   Begin VB.TextBox txtnom 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6600
      TabIndex        =   31
      Top             =   3720
      Width           =   2655
   End
   Begin VB.TextBox txtapellido 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6600
      TabIndex        =   30
      Top             =   3240
      Width           =   2655
   End
   Begin VB.TextBox txtcuit 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6600
      TabIndex        =   29
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H0080FFFF&
      Caption         =   "Datos del comprador"
      BeginProperty Font 
         Name            =   "Bodoni MT Poster Compressed"
         Size            =   15.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   6360
      TabIndex        =   25
      Top             =   240
      Width           =   3375
      Begin VB.TextBox txttel 
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   4920
         Width           =   2655
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H0080FFFF&
         Caption         =   "Búsqueda por CUIT"
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
         Left            =   240
         Picture         =   "Ventas.frx":5FE4
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   26
         Top             =   600
         Width           =   2895
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FFFF&
      Caption         =   "CERRAR"
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
      Left            =   3480
      Picture         =   "Ventas.frx":68AE
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8280
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Calcular"
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
      Left            =   1440
      Picture         =   "Ventas.frx":7178
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8280
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0080FFFF&
      Caption         =   "Forma de Pago"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   4920
      TabIndex        =   12
      Top             =   5880
      Width           =   3375
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   1920
         TabIndex        =   20
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   2040
         TabIndex        =   18
         Top             =   960
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Ventas.frx":7A42
         Left            =   1920
         List            =   "Ventas.frx":7A55
         TabIndex        =   16
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H0080FFFF&
         Caption         =   "Contado"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   1440
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0080FFFF&
         Caption         =   "Cuotas"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080FFFF&
         Caption         =   "Interes Simple %"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Width           =   1815
      End
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   495
      Left            =   3240
      TabIndex        =   11
      Top             =   7560
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   495
      Left            =   3240
      TabIndex        =   9
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   495
      Left            =   3240
      TabIndex        =   7
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Datos del Vehículo"
      BeginProperty Font 
         Name            =   "Bodoni MT Poster Compressed"
         Size            =   15.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6135
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   2400
         TabIndex        =   24
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H0080FFFF&
         Caption         =   "Búsqueda por ID"
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
         Left            =   360
         Picture         =   "Ventas.frx":7A6D
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtfoto 
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   4920
         Width           =   2655
      End
      Begin VB.TextBox txtprecio 
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   4440
         Width           =   2655
      End
      Begin VB.TextBox txtaño 
         Enabled         =   0   'False
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   5
         Top             =   3960
         Width           =   2655
      End
      Begin VB.TextBox txtnombre 
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   3480
         Width           =   2655
      End
      Begin VB.TextBox txtmarca 
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   3000
         Width           =   2655
      End
      Begin VB.TextBox txtcategoria 
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   2520
         Width           =   2655
      End
      Begin VB.TextBox txtid 
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   2040
         Width           =   2655
      End
      Begin VB.Image Image4 
         Height          =   3255
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   3015
      End
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   2640
      Picture         =   "Ventas.frx":8337
      Top             =   7560
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   2640
      Picture         =   "Ventas.frx":8C01
      Top             =   6720
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2640
      Picture         =   "Ventas.frx":94CB
      Top             =   5880
      Width           =   480
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FFFF&
      Caption         =   "Saldo"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   10
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Anticipo 50%"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   8
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Valor del Vehículo"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   6
      Top             =   5880
      Width           =   1335
   End
End
Attribute VB_Name = "Ventas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
rs_ventas.AddNew
    
    
    
    rs_ventas("fecha_venta") = fecha.Value
    rs_ventas("cuit_venta") = txtcuitv.Text
    rs_ventas("name_person") = txtapenom.Text
    rs_ventas("contacto") = txtcontacto.Text
    rs_ventas("name_car") = txtvendido.Text
    rs_ventas("año_v") = txtañov.Text
    rs_ventas("precio_v") = txtsaldoo.Text
    rs_ventas("saldo_v") = txtsaldo.Text
    rs_ventas("apagar_v") = txtpago.Text
   
    rs_ventas.Update
    MsgBox "Venta cargada al sistema", 32, "Correctamente"
    limpiartodo
End Sub

Private Sub Command2_Click()

If Option1.Value = True And Combo1.Text = "12" And IsNumeric(Text8.Text) Then

meses = Val(Combo1.Text)
tiempo = meses / 12
interes = Val(Text8.Text) / 100
capital = Val(Text5.Text)
calculo = capital * interes * tiempo
interesmes = calculo / Val(Combo1.Text)
cuotasininteres = Val(Text5.Text) / 12
interesfinalpormes = interesmes + cuotasininteres
MsgBox "El valor a pagar es de 12 cuota de $ " & interesfinalpormes, 32, "Valor de las cuotas"
txtpago.Text = "12 Cuotas de $ " & interesfinalpormes

ElseIf Option1.Value = True And Combo1.Text = "18" And IsNumeric(Text8.Text) Then
meses = Val(Combo1.Text)
tiempo = meses / 12
interes = Val(Text8.Text) / 100
capital = Val(Text5.Text)
calculo = capital * interes * tiempo
interesmes = calculo / Val(Combo1.Text)
cuotasininteres = Val(Text5.Text) / 18
interesfinalpormes = interesmes + cuotasininteres
MsgBox "El valor a pagar es de 18 cuota de $ " & interesfinalpormes, 32, "Valor de las cuotas"
txtpago.Text = "18 Cuotas de $ " & interesfinalpormes

ElseIf Option1.Value = True And Combo1.Text = "24" And IsNumeric(Text8.Text) Then
meses = Val(Combo1.Text)
tiempo = meses / 12
interes = Val(Text8.Text) / 100
capital = Val(Text5.Text)
calculo = capital * interes * tiempo
interesmes = calculo / Val(Combo1.Text)
cuotasininteres = Val(Text5.Text) / 24
interesfinalpormes = interesmes + cuotasininteres
MsgBox "El valor a pagar es de 24 cuota de $ " & interesfinalpormes, 32, "Valor de las cuotas"
txtpago.Text = "24 Cuotas de $ " & interesfinalpormes

ElseIf Option1.Value = True And Combo1.Text = "36" And IsNumeric(Text8.Text) Then
meses = Val(Combo1.Text)
tiempo = meses / 12
interes = Val(Text8.Text) / 100
capital = Val(Text5.Text)
calculo = capital * interes * tiempo
interesmes = calculo / Val(Combo1.Text)
cuotasininteres = Val(Text5.Text) / 36
interesfinalpormes = interesmes + cuotasininteres
MsgBox "El valor a pagar es de 36 cuota de $ " & interesfinalpormes, 32, "Valor de las cuotas"
txtpago.Text = "36 Cuotas de $ " & interesfinalpormes


ElseIf Option1.Value = True And Combo1.Text = "48" And IsNumeric(Text8.Text) Then
meses = Val(Combo1.Text)
tiempo = meses / 12
interes = Val(Text8.Text) / 100
capital = Val(Text5.Text)
calculo = capital * interes * tiempo
interesmes = calculo / Val(Combo1.Text)
cuotasininteres = Val(Text5.Text) / 24
interesfinalpormes = interesmes + cuotasininteres
MsgBox "El valor a pagar es de 48 cuota de $ " & interesfinalpormes, 32, "Valor de las cuotas"
txtpago.Text = "48 Cuotas de $ " & interesfinalpormes
ElseIf Option2.Value = True Then
descuento = Val(Text3.Text) * 0.25
contado = Val(Text3.Text) - descuento
MsgBox "El valor al contado a pagar es de $ " & contado, 32, "Debe abonar"
txtpago.Text = "25% descuento Contado $" & contado
txtsaldo.Text = 0
Else
MsgBox "Debe tildar una opcion o falta colocar el interes"

End If




End Sub


Private Sub Command3_Click()
respuesta = MsgBox("¿Desea salir?", vbYesNo + 16, "Continuar")
If respuesta = vbYes Then
Me.Hide
limpiartodo
End If
End Sub

Private Sub Command4_Click()
ListadoClientes.Show

End Sub


Private Sub Command7_Click()

Set rs = db.OpenRecordset("select * from auto where id_auto = " & Val(Text12.Text) & "")

If Not (rs.EOF And rs.BOF) Then
    txtid(0).Text = rs.Fields("id_auto")
    txtcategoria(1).Text = rs.Fields("categoria")
    txtmarca(2).Text = rs.Fields("marca")
    txtnombre(3).Text = rs.Fields("nombre")
    txtaño(4).Text = rs.Fields("año")
    txtprecio.Text = rs.Fields("precio")
    txtfoto.Text = rs.Fields("foto")
    X = App.Path
    Image4.Picture = LoadPicture(X + "\" + txtfoto.Text)
    Text3.Text = txtprecio.Text
    anticipo = Val(Text3.Text) * 0.5
    Text4.Text = anticipo
    saldo = Val(Text3.Text) - Val(Text4.Text)
    Text5.Text = saldo
    txtvendido.Text = txtcategoria(1) + "," + txtmarca(2).Text + "," + txtnombre(3).Text
    txtsaldo.Text = Text5.Text
    txtañov.Text = txtaño(4).Text
    txtsaldoo.Text = txtprecio.Text
    
    
    
    
    
     
Else
    MsgBox "No se encontró ningún dato"
    
End If
End Sub

Private Sub Command9_Click()
Set rs_clientes = db.OpenRecordset("select * from cliente where cuit = '" & Val(Text1.Text) & "'")

If Not (rs_clientes.EOF And rs_clientes.BOF) Then
    txtcuit.Text = rs_clientes.Fields("cuit")
    txtapellido.Text = rs_clientes.Fields("apellido")
    txtnom.Text = rs_clientes.Fields("nombre")
    txtfnac.Text = rs_clientes.Fields("fechanacimiento")
    txtdire.Text = rs_clientes.Fields("domicilio")
    txttel.Text = rs_clientes.Fields("telefono")
    txtcuitv.Text = txtcuit.Text
    txtapenom.Text = txtapellido.Text + "," + txtnom.Text
    txtcontacto.Text = txttel.Text + " - " + txtdire.Text
    
    
    
     
Else
    MsgBox "No se encontró ningún dato"
    
End If
End Sub

Private Sub Form_Load()
Set db = OpenDatabase("D:\Usuarios\User\Desktop\Consecionaria\Ventaautos")
Set rs_ventas = db.OpenRecordset("select * from ventas")
Option1.Value = False
Option2.Value = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
db.Close
rs_ventas.Close
rs.Close
rs_clientes.Close
End Sub
Sub limpiartodo()
    txtcuit.Text = ""
    txtapellido.Text = ""
    txtnom.Text = ""
    txtfnac.Text = ""
    txtdire.Text = ""
    txttel.Text = ""
    txtcuitv.Text = ""
    txtapenom.Text = ""
    txtcontacto.Text = ""
    txtid(0).Text = ""
    txtcategoria(1).Text = ""
    txtmarca(2).Text = ""
    txtnombre(3).Text = ""
    txtaño(4).Text = ""
    txtprecio.Text = ""
    txtfoto.Text = ""
    Image4.Picture = LoadPicture("")
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    txtvendido.Text = ""
    txtsaldo.Text = ""
    txtañov.Text = ""
    txtsaldoo.Text = ""
    Text1.Text = ""
    Text12.Text = ""
    Text8.Text = ""
    txtpago.Text = ""
    Option1.Value = False
    Option2.Value = False
    
End Sub
