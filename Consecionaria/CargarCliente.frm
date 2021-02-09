VERSION 5.00
Begin VB.Form CargarCliente 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   0  'None
   Caption         =   "CargarCliente"
   ClientHeight    =   7875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text12 
      Height          =   495
      Left            =   1800
      TabIndex        =   25
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H0080FFFF&
      Caption         =   "Buscar CUIT"
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
      Picture         =   "CargarCliente.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   120
      Width           =   1355
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H0080FFFF&
      Caption         =   "Modificar"
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
      Left            =   5640
      Picture         =   "CargarCliente.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6840
      Width           =   1355
   End
   Begin VB.TextBox Text10 
      Height          =   495
      Left            =   6000
      MaxLength       =   4
      TabIndex        =   22
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox Text9 
      Height          =   495
      Left            =   4440
      MaxLength       =   2
      TabIndex        =   21
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080FFFF&
      Caption         =   "Nuevo"
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
      Left            =   960
      Picture         =   "CargarCliente.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0080FFFF&
      Caption         =   "Agregar"
      Enabled         =   0   'False
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
      Left            =   2520
      Picture         =   "CargarCliente.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6840
      Width           =   1335
   End
   Begin VB.TextBox Text8 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      Height          =   495
      Left            =   6360
      MaxLength       =   1
      TabIndex        =   18
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox Text7 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      Height          =   495
      Left            =   3000
      MaxLength       =   2
      TabIndex        =   17
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   3000
      MaxLength       =   2
      TabIndex        =   16
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H0080FFFF&
      Caption         =   "Siguiente"
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
      Left            =   4080
      Picture         =   "CargarCliente.frx":2328
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080FFFF&
      Caption         =   "Anterior"
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
      Left            =   2520
      Picture         =   "CargarCliente.frx":2BF2
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H0080FFFF&
      Caption         =   "Último"
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
      Left            =   5640
      Picture         =   "CargarCliente.frx":34BC
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H0080FFFF&
      Caption         =   "Primero"
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
      Left            =   960
      Picture         =   "CargarCliente.frx":3D86
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
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
      Left            =   4080
      Picture         =   "CargarCliente.frx":4650
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6840
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   4080
      Width           =   3975
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   2640
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   1920
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      Height          =   495
      Left            =   3840
      MaxLength       =   8
      TabIndex        =   0
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   4800
      Width           =   3975
   End
   Begin VB.Image Image9 
      Height          =   495
      Left            =   6360
      Picture         =   "CargarCliente.frx":4F1A
      Stretch         =   -1  'True
      Top             =   360
      Width           =   495
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   5280
      X2              =   5760
      Y1              =   3840
      Y2              =   3360
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   3720
      X2              =   4200
      Y1              =   3840
      Y2              =   3360
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   2280
      Picture         =   "CargarCliente.frx":57E4
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   2280
      Picture         =   "CargarCliente.frx":60AE
      Top             =   4800
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   2280
      Picture         =   "CargarCliente.frx":6978
      Top             =   2640
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   2280
      Picture         =   "CargarCliente.frx":7242
      Top             =   1920
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   2280
      Picture         =   "CargarCliente.frx":7B0C
      Top             =   4080
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2280
      Picture         =   "CargarCliente.frx":83D6
      Top             =   3360
      Width           =   480
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DOMICILIO"
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
      Left            =   0
      TabIndex        =   10
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA DE NACIMIENTO"
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
      Left            =   0
      TabIndex        =   9
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE"
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
      Left            =   0
      TabIndex        =   8
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "APELLIDO"
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
      Left            =   0
      TabIndex        =   7
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CUIT"
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
      Left            =   0
      TabIndex        =   6
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TELÉFONO"
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
      Left            =   0
      TabIndex        =   5
      Top             =   4800
      Width           =   1935
   End
End
Attribute VB_Name = "CargarCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim connection As New ADODB.connection
Dim record As New ADODB.Recordset
'Option Explicit
'Dim datos As New ADODB.Connection
'Private WithEvents s As ADODB.Recordset
Private Sub Command1_Click()
Me.Hide
End Sub

Private Sub Command10_Click()
'------------------------ BOTÓN DE BÚSQUEDA---------------------------------
record.Close
record.Open "select * from cliente where cuit= '" & Text12.Text & "'", connection, adOpenDynamic, adLockPessimistic
If Not record.EOF Then
    mostrar
    reiniciar
Else
MsgBox "Cliente no encontrado", vbCritical, "Mensaje"
escondercuit
End If
'------------------------ BÚSQUEDA---------------------------------
's.Find "cuit = '" & Text12.Text & "'", 1
'If s.EOF = False And s.BOF = False Then
    'Text1.Text = s.Fields("cuit")
    'Text2.Text = s.Fields("apellido")
    'Text3.Text = s.Fields("nombre")
    'Text10.Text = s.Fields("fechanacimiento")
    'Text5.Text = s.Fields("domicilio")
    'Text6.Text = s.Fields("telefono")
'Else
    'MsgBox "No se encontró ningún dato", vbCritical, "ERROR"
'End If
End Sub

Private Sub Command11_Click()
'------------------------------- MODIFICAR ----------------------------------------
record.Update Array("cuit", "apellido", "nombre", "fechanacimiento", "domicilio", "telefono"), Array(Text1.Text, Text2.Text, Text3.Text, Text10.Text, Text5.Text, Text6.Text)
record.Update
If record.State = 1 Or record.State = 0 Then
MsgBox "Registro actualizado con exito", vbInformation, "Mensaje"
Else
MsgBox "Ha ocurrido un error", vbCritical, "Mensaje"
escondercuit
End If
'------------------------------- MODIFICAR REGISTRO----------------------------------------
'record.Fields("cuit").Value = Text1.Text
'record.Fields("apellido").Value = Text2.Text
'record.Fields("nombre").Value = Text3.Text
'record.Fields("fechanacimiento").Value = Text10.Text
'record.Fields("domicilio").Value = Text5.Text
'record.Fields("telefono").Value = Text6.Text
'MsgBox "Registro actualizado con exito", vbInformation, "Mensaje"
'record.Update
End Sub

Private Sub Command2_Click()
'----------------------------- ELIMINAR REGISTRO ACTUAL---------------------------------------------
respuesta = MsgBox("¿Desea eliminar el registro actual?", vbYesNo + 16, "Continuar")
If respuesta = vbYes Then
    record.Delete adAffectCurrent
    MsgBox "Registro borrado con exito", vbInformation, "Mensaje"
    record.Update
    actualizardata
        Else
            MsgBox "Registro no eliminado", vbCritical, "Mensaje error"
            End If
            escondercuit
'----------------------------- ELIMINAR ---------------------------------------------
'respuesta = MsgBox("¿Desea eliminar el registro actual?", vbYesNo + 16, "Continuar")
'If respuesta = vbYes Then
    's.Delete
    'MsgBox "Registro borrado con exito", 16, "Concretado"
    's.MoveNext
    'End If
'If s.EOF Then
    's.MoveLast
End Sub
Private Sub Command4_Click()
'--------------------------------------MOVERSE AL ANTERIOR REGISTRO--------------------------
record.MovePrevious 'Me Muevo al anterior registro
If record.BOF Then  'Si llego al comienzo entonces
record.MoveLast     'Voy al último o podria hacer que se quede en el primero
mostrar             'Carga los registros en cada objeto(textbox)
Else
mostrar
End If
escondercuit
's.MovePrevious
'If s.BOF Then
's.MoveFirst
'End If
'Text1.Text = s.Fields("cuit")
'Text2.Text = s.Fields("apellido")
'Text3.Text = s.Fields("nombre")
'Text10.Text = s.Fields("fechanacimiento")
'Text5.Text = s.Fields("domicilio")
'Text6.Text = s.Fields("telefono")
End Sub

Private Sub Command5_Click()
'--------------------------------------------------NUEVO REGISTRO-----------------------------------
record.AddNew
limpiar
Text1.MaxLength = 11
'Text1.Text = ""
'Text2.Text = ""
'Text3.Text = ""
'Text4.Text = ""
'Text5.Text = ""
'Text6.Text = ""
'Text7.Text = ""
'Text8.Text = ""
'Text10.Text = ""
Text7.Visible = True
Text8.Visible = True
Command6.Enabled = True
Command5.Enabled = False
End Sub

Private Sub Command6_Click()
'Boton Guardar
Dim dia, mes, anio, fnac, dos, ocho, uno, cuit As String
dia = Text4.Text
mes = Text9.Text
anio = Text10.Text
fnac = dia & "/" & mes & "/" & anio
dos = Text7.Text
ocho = Text1.Text
uno = Text8.Text
cuit = dos + ocho + uno
'4 dia
'9 mes
'10 año
If Text2.Text <> "" And Text3.Text <> "" And Text4.Text <> "" And Text5.Text <> "" And cuit <> "" And fnac <> "" And Text6.Text <> "" And Text9.Text <> "" And Text10.Text <> "" And Text7.Text <> "" And Text8.Text <> "" Then
    '-----------------------------CARGAR REGISTRO CONDICION------------------------------
    record.Fields("cuit").Value = cuit
    record.Fields("apellido").Value = Text2.Text
    record.Fields("nombre").Value = Text3.Text
    record.Fields("fechanacimiento").Value = fnac
    record.Fields("domicilio").Value = Text5.Text
    record.Fields("telefono").Value = Text6.Text
    MsgBox "Registro cargado con exito", 32, "Correctamente"
    record.Update
    Text10.MaxLength = 10
    's.AddNew
    's("cuit") = cuit
    's("apellido") = Text2.Text
    's("nombre") = Text3.Text
    's("fechanacimiento") = fnac
    's("domicilio") = Text5.Text
    's("telefono") = Text6.Text
    's.Update
    'MsgBox "Registro cargado con exito", 32, "Correctamente"
    's.MoveFirst
    Command6.Enabled = False
    Command5.Enabled = True
    'Text10.MaxLength = 10
    escondercuit
    Else
        MsgBox "Faltan rellenar campos", vbCritical, "Mensaje"
    End If
    Text4.Text = ""
    Text9.Text = ""
    
End Sub

Private Sub Command7_Click()
'--------------------------------------MOVERSE AL SIGUIENTE REGISTRO

record.MoveNext     'Me Muevo al siguiente
If record.EOF Then  'Si llego al final entonces
record.MoveFirst    'Vuelvo al primero o podria hacer que se quede en el último
mostrar             'Carga los registros en cada objeto(textbox)
Else
mostrar
End If

escondercuit
's.MoveNext
'If s.EOF Then
's.MoveLast
'End If
'Text1.Text = s.Fields("cuit")
'Text2.Text = s.Fields("apellido")
'Text3.Text = s.Fields("nombre")
'Text10.Text = s.Fields("fechanacimiento")
'Text5.Text = s.Fields("domicilio")
'Text6.Text = s.Fields("telefono")
End Sub

Private Sub Command8_Click()
'------------------------------MOVERSE AL ÚLTIMO REGISTRO----------------------
record.MoveLast  'Muevo al último registro
mostrar          'Llamo al procedimiento para mostrar los registros en las textbox

escondercuit

's.MoveLast
'Text1.Text = s.Fields("cuit")
'Text2.Text = s.Fields("apellido")
'Text3.Text = s.Fields("nombre")
'Text10.Text = s.Fields("fechanacimiento")
'Text5.Text = s.Fields("domicilio")
'Text6.Text = s.Fields("telefono")
End Sub

Private Sub Command9_Click()
'------------------------------MOVERSE AL PRIMER REGISTRO----------------------
record.MoveFirst 'Muevo al primero
mostrar          'Llamo al procedimiento para mostrar los registros en las textbox

escondercuit

's.MoveFirst
'Text1.Text = s.Fields("cuit")
'Text2.Text = s.Fields("apellido")
'Text3.Text = s.Fields("nombre")
'Text10.Text = s.Fields("fechanacimiento")
'Text5.Text = s.Fields("domicilio")
'Text6.Text = s.Fields("telefono")
End Sub

Private Sub Form_Load()
connection.Open "provider=Microsoft.JET.OLEDB.4.0;data source=" & App.Path & "\Ventaautos.mdb" & ""
record.Open "select * from cliente", connection, adOpenDynamic, adLockPessimistic
Text10.MaxLength = 10
Text1.MaxLength = 11
Text7.Visible = False
Text8.Visible = False
mostrar
'Set s = New ADODB.Recordset
'datos.Open "provider=Microsoft.JET.OLEDB.4.0;data source=C:\Users\Mona\Desktop\Consecionaria\Ventaautos.mdb"
's.Source = "cliente"
's.CursorType = adOpenKeyset
's.LockType = adLockOptimistic
's.Open "select * from cliente", datos
's.MoveFirst
    
    'Text1.Text = s.Fields("cuit")
    'Text2.Text = s.Fields("apellido")
    'Text3.Text = s.Fields("nombre")
    'Text10.Text = s.Fields("fechanacimiento")
    'Text5.Text = s.Fields("domicilio")
    'Text6.Text = s.Fields("telefono")
    
End Sub

Sub reiniciar()
record.Close
record.Open "Select * from cliente", connection, adOpenDynamic, adLockPessimistic
End Sub
Sub actualizardata()
record.Close
record.Open "Select * from cliente", connection, adOpenDynamic, adLockPessimistic
If Not record.EOF Then
record.MoveNext
mostrar
Else
MsgBox "Cierre la aplicacion y vuelva a abrirla", 16, "Mensaje"
End If

End Sub
Sub limpiar()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
End Sub
Sub mostrar()
'-------------------------PROCEDIMIENTO QUE TIENE CARGADO PARA MOSTRAR LOS REGISTRO EN CADA CAJA DE TEXTO
Text1.Text = record!cuit
Text2.Text = record!apellido
Text3.Text = record!nombre
Text10.Text = record!fechanacimiento
Text5.Text = record!domicilio
Text6.Text = record!telefono
escondercuit
End Sub

Private Sub Image9_Click()
Me.Hide
End Sub
Sub escondercuit()
    Text7.Visible = False
    Text8.Visible = False
End Sub
