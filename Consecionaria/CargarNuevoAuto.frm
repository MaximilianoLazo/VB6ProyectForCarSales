VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form CargarNuevoAuto 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Carga"
   ClientHeight    =   8340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13890
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   13890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2400
      TabIndex        =   5
      Top             =   1800
      Width           =   3975
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H0080FFFF&
      Caption         =   "borrar registro actual"
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
      Left            =   6600
      Picture         =   "CargarNuevoAuto.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7320
      Width           =   2895
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   9600
      TabIndex        =   23
      Top             =   6240
      Width           =   3975
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H0080FFFF&
      Caption         =   "Buscar por id_auto"
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
      Left            =   6600
      Picture         =   "CargarNuevoAuto.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6120
      Width           =   2895
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
      Left            =   3480
      Picture         =   "CargarNuevoAuto.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7320
      Width           =   1355
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
      Left            =   360
      Picture         =   "CargarNuevoAuto.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6120
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
      Left            =   5040
      Picture         =   "CargarNuevoAuto.frx":2328
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
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
      Left            =   1920
      Picture         =   "CargarNuevoAuto.frx":2BF2
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6120
      Width           =   1335
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
      Left            =   3480
      Picture         =   "CargarNuevoAuto.frx":34BC
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0080FFFF&
      Caption         =   "Guardar"
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
      Left            =   1920
      Picture         =   "CargarNuevoAuto.frx":3D86
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7320
      Width           =   1335
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
      Left            =   360
      Picture         =   "CargarNuevoAuto.frx":4650
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7320
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog abrir 
      Left            =   12720
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   5280
      Width           =   3975
   End
   Begin VB.TextBox Text5 
      DataSource      =   "Ado"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   3840
      Width           =   3975
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2400
      MaxLength       =   17
      TabIndex        =   0
      Top             =   2400
      Width           =   3975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Examinar"
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
      Left            =   5040
      Picture         =   "CargarNuevoAuto.frx":4F1A
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7320
      Width           =   1355
   End
   Begin VB.TextBox Text6 
      DataSource      =   "Ado"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   4560
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2400
      MaxLength       =   17
      TabIndex        =   6
      Top             =   960
      Width           =   3975
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   3120
      Width           =   3975
   End
   Begin VB.Image Image9 
      Height          =   495
      Left            =   13200
      Picture         =   "CargarNuevoAuto.frx":57E4
      Stretch         =   -1  'True
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Carga de vehículos para la empresa"
      BeginProperty Font 
         Name            =   "Bodoni MT Poster Compressed"
         Size            =   24
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   840
      TabIndex        =   25
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Foto"
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
      TabIndex        =   14
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Image Image8 
      Height          =   480
      Left            =   1680
      Picture         =   "CargarNuevoAuto.frx":60AE
      Top             =   5280
      Width           =   480
   End
   Begin VB.Image Image7 
      Height          =   4815
      Left            =   6720
      Stretch         =   -1  'True
      Top             =   960
      Width           =   6855
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   1680
      Picture         =   "CargarNuevoAuto.frx":6978
      Top             =   960
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   1680
      Picture         =   "CargarNuevoAuto.frx":7A42
      Top             =   1680
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   1680
      Picture         =   "CargarNuevoAuto.frx":8B0C
      Top             =   2280
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   1680
      Picture         =   "CargarNuevoAuto.frx":9BD6
      Top             =   3120
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   1680
      Picture         =   "CargarNuevoAuto.frx":A4A0
      Top             =   3840
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   1680
      Picture         =   "CargarNuevoAuto.frx":AD6A
      Top             =   4560
      Width           =   480
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Precio $"
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
      TabIndex        =   12
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ID_Auto"
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
      TabIndex        =   11
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Categoría"
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
      TabIndex        =   10
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Marca"
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
      TabIndex        =   9
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
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
      TabIndex        =   8
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      TabIndex        =   7
      Top             =   3840
      Width           =   1215
   End
End
Attribute VB_Name = "CargarNuevoAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.connection
Dim re As New ADODB.Recordset

'Option Explicit
'Dim basededatos As New ADODB.Connection
'Private WithEvents r As ADODB.Recordset
'Dim x As String
Private Sub Command10_Click()
'------------------------ BÚSQUEDA---------------------------------
re.Close
re.Open "select * from auto where id_auto= " & Val(Text8.Text) & "", con, adOpenDynamic, adLockPessimistic
If Not re.EOF Then
    mostrar
    reiniciar
Else
MsgBox "Vehículo no encontrado", vbCritical, "Mensaje"
End If
'r.Find "id_auto = " & Val(Text8.Text) & "", 1
'If r.EOF = False And r.BOF = False Then
    'Text1.Text = r.Fields("id_auto")
    'Text2.Text = r.Fields("categoria")
    'Text3.Text = r.Fields("marca")
    'Text4.Text = r.Fields("nombre")
    'Text5.Text = r.Fields("año")
    'Text6.Text = r.Fields("precio")
    'Text7.Text = r.Fields("foto")
    'x = App.Path
    'Image7.Picture = LoadPicture(x + "\" + Text7.Text)
    
    
'Else
    'MsgBox "No se encontró ningún dato", vbCritical, "ERROR"
'End If
End Sub

Private Sub Command11_Click()
'------------------------------- MODIFICAR ----------------------------------------
're.Fields("categoria").Value = Text2.Text
're.Fields("marca").Value = Text3.Text
're.Fields("nombre").Value = Text4.Text
're.Fields("año").Value = Text5.Text
're.Fields("precio").Value = Text6.Text
're.Fields("foto").Value = Text7.Text
'MsgBox "Registro actualizado con exito", vbInformation, "Mensaje"
're.Update
re.Update Array("id_auto", "categoria", "marca", "nombre", "año", "precio", "foto"), Array(Text1.Text, Combo1.Text, Text3.Text, Text4.Text, Text5.Text, Text6.Text, Text7.Text)
're.Update
If re.State = 1 Or re.State = 0 Then
MsgBox "Registro actualizado con exito"
Else
MsgBox "Ha ocurrido un error"
End If
'r.Update Array("id_auto", "categoria", "marca", "nombre", "año", "precio", "foto"), Array(Text1.Text, Text2.Text, Text3.Text, Text4.Text, Text5.Text, Text6.Text, Text7.Text)
'r.Update
'If r.State = 1 Or r.State = 0 Then
'MsgBox "Registro actualizado con exito"
'Else
'MsgBox "Ha ocurrido un error"
'End If
End Sub

Private Sub Command12_Click()
'----------------------------- ELIMINAR ---------------------------------------------
respuesta = MsgBox("¿Desea eliminar el registro actual?", vbYesNo + 16, "Continuar")
If respuesta = vbYes Then
    re.Delete adAffectCurrent
    MsgBox "Registro borrado con exito", vbInformation, "Mensaje"
    re.Update
    actualizardata
        Else
            MsgBox "Registro no eliminado", vbCritical, "Mensaje error"
            End If
            





'respuesta = MsgBox("¿Desea eliminar el registro actual?", vbYesNo + 16, "Continuar")
'If respuesta = vbYes Then
    'r.Delete
    'MsgBox "Registro borrado con exito", 16, "Concretado"
    'r.MoveNext
    'End If
'If r.EOF Then
    'r.MoveLast
   ' End If
End Sub

Private Sub Command2_Click()
abrir.ShowOpen
Image7.Picture = LoadPicture(abrir.FileName)
Text7.Text = abrir.FileTitle
If Text7.Text = "" Then
    MsgBox "Falta imagen", 16, "Error"
    Else
    Text7.Text = abrir.FileTitle
End If
End Sub

Private Sub Command3_Click()
'-----------------------------MOVERSE AL ANTERIOR-------------------------------------
re.MovePrevious
If re.BOF Then
re.MoveLast
mostrar
Else
mostrar
End If
'r.MovePrevious
'If r.BOF Then
'r.MoveFirst
'End If
'Text1.Text = r.Fields("id_auto")
'Text2.Text = r.Fields("categoria")
'Text3.Text = r.Fields("marca")
'Text4.Text = r.Fields("nombre")
'Text5.Text = r.Fields("año")
'Text6.Text = r.Fields("precio")
'Text7.Text = r.Fields("foto")
'x = App.Path
'Image7.Picture = LoadPicture(x & "\" & Text7.Text)

End Sub


Private Sub Command4_Click()
Me.Hide
Form1.Show
End Sub

Private Sub Command5_Click()
'---------------------------NUEVO-------------------------------------
re.AddNew
limpiar
Command11.Enabled = False ' Botón modificar desactivado
'Text2.Enabled = True
'Text3.Enabled = True
'Text4.Enabled = True
'Text5.Enabled = True
'Text6.Enabled = True
'Text7.Enabled = True
'Command6.Enabled = True
'Command2.Enabled = True
'Text1.Text = ""
'Text2.Text = ""
'Text3.Text = ""
'Text4.Text = ""
'Text5.Text = ""
'Text6.Text = ""
'Text7.Text = ""
'Image7.Picture = LoadPicture(Text7.Text)
'Command5.Enabled = False
End Sub

Private Sub Command6_Click()
    '-----------------------------CARGAR REGISTRO------------------------------
    If Combo1.Text <> "" And Text3.Text <> "" And Text4.Text <> "" And Text5.Text <> "" And Text6.Text <> "" And Text7.Text <> "" Then
    re.Fields("categoria").Value = Combo1.Text
    re.Fields("marca").Value = Text3.Text
    re.Fields("nombre").Value = Text4.Text
    re.Fields("año").Value = Text5.Text
    re.Fields("precio").Value = Text6.Text
    re.Fields("foto").Value = Text7.Text
    X = App.Path
    Image7.Picture = LoadPicture(X & "\" & Text7.Text)
    MsgBox "Registro cargado con exito", 32, "Correctamente"
    re.Update
    Command11.Enabled = True ' Luego de cargar un registro activo el boton modificar
    Else
    MsgBox "Faltan rellenar campos", vbCritical, "Mensaje:"
    End If
    'r.AddNew
    'r("categoria") = Text2.Text
    'r("marca") = Text3.Text
    'r("nombre") = Text4.Text
    'r("año") = Text5.Text
    'r("precio") = Text6.Text
    'r("foto") = Text7.Text
    'x = App.Path
    'Image7.Picture = LoadPicture(x & "\" & Text7.Text)
    'r.Update
    'MsgBox "Registro cargado con exito", 32, "Correctamente"
    'r.MoveFirst
    'Command5.Enabled = True
    'Command6.Enabled = False
    'Command2.Enabled = False
End Sub

Private Sub Command7_Click()
'--------------------------------------MOVERSE AL SIGUIENTE REGISTRO

re.MoveNext 'Me Muevo al siguiente
If re.EOF Then 'Si llego al final entonces
re.MoveFirst    'Vuelvo al primero
mostrar
Else
mostrar
End If

'r.MoveNext
'If r.EOF Then
'r.MoveLast
'End If
'Text1.Text = r.Fields("id_auto")
'Text2.Text = r.Fields("categoria")
'Text3.Text = r.Fields("marca")
'Text4.Text = r.Fields("nombre")
'Text5.Text = r.Fields("año")
'Text6.Text = r.Fields("precio")
'Text7.Text = r.Fields("foto")
'x = App.Path
'Image7.Picture = LoadPicture(x & "\" & Text7.Text)
End Sub
Private Sub Command8_Click()
'---------------------------MOVERSE AL ÚLTIMO REGISTRO-----------------------

re.MoveLast
mostrar

'r.MoveLast
'Text1.Text = r.Fields("id_auto")
'Text2.Text = r.Fields("categoria")
'Text3.Text = r.Fields("marca")
'Text4.Text = r.Fields("nombre")
'Text5.Text = r.Fields("año")
'Text6.Text = r.Fields("precio")
'Text7.Text = r.Fields("foto")
'x = App.Path
'Image7.Picture = LoadPicture(x & "\" & Text7.Text)
End Sub

Private Sub Command9_Click()
'------------------------------MOVERSE AL PRIMER REGISTRO----------------------
re.MoveFirst
mostrar

'r.MoveFirst
'Text1.Text = r.Fields("id_auto")
'Text2.Text = r.Fields("categoria")
'Text3.Text = r.Fields("marca")
'Text4.Text = r.Fields("nombre")
'Text5.Text = r.Fields("año")
'Text6.Text = r.Fields("precio")
'Text7.Text = r.Fields("foto")
'x = App.Path
'Image7.Picture = LoadPicture(x & "\" & Text7.Text)
End Sub
Sub reiniciar()
re.Close
re.Open "Select * from auto", con, adOpenDynamic, adLockPessimistic
End Sub
Sub actualizardata()
re.Close
re.Open "Select * from auto", con, adOpenDynamic, adLockPessimistic
If Not re.EOF Then
re.MoveNext
mostrar
Else
MsgBox "Cierre la aplicacion y vuelva a abrirla", 16, "Mensaje"
End If

End Sub
Sub limpiar()
Text1.Text = ""
Combo1.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Image7.Picture = LoadPicture("")
End Sub
Sub mostrar()
'-------------------------PROCEDIMIENTO QUE TIENE CARGADO PARA MOSTRAR LOS REGISTRO EN CADA CAJA DE TEXTO
Text1.Enabled = False
Text1.Text = re!id_auto
Combo1.Text = re!categoria
Text3.Text = re!marca
Text4.Text = re!nombre
Text5.Text = re!año
Text6.Text = re!precio
Text7.Text = re!foto
X = App.Path
Image7.Picture = LoadPicture(X + "\" + Text7.Text)
End Sub
Private Sub Form_Load()
con.Open "provider=Microsoft.JET.OLEDB.4.0;data source=" & App.Path & "\Ventaautos.mdb" & ""
re.Open "select * from auto", con, adOpenDynamic, adLockPessimistic
mostrar
re.Update
Combo1.AddItem "Camioneta"
Combo1.AddItem "Auto"
Combo1.AddItem "SUV"

'Set r = New ADODB.Recordset
'basededatos.Open "provider=Microsoft.JET.OLEDB.4.0;data source=C:\Users\Mona\Desktop\Consecionaria\Ventaautos.mdb"
'r.Source = "auto"
'r.CursorType = adOpenKeyset
'r.LockType = adLockOptimistic
'r.Open "select * from auto", basededatos
'r.MoveFirst
'Text1.Text = r.Fields("id_auto")
'Text2.Text = r.Fields("categoria")
'Text3.Text = r.Fields("marca")
'Text4.Text = r.Fields("nombre")
'Text5.Text = r.Fields("año")
'Text6.Text = r.Fields("precio")
'Text7.Text = r.Fields("foto")
'x = App.Path
'Image7.Picture = LoadPicture(x + "\" + Text7.Text)
'Text1.Enabled = False
'Command6.Enabled = False
'Command2.Enabled = False
End Sub

Private Sub Image9_Click()
'------------------------- CERRAR EL FORMULARIO---------------------
respuesta = MsgBox("¿Desea cerrar la carga de Autos?", vbYesNo + 16, "Continuar")
If respuesta = vbYes Then
Me.Hide
End If
End Sub
