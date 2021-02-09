Attribute VB_Name = "Module1"
Public dia, mes, anio, fnac As String
Public interes As Single
Public capital As Single
Public tiempo As Single
Public interesmes As Single
Public interesfinalpormes As Single
Public meses As Integer
Public calculo As Single
Public cuotasininteres As Single
Public descuento As Single
Public contado As Single
Public anticipo As Single
Public saldo As Single
Public contra As String
Public rs As Recordset
Public i As Integer
Public respuesta As Integer
Public rs_clientes As Recordset
Public db As Database
Public rs_listado As Recordset
Public bd As Connection
Public rs_ventas As Recordset
Public bdd As Connection
Public cn As New ADODB.Connection
Public tablas As New ADODB.Recordset
Public ruta As String

Public Function conectlistado()
ruta = App.Path & "\ventaautos.mdb"
cn.Open "provider=Microsoft.JET.OLEDB.4.0;data source=" & ruta & ";"
End Function

Public Function login()

End Function









