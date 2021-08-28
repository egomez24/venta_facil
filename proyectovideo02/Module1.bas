Attribute VB_Name = "Module1"
Option Explicit

Private cnnconexionA01 As ADODB.Connection
Public registronA01 As ADODB.Recordset
Private cmdcomandoA01 As ADODB.Command


Public Sub consulta(clocal01 As String)
Set cnnconexionA01 = New ADODB.Connection
cnnconexionA01.ConnectionString = "provider=sqloledb.1;server=DESKTOP-15TA2GK\SQLEXPRESS;uid=sa;pwd=hola27;database=Inventario;"
cnnconexionA01.CursorLocation = adUseClient
cnnconexionA01.ConnectionTimeout = 15
cnnconexionA01.Open
Set cmdcomandoA01 = New ADODB.Command
With cmdcomandoA01
.ActiveConnection = cnnconexionA01
.CommandType = adCmdText
.CommandTimeout = 15
.CommandText = clocal01
End With
Set registronA01 = cmdcomandoA01.Execute()

End Sub

