VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   Caption         =   "Form1"
   ClientHeight    =   6180
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   13380
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   13380
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   480
      Top             =   4560
   End
   Begin VB.CommandButton Command6 
      Caption         =   "get_value"
      Height          =   495
      Left            =   10680
      TabIndex        =   5
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      Picture         =   "Form1.frx":182C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      MaskColor       =   &H00404040&
      Picture         =   "Form1.frx":3058
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      Picture         =   "Form1.frx":4884
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      MaskColor       =   &H00404040&
      Picture         =   "Form1.frx":60B0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Adobe Fan Heiti Std B"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   10560
      TabIndex        =   8
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Left            =   1080
      TabIndex        =   7
      Top             =   240
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   780
   End
   Begin VB.Image Image1 
      Height          =   4590
      Left            =   3240
      Picture         =   "Form1.frx":78DC
      Stretch         =   -1  'True
      Top             =   120
      Width           =   7260
   End
   Begin VB.Menu fcrud01 
      Caption         =   "Archivo"
      Index           =   0
      Begin VB.Menu bagregar01 
         Caption         =   "Agregar"
         Index           =   1
      End
      Begin VB.Menu beditar01 
         Caption         =   "Editar"
         Index           =   2
      End
      Begin VB.Menu beliminar01 
         Caption         =   "Eliminar"
         Index           =   3
      End
      Begin VB.Menu bbuscar01 
         Caption         =   "Buscar"
         Index           =   4
      End
   End
   Begin VB.Menu fventa01 
      Caption         =   "Venta"
      Index           =   5
      Begin VB.Menu bfactura01 
         Caption         =   "Facturar"
         Index           =   6
      End
   End
   Begin VB.Menu reporte 
      Caption         =   "Reportes"
      Index           =   7
      Begin VB.Menu cierrecaja 
         Caption         =   "Cierre de caja"
         Index           =   8
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Private cnnconexion As ADODB.Connection
Private registron As ADODB.Recordset
Private cmdcomando As ADODB.Command


Private Sub Command1_Click()
Dim consulta03 As String
consulta03 = "select * From Items"
consulta (consulta03)
'Set DataGrid1.DataSource = registron
Set Form5.DataGrid1.DataSource = registronA01
Form5.Show
End Sub

Private Sub consulta01(consulta02 As String)
    Set cnnconexion = New ADODB.Connection
    
    cnnconexion.ConnectionString = "Provider=sqloledb.1;server=DESKTOP-15TA2GK\SQLEXPRESS;uid=sa;pwd=usado123;database=Inventario;"
    cnnconexion.ConnectionTimeout = 15
    cnnconexion.CursorLocation = adUseClient
    cnnconexion.Open
    
    Set cmdcomando = New ADODB.Command
    With cmdcomando
    .ActiveConnection = cnnconexion
    .CommandType = adCmdText
    .CommandTimeout = 15
    .CommandText = consulta02
    End With
    Set registron = cmdcomando.Execute()
    

End Sub

Private Sub Command2_Click()
Dim consulta04 As String
consulta04 = "select * From Items Where Codigo = '" + Text1.Text + "'"
consulta01 (consulta04)
If registron.RecordCount = 0 Then
MsgBox "No existe"
Else
Set DataGrid1.DataSource = registron
End If


End Sub

Private Sub Command3_Click()
Form3.Show

End Sub

Private Sub Command4_Click()
Form6.Show

'Dim consulta05 As String
'consulta05 = "Delete From Items where Codigo = '" + Text1.Text + "'"
'consulta01 (consulta05)
'MsgBox "Se ha eliminado"
End Sub

Private Sub Command5_Click()

Dim campoabuscar As String
Dim consulta06 As String
campoabuscar = InputBox("Ingrese registro")
If campoabuscar <> "" Then


    consulta06 = "Select * From Items where Codigo = '" + campoabuscar + "'"
    consulta01 (consulta06)

    If registron.RecordCount = 0 Then
        MsgBox "No Existe"
    Else
       Form2.Text1.Text = registron!Codigo
        Form2.Text2.Text = registron!Nombre
        Form2.Text3.Text = Trim(registron!Precio)
        Form2.Text4.Text = registron!Descuento
        Form2.Show
    End If
End If

End Sub

Private Sub Command6_Click()
'varBookMark = DataGrid1.RowBookmark(0)
'MsgBox DataGrid1.Columns(1).CellValue(varBookMark)
'MsgBox DataGrid1.ApproxCount
Dim j As Integer
For j = 0 To (DataGrid1.ApproxCount - 1) Step 1
'MsgBox j
varBookMark = DataGrid1.RowBookmark(j)
MsgBox CStr(DataGrid1.Columns(3).CellValue(varBookMark))
Next

End Sub

Private Sub Timer1_Timer()
Label3.Caption = Now()
End Sub
