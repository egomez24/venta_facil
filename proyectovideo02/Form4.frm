VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00404040&
   Caption         =   "Pantalla de inicio"
   ClientHeight    =   3900
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7875
   ClipControls    =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   7875
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ingresar"
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1920
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   960
      Width           =   3735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "RapiFacil v1.0"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   2880
      TabIndex        =   6
      Top             =   120
      Width           =   1860
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Left            =   1440
      TabIndex        =   2
      Top             =   1560
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
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
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   720
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnnconexion03 As ADODB.Connection
Private registron03 As ADODB.Recordset
Private cmdcomando03 As ADODB.Command

Private Sub consultaxf01(consultax02 As String)
Set cnnconexion03 = New ADODB.Connection
Set cmdcomando03 = New ADODB.Command

cnnconexion03.ConnectionString = "Provider=sqloledb.1;server=DESKTOP-15TA2GK\SQLEXPRESS;uid=sa;pwd=hola27;database=Inventario;"
cnnconexion03.CursorLocation = adUseClient
cnnconexion03.ConnectionTimeout = 15
cnnconexion03.Open

With cmdcomando03
.ActiveConnection = cnnconexion03
.CommandType = adCmdText
.CommandTimeout = 15
.CommandText = consultax02
End With

Set registron03 = cmdcomando03.Execute()



End Sub

Private Sub Command1_Click()
Dim consultax03 As String
consultax03 = "Select usuario,pwd From usuarios where usuario='" + Text1.Text + "'" + "and pwd='" + Text2.Text + "'"
consultaxf01 (consultax03)
If registron03.RecordCount = 0 Then
MsgBox "Datos incorrectos"
Else
Unload Me
Form1.Show
Form1.Label2.Caption = registron03!usuario
End If

End Sub

Private Sub Form_Load()
Text1.Text = "admin"
Text2.Text = "123"
End Sub

