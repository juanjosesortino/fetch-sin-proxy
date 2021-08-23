VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   5610
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text 
      Height          =   525
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Text            =   "Provider=OraOLEDB.Oracle.1;Password=apfrms2001;User ID=SYSADMIN;Data Source=BASE"
      Top             =   60
      Width           =   5415
   End
   Begin VB.TextBox Text 
      Height          =   3045
      Index           =   0
      Left            =   1470
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   870
      Width           =   3135
   End
   Begin VB.CommandButton Command 
      Caption         =   "Fetch"
      Height          =   525
      Left            =   150
      TabIndex        =   2
      Top             =   900
      Width           =   1245
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command_Click()
      Dim rst  As ADODB.Recordset
         
10       On Error GoTo GestErr
   
20       Set cnn = New ADODB.Connection 'HKEY_LOCAL_MACHINE\SOFTWARE\Algoritmo\DatabaseSettings\ConnectionsStrings
30       cnn.ConnectionString = Text(1).Text
40       cnn.Open
50       Set rst = New ADODB.Recordset
60       rst.CursorLocation = adUseClient
70       rst.LockType = adLockReadOnly
80       rst.CursorType = adOpenStatic

90       SQL = "SELECT PRV_PROVINCIA, PRV_NOMBRE, PRV_CODIGO_AFIP, PRV_CODIGO_SAGPYA , PRV_CODIGO_PAIS FROM PROVINCIAS"
100      rst.Open SQL, cnn

110      Do While Not rst.EOF
120         Text(0).Text = Text(0).Text & rst("PRV_NOMBRE").Value & vbCrLf
130         rst.MoveNext
140      Loop
150      rst.Close
   
160      Exit Sub
   
GestErr:
170      MsgBox Err.Description & Erl

End Sub

