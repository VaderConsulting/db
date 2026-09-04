VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Troll"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   6225
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstDate 
      Height          =   2595
      Left            =   4320
      TabIndex        =   5
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "2"
      Height          =   255
      Left            =   6480
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox lstHostname 
      Height          =   2595
      Left            =   2640
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdGet 
      Caption         =   "Start"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "\\cbdxaai\kixlog$\07-07"
      Top             =   120
      Width           =   2415
   End
   Begin VB.FileListBox filLogs 
      Height          =   2625
      Left            =   120
      Pattern         =   "*.log"
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Idle"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   6015
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExtract_Click()
    Dim adoconn As ADODB.Connection
    Dim adors As ADODB.Recordset
    Dim strConn As String
    Dim vFilename As String
    Dim vFileDateTime As Date
    
    Set adoconn = New ADODB.Connection
    
    strConn = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=EUC;Data Source=PERTHXSAC"
    adoconn.Open strConn
    
    Set adors = New ADODB.Recordset
    adors.CursorType = adOpenKeyset
    adors.LockType = adLockOptimistic
    
    SQL = "CREATE TABLE tblDaveTemp (Hostname nVarChar(50), FileDateTime DateTime, xImportMe Bit, xIP VarChar(50), xSiteloc VarChar(50))"
    
    adoconn.Execute SQL
    adors.Open "tblDaveTemp", adoconn, , , adCmdTable
    
    For lp = 0 To filLogs.ListCount - 1
        vFilename = Left(filLogs.List(lp), Len(filLogs.List(lp)) - 4)
        vFileDateTime = FileDateTime(txtPath & "\" & filLogs.List(lp))
        lstHostname.AddItem vFilename
        lstDate.AddItem vFileDateTime
        adors.AddNew
        adors!HostName = vFilename
        adors!FileDateTime = vFileDateTime
        adors.Update
        lblStatus = lp & " of " & filLogs.ListCount - 1
        lblStatus.Refresh
        DoEvents
    Next lp
    'adoconn.Execute "DROP TABLE tblDaveTemp"
    

    MsgBox "Jobs Done"
End Sub

Private Sub cmdGet_Click()
    filLogs.Path = txtPath
    cmdExtract_Click
End Sub

