VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2310
      Left            =   210
      TabIndex        =   1
      Top             =   255
      Width           =   5670
      _ExtentX        =   10001
      _ExtentY        =   4075
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Title"
         Caption         =   "Title"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1995.024
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   495
      Left            =   4680
      TabIndex        =   0
      Top             =   2655
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call Contacts
End Sub
'*********************************************************
'* Procedure   : Contacts()
'* Created     : 07/06/2004 14:39:19
'* Generator   : ADOCodeGen
'* Called From :
'* Notes       : Add Microsoft ADO.ext for DLL and
'*             : Security
'*             : Add Microsoft ActiveX Data Objects
'*             : 2.5 Library
'*********************************************************
Private Sub Contacts()
    
    Dim dbFile As String
    Dim cnContacts As ADODB.Connection
    Dim rsContacts As ADODB.Recordset
    Dim SQL As String
    
    '------- Set the database Application Path ------
    
    dbFile = App.Path & "\Contacts.mdb"
    
    '----------- Establish the connection -----------
    
    Set cnContacts = New ADODB.Connection
    cnContacts.CursorLocation = adUseClient
    cnContacts.ConnectionString = "Provider=Microsoft.jet.OLEDB.4.0;" & "Data Source=" & dbFile & ";" & "Persist Security Info=False"
    
    '-------------- Open the connection -------------
    
    cnContacts.Open
    
    '---------- Enter your SQL Statement Here -------
    
    SQL = "SELECT Contacts.* "
    SQL = SQL & "FROM Contacts;"
    
    '--------------- Get the Records ----------------
    
    Set rsContacts = New ADODB.Recordset
    rsContacts.Open SQL, cnContacts, adOpenStatic, adLockOptimistic, adCmdText
    
    '----- Carry out functions on the database  -----
    Set DataGrid1.DataSource = rsContacts
    DataGrid1.Refresh
    
    '------------ Close the Connection --------------
    
    'cnContacts.Close
    'Set cnContacts = Nothing
    
End Sub
