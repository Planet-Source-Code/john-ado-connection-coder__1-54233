VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmConectionCoder 
   BackColor       =   &H00E0E0E0&
   Caption         =   "ADO Connection Coder"
   ClientHeight    =   5325
   ClientLeft      =   1740
   ClientTop       =   1245
   ClientWidth     =   8985
   Icon            =   "frmADOConectionCoder.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   8985
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   4950
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Load"
            Object.ToolTipText     =   "Open a database"
            Object.Tag             =   "Load"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Make"
            Object.ToolTipText     =   "Make the file"
            Object.Tag             =   "Make"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save as text"
            Object.Tag             =   "Save"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy to Clip Board"
            Object.Tag             =   "Copy"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste clipboard to text"
            Object.Tag             =   "Paste"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut selected text"
            Object.Tag             =   "Cut"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Close Program"
            Object.Tag             =   "Exit"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.PictureBox Picture1 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   16
         Top             =   0
         Width           =   0
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8370
      Top             =   2130
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmADOConectionCoder.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmADOConectionCoder.frx":0984
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmADOConectionCoder.frx":0A96
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmADOConectionCoder.frx":0FD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmADOConectionCoder.frx":132A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmADOConectionCoder.frx":186C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmADOConectionCoder.frx":1DAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmADOConectionCoder.frx":2140
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmADOConectionCoder.frx":24D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmADOConectionCoder.frx":25E4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1665
      Left            =   75
      TabIndex        =   1
      Top             =   360
      Width           =   8745
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   6195
         TabIndex        =   11
         Top             =   1215
         Width           =   2355
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00F3F3F3&
         Height          =   285
         Left            =   1530
         TabIndex        =   5
         Top             =   285
         Width           =   4545
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1530
         TabIndex        =   4
         Top             =   600
         Width           =   4545
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1530
         TabIndex        =   3
         Top             =   915
         Width           =   4545
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1545
         TabIndex        =   2
         Top             =   1245
         Width           =   4545
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ADO Connection Coder Ver 1.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   570
         Left            =   6345
         TabIndex        =   14
         Top             =   315
         Width           =   2115
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ADO Connection Coder Ver 1.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   570
         Left            =   6330
         TabIndex        =   13
         Top             =   300
         Width           =   2115
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Table"
         Height          =   210
         Left            =   6225
         TabIndex        =   10
         Top             =   930
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Database Name:"
         Height          =   240
         Left            =   75
         TabIndex        =   9
         Top             =   315
         Width           =   1365
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Proceedure Name:"
         Height          =   240
         Left            =   75
         TabIndex        =   8
         Top             =   615
         Width           =   1365
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Full Address:"
         Height          =   240
         Left            =   75
         TabIndex        =   7
         Top             =   930
         Width           =   1365
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Table Name:"
         Height          =   240
         Left            =   90
         TabIndex        =   6
         Top             =   1245
         Width           =   1365
      End
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Left            =   150
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   2070
      Width           =   8715
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8430
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   17
      Top             =   0
      Width           =   0
   End
   Begin VB.Menu mnuOpenDatabase 
      Caption         =   "&File"
      Begin VB.Menu mnuLoadDatabase 
         Caption         =   "Load Database"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuSaveCode 
         Caption         =   "Save Code"
         Shortcut        =   ^S
      End
      Begin VB.Menu MnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu MnuCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
         Shortcut        =   ^U
      End
   End
   Begin VB.Menu mnuCode 
      Caption         =   "&Code"
      Begin VB.Menu mnuMakeCode 
         Caption         =   "Make Code"
         Shortcut        =   ^M
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmConectionCoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************
'* Module      : frmCodeGen
'* Project     : ADOCodeGen
'* Created     : 05/06/2004 09:48
'* Author      : John Attfield
'* Last Update : 05/06/2004 09:48
'* Purpose     : To generate the code to set the database
'*             : connection and create the RecordSet.
'*             : Error detection has been kept to a
'*             : minimum to allow visual basic to report
'*             : any bugs
'* Thanks      : Thanks to CarlosVara for the idea,
'*             : and the method of geting table names from
'*             : a database.
'*********************************************************

Option Explicit

Dim cn As ADODB.Connection
Dim strFileName As String              'Database Path
Dim strFullPath As String
Dim strProceedureName As String
Dim strProceedure As String
Dim strTableName As String
Dim strDateTime As String

'*********************************************************
'* Procedure   : GetDBPath
'* Created     : 05/06/2004 09:50
'* -------------------------------------------------------
'* Notes       : Get the path to the database using the
'*             : commonDialog control.
'*********************************************************

Private Sub GetDBPath()
    
    On Error GoTo ErrGetPath
    ' Set CancelError is True
    CommonDialog1.CancelError = True
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly
    ' Set filters
    CommonDialog1.Filter = "All Files(*.*)|*.*|Access 2000 DB (*.mdb)|*.mdb"
    ' Specify default filter
    CommonDialog1.FilterIndex = 2
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Display name of selected file
    
    strFileName = CommonDialog1.FileTitle
    strFullPath = CommonDialog1.FileName
    strProceedureName = Replace(strFileName, " ", "")
    strProceedureName = Left(strProceedureName, InStr(strProceedureName, ".") - 1)
    strProceedure = strProceedureName
    Exit Sub
    
ErrGetPath:
    'User pressed the Cancel button
    strFullPath = vbNullString
    Exit Sub
    
End Sub

'*********************************************************
'* Procedure   : cmdExit_Click
'* Created     : 05/06/2004 09:51
'* -------------------------------------------------------
'* Notes       : Just as it says on the tin.
'*********************************************************

Private Sub cmdExit_Click()
    End
End Sub

'*********************************************************
'* Procedure   : Combo1_Click
'* Created     : 05/06/2004 21:13
'* -------------------------------------------------------
'* Notes       :
'*********************************************************

Private Sub Combo1_Click()
    Text4.Text = Combo1.Text
End Sub

'*********************************************************
'* Procedure   : Command1_Click
'* Created     : 05/06/2004 09:53
'* -------------------------------------------------------
'* Notes       : Calls the CommonDialog routine and fills
'*             : the on screen text boxes with the results.
'*             : these will be used later to generate the
'*             : connection code.
'*********************************************************

Private Sub Make()
    Call GetDBPath
    
    Text1.Text = strFileName
    Text2.Text = strProceedureName & "()"
    Text3.Text = strFullPath
    Text4.Text = ""
    
    Call Connect
    
End Sub

'*********************************************************
'* Procedure   : Connect
'* Created     : 05/06/2004 10:51
'* -------------------------------------------------------
'* Notes       :
'*********************************************************

Private Sub Connect()
    
    Dim dbFile As String
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim Ctl As ADOX.Catalog
    Dim CtlTbl As ADOX.Table
    Dim i As Long
    
    '------- Set the database Application Path ------
    
    dbFile = strFullPath
    
    '----------- Establish the connection -----------
    
    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseClient
    cn.ConnectionString = "Provider=Microsoft.jet.OLEDB.4.0;" & "Data Source=" & dbFile & ";" & "Persist Security Info=False"
    
    '-------------- Open the connection -------------
    
    cn.Open
    
    '----------- Open the Database Catalog ----------
    Set Ctl = New ADOX.Catalog
    Ctl.ActiveConnection = cn
    
    'Table Definitions
    i = 0
    For Each CtlTbl In Ctl.Tables
        
        If CtlTbl.Type = "TABLE" Then
            i = i + 1
            If i = 1 Then
                Combo1.Text = CtlTbl.Name
            End If
            Combo1.AddItem CtlTbl.Name
        End If
    Next CtlTbl
    Text4.Text = Combo1.Text
    '------------ Close the Connection --------------
    
    cn.Close
    Set cn = Nothing
    
End Sub

'*********************************************************
'* Procedure   : Form_Unload
'* Created     : 05/06/2004 21:14
'* -------------------------------------------------------
'* Notes       :
'*********************************************************

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

'*********************************************************
'* Procedure   : mnuAbout_Click
'* Created     : 05/06/2004 21:14
'* -------------------------------------------------------
'* Notes       :
'*********************************************************

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

'*********************************************************
'* Procedure   : MnuCopy_Click
'* Created     : 05/06/2004 21:14
'* -------------------------------------------------------
'* Notes       : In the event that nothing is selected
'*             : then all is selected provided there is
'*             : something in the Clipboard
'*********************************************************

Private Sub MnuCopy_Click()
    If Len(Text5.Text) > 0 Then
        If Text5.SelText <> "" Then
            Clipboard.Clear
            Clipboard.SetText Text5.SelText, vbCFText
            StatusBar1.Panels(1).Text = " Selected text in 'ClipBoard'   "
        Else
            Clipboard.SetText Text5.Text, vbCFText
            StatusBar1.Panels(1).Text = " All placed in 'ClipBoard'  "
        End If
    End If
    
End Sub

'*********************************************************
'* Procedure   : mnuCut_Click
'* Created     : 05/06/2004 23:02
'* -------------------------------------------------------
'* Notes       :
'*********************************************************

Private Sub mnuCut_Click()
    If Text5.SelText <> "" Then
        Clipboard.Clear
        Clipboard.SetText Text5.SelText, vbCFText
        Text5.SelText = "" 'Same as copy except this clears the text from the text box'
    End If
End Sub

'*********************************************************
'* Procedure   : mnuPaste_Click
'* Created     : 05/06/2004 23:02
'* -------------------------------------------------------
'* Notes       :
'*********************************************************

Private Sub mnuPaste_Click()
    Text5.SelText = Clipboard.GetText(vbCFText) 'Recovers any text currently in the clipboard'
End Sub

'*********************************************************
'* Procedure   : MnuExit_Click
'* Created     : 05/06/2004 21:14
'* -------------------------------------------------------
'* Notes       :
'*********************************************************

Private Sub MnuExit_Click()
    End
End Sub

'*********************************************************
'* Procedure   : mnuMakeCode_Click
'* Created     : 05/06/2004 15:31
'* -------------------------------------------------------
'* Notes       : This loads the main text box with the
'*             : generated code ready for copying to your
'*             : project. Note the use of a textBox or
'*             : (a string) repeatedly added to).
'*             : Using the " _" character used to break up
'*             : lines of text is limited to about 25
'*             : lines.
'*********************************************************

Private Sub mnuMakeCode_Click()
    
    If Len(Text1.Text) = 0 Then
        Select Case MsgBox("Database not selected" _
            & vbCrLf & "Would you like to select one now ?" _
            , vbYesNo + vbQuestion + vbDefaultButton1, App.Title)
            
        Case vbYes
            Call mnuLoadDatabase_Click
        Case vbNo
            Exit Sub
    End Select
    Exit Sub
End If

strFileName = Text1.Text
strProceedureName = Text2.Text
strFullPath = Text3.Text
strTableName = Text4.Text
strDateTime = Now

Text5.Text = ""

Text5.Text = Text5.Text & "'*********************************************************" & vbCrLf
Text5.Text = Text5.Text & "'* Procedure   : " & strProceedureName & vbCrLf
Text5.Text = Text5.Text & "'* Created     : " & strDateTime & vbCrLf
Text5.Text = Text5.Text & "'* Generator   : ADOCodeGen" & vbCrLf
Text5.Text = Text5.Text & "'* Called From :" & vbCrLf
Text5.Text = Text5.Text & "'* Notes       : Add Microsoft ADO.ext for DLL and" & vbCrLf
Text5.Text = Text5.Text & "'*             : Security" & vbCrLf
Text5.Text = Text5.Text & "'*             : Add Microsoft ActiveX Data Objects" & vbCrLf
Text5.Text = Text5.Text & "'*             : 2.5 Library" & vbCrLf
Text5.Text = Text5.Text & "'*********************************************************" & vbCrLf
Text5.Text = Text5.Text & "Private Sub " & strProceedureName & vbCrLf
Text5.Text = Text5.Text & vbCrLf
Text5.Text = Text5.Text & Chr(9) & "Dim dbFile as String" & vbCrLf
Text5.Text = Text5.Text & Chr(9) & "Dim cn" & strProceedure & " As ADODB.Connection" & vbCrLf
Text5.Text = Text5.Text & Chr(9) & "Dim rs" & strTableName & " As ADODB.Recordset" & vbCrLf
Text5.Text = Text5.Text & Chr(9) & "Dim SQL as String" & vbCrLf
Text5.Text = Text5.Text & vbCrLf
Text5.Text = Text5.Text & Chr(9) & "'------- Set the database Application Path ------" & vbCrLf
Text5.Text = Text5.Text & vbCrLf
Text5.Text = Text5.Text & Chr(9) & "DbFile = App.Path " & "& " & Chr(34) & "\" & strFileName & Chr(34) & vbCrLf
Text5.Text = Text5.Text & vbCrLf
Text5.Text = Text5.Text & Chr(9) & "'----------- Establish the connection -----------" & vbCrLf
Text5.Text = Text5.Text & Chr(9) & vbCrLf
Text5.Text = Text5.Text & Chr(9) & "Set cn" & strProceedure & " = New ADODB.Connection" & vbCrLf
Text5.Text = Text5.Text & Chr(9) & "cn" & strProceedure & ".CursorLocation = adUseClient" & vbCrLf
Text5.Text = Text5.Text & Chr(9) & "cn" & strProceedure & ".ConnectionString = " & Chr(34) & "Provider=Microsoft.jet.OLEDB.4.0;" & Chr(34) & " & " & Chr(34) & "Data Source=" & Chr(34) & " & DbFile & " & Chr(34) & ";" & Chr(34) & " & " & Chr(34) & "Persist Security Info=False" & Chr(34) & vbCrLf
Text5.Text = Text5.Text & vbCrLf
Text5.Text = Text5.Text & Chr(9) & "'-------------- Open the connection -------------" & vbCrLf
Text5.Text = Text5.Text & vbCrLf
Text5.Text = Text5.Text & Chr(9) & "cn" & strProceedure & ".Open" & vbCrLf
Text5.Text = Text5.Text & vbCrLf
Text5.Text = Text5.Text & Chr(9) & "'---------- Enter your SQL Statement Here -------" & vbCrLf
Text5.Text = Text5.Text & vbCrLf
If Len(strTableName) > 0 Then
    Text5.Text = Text5.Text & Chr(9) & "SQL = " & Chr(34) & "SELECT " & strTableName & ".* " & Chr(34) & vbCrLf
    Text5.Text = Text5.Text & Chr(9) & "SQL = SQL & " & Chr(34) & "FROM " & strTableName & ";" & Chr(34) & vbCrLf
End If
Text5.Text = Text5.Text & vbCrLf
Text5.Text = Text5.Text & Chr(9) & "'--------------- Get the Records ----------------" & vbCrLf
Text5.Text = Text5.Text & vbCrLf
Text5.Text = Text5.Text & Chr(9) & "Set rs" & strTableName & " = New ADODB.Recordset" & vbCrLf
Text5.Text = Text5.Text & Chr(9) & "rs" & strTableName & ".Open SQL, " & "cn" & strProceedure & ", adOpenStatic, adLockOptimistic, adCmdText" & vbCrLf
Text5.Text = Text5.Text & vbCrLf
Text5.Text = Text5.Text & Chr(9) & "'----- Carry out functions on the database  -----" & vbCrLf
Text5.Text = Text5.Text & vbCrLf
Text5.Text = Text5.Text & vbCrLf
Text5.Text = Text5.Text & Chr(9) & "'------------ Close the Connection --------------" & vbCrLf
Text5.Text = Text5.Text & vbCrLf
Text5.Text = Text5.Text & Chr(9) & "'cn" & strProceedure & ".Close" & vbCrLf
Text5.Text = Text5.Text & Chr(9) & "'Set cn" & strProceedure & " = Nothing" & vbCrLf
Text5.Text = Text5.Text & vbCrLf
Text5.Text = Text5.Text & "End Sub" & vbCrLf

End Sub

'*********************************************************
'* Procedure   : Form_Resize
'* Created     : 05/06/2004 21:21
'* -------------------------------------------------------
'* Notes       :
'*********************************************************

Private Sub Form_Resize()
    
    'Check to see if the form has been minimized
    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If
    
    If Me.Width < 6000 Or Me.Height < 4000 Then
        Me.Width = 6000
        Me.Height = 4000
    End If
    
    Frame1.Top = 300
    Frame1.Left = 100
    Frame1.Width = Me.Width - 300
    
    Text5.Top = Frame1.Top + Frame1.Height + 100
    Text5.Left = Frame1.Left
    Text5.Width = Frame1.Width
    Text5.Height = Me.Height - Text5.Top - 1200
    
    Text1.Width = Frame1.Width - 4545
    Text2.Width = Text1.Width
    Text3.Width = Text1.Width
    Text4.Width = Text1.Width
    
    Label5.Left = Text1.Left + Text1.Width + 100
    Combo1.Left = Label5.Left
    Combo1.Width = (Frame1.Width - Text1.Width) - 1750
    
    Label7.Left = Combo1.Left + 430
    Label6.Left = Combo1.Left + 400
End Sub

'*********************************************************
'* Procedure   : mnuLoadDatabase_Click
'* Created     : 05/06/2004 21:21
'* -------------------------------------------------------
'* Notes       :
'*********************************************************

Private Sub mnuLoadDatabase_Click()
    Call GetDBPath
    
    Text1.Text = strFileName
    Text2.Text = strProceedureName & "()"
    Text3.Text = strFullPath
    Text4.Text = ""
    
    Call Connect
End Sub

Private Sub mnuPast_Click()
    
End Sub

'*********************************************************
'* Procedure   : mnuSaveCode_Click
'* Created     : 05/06/2004 21:21
'* -------------------------------------------------------
'* Notes       :
'*********************************************************

Private Sub mnuSaveCode_Click()
    Dim lngSaveFile As Long
    Dim strFileName As String
    
    Dim temp As String
    
    strFileName = strProceedure & "Code.txt"
    lngSaveFile = FreeFile
    
    Open AppPath() & strFileName For Output As lngSaveFile
    Print #lngSaveFile, Text5.Text
    Close #lngSaveFile
    
End Sub
'*********************************************************
'* Procedure   : AppPath
'* Created     : 05/06/2004 21:22
'* -------------------------------------------------------
'* Notes       :
'*********************************************************
Public Function AppPath() As String
    
    If Right$(App.Path, 1) <> "\" Then
        AppPath = App.Path & "\"
    Else 'NOT RIGHT$(APP.PATH,...
        AppPath = App.Path
    End If
    
End Function

'*********************************************************
'* Procedure   : Toolbar1_ButtonClick
'* Created     : 05/06/2004 21:22
'* -------------------------------------------------------
'* Notes       :
'*********************************************************

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Tag
        Case "Load"
            Call mnuLoadDatabase_Click
        Case "Make"
            Call mnuMakeCode_Click
        Case "Save"
            Call mnuSaveCode_Click
        Case "Copy"
            Call MnuCopy_Click
        Case "Paste"
            Call mnuPaste_Click
        Case "Cut"
            Call mnuCut_Click
        Case "Help"
            Call mnuAbout_Click
        Case "Exit"
            Call MnuExit_Click
    End Select
    
End Sub

