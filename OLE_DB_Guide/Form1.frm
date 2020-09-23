VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Complete OLE DB Provider Guide"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10800
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   10800
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   6015
      Left            =   4755
      TabIndex        =   3
      Top             =   45
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   10610
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":058A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   10935
      TabIndex        =   2
      Top             =   0
      Width           =   10935
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   10815
      TabIndex        =   1
      Top             =   15
      Width           =   10815
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2760
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":064C
            Key             =   "book"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0BE6
            Key             =   "page"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6015
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   10610
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   176
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "Form1.frx":1180
   End
   Begin VB.Menu mnuF 
      Caption         =   "&File"
      Begin VB.Menu mniAOT 
         Caption         =   "&Form Always On Top"
         Checked         =   -1  'True
      End
      Begin VB.Menu blnk2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEx 
         Caption         =   "&Exit       "
      End
   End
   Begin VB.Menu mnuHlp 
      Caption         =   "&Help"
      Begin VB.Menu mnucon 
         Caption         =   "Contents..."
      End
      Begin VB.Menu blnk1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbt 
         Caption         =   "About...       "
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" _
            (ByVal hwnd As Long, _
            ByVal hWndInsertAfter As Long, _
            ByVal X As Long, _
            ByVal Y As Long, _
            ByVal cX As Long, _
            ByVal cY As Long, _
            ByVal wFlags As Long) As Long
            
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
            
Public Function getURL(urlADD As String, sourceHWND As String)
On Error Resume Next
Dim gotoURL
gotoURL = ShellExecute(sourceHWND, vbNullString, urlADD, "", vbNullString, 1)
End Function

Private Sub Form_Load()
Call SetTopWindow(Me.hwnd, True)
Dim X As Node
Set X = TreeView1.Nodes.Add(, , "OlePro", "OLE DB Provider", 1, 1)
X.Tag = "OlePro"
X.Expanded = False
'----- 1-5 child
Set X = TreeView1.Nodes.Add("OlePro", tvwChild, "actvS", "for Active Directory Service", 2, 2)
X.Tag = "actvS"
Set X = TreeView1.Nodes.Add("OlePro", tvwChild, "adVTG", "for Advantage", 2, 2)
X.Tag = "adVTG"
Set X = TreeView1.Nodes.Add("OlePro", tvwChild, "as/400", "for for AS/400 (from IBM)", 2, 2)
X.Tag = "as/400"
Set X = TreeView1.Nodes.Add("OlePro", tvwChild, "vSam", "for AS/400 and VSAM (from Microsoft)", 2, 2)
X.Tag = "vSam"
Set X = TreeView1.Nodes.Add("OlePro", tvwChild, "comServer", "for Commerce Server", 2, 2)
X.Tag = "comServer"
'----- 6-10 child
Set X = TreeView1.Nodes.Add("OlePro", tvwChild, "db2MS", "for DB2 (from Microsoft)", 2, 2)
X.Tag = "db2MS"
Set X = TreeView1.Nodes.Add("OlePro", tvwChild, "dtsPackgs", "for DTS Packages", 2, 2)
X.Tag = "dtsPackgs"
Set X = TreeView1.Nodes.Add("OlePro", tvwChild, "xChange", "for Exchange", 2, 2)
X.Tag = "xChange"
Set X = TreeView1.Nodes.Add("OlePro", tvwChild, "xCell", "for Excel", 2, 2)
X.Tag = "xCell"
Set X = TreeView1.Nodes.Add("OlePro", tvwChild, "inServer", "for Index Server", 2, 2)
X.Tag = "inServer"
'----- 11-15 child
Set X = TreeView1.Nodes.Add("OlePro", tvwChild, "netPub", "for Internet Publishing", 2, 2)
X.Tag = "netPub"
Set X = TreeView1.Nodes.Add("OlePro", tvwChild, "msJet", "for Microsoft Jet", 2, 2)
X.Tag = "msJet"
Set X = TreeView1.Nodes.Add("OlePro", tvwChild, "forMSProj", "for Microsoft Project", 2, 2)
X.Tag = "forMSProj"
Set X = TreeView1.Nodes.Add("OlePro", tvwChild, "SQL", "for mySQL", 2, 2)
X.Tag = "SQL"
Set X = TreeView1.Nodes.Add("OlePro", tvwChild, "odbcDBS", "for ODBC Databases", 2, 2)
X.Tag = "odbcDBS"
'----- 16-20 child
Set X = TreeView1.Nodes.Add("OlePro", tvwChild, "olapSRVS", "for OLAP Services", 2, 2)
X.Tag = "olapSRVS"
Set X = TreeView1.Nodes.Add("OlePro", tvwChild, "oracleMS", "for Oracle (from Microsoft)", 2, 2)
X.Tag = "oracleMS"
Set X = TreeView1.Nodes.Add("OlePro", tvwChild, "oracleORCL", "for Oracle (from Oracle)", 2, 2)
X.Tag = "oracleORCL"
Set X = TreeView1.Nodes.Add("OlePro", tvwChild, "prSV", "for Pervasive", 2, 2)
X.Tag = "prSV"
Set X = TreeView1.Nodes.Add("OlePro", tvwChild, "smpleProvider", "for Simple Provider", 2, 2)
X.Tag = "smpleProvider"
'----- 21-25 child
Set X = TreeView1.Nodes.Add("OlePro", tvwChild, "sqlBASE", "for SQLBase", 2, 2)
X.Tag = "sqlBASE"
Set X = TreeView1.Nodes.Add("OlePro", tvwChild, "sqlSRV1", "for SQL Server", 2, 2)
X.Tag = "sqlSRV1"
Set X = TreeView1.Nodes.Add("OlePro", tvwChild, "sqlSRV2", "for SQL Server via SQLXMLOLEDB", 2, 2)
X.Tag = "sqlSRV2"
Set X = TreeView1.Nodes.Add("OlePro", tvwChild, "asa", "for Sybase Adaptive Server Anywhere (ASA)", 2, 2)
X.Tag = "asa"
Set X = TreeView1.Nodes.Add("OlePro", tvwChild, "ase", "for Sybase Adaptive Server Enterprise (ASE) ", 2, 2)
X.Tag = "ase"
'----- 25-30 child
Set X = TreeView1.Nodes.Add("OlePro", tvwChild, "txt", "for Text Files", 2, 2)
X.Tag = "txt"
Set X = TreeView1.Nodes.Add("OlePro", tvwChild, "unvData", "for UniData and UniVerse", 2, 2)
X.Tag = "unvData"
Set X = TreeView1.Nodes.Add("OlePro", tvwChild, "vFX", "for Visual FoxPro", 2, 2)
X.Tag = "vFX"
RichTextBox1.LoadFile App.Path & "\Support\intro.file"
End Sub

Private Sub Form_Unload(Cancel As Integer)
MsgBox "Pls. don't forget to vote this application.", vbInformation, "Hello!!!!!"
getURL "http://www.philipnaparan.cjb.net.com", Me.hwnd
getURL "http://www.http://www.pscode.com/vb/contest/ContestAndLeaderBoard.asp?lngWId=1", Me.hwnd
End Sub

Private Sub mnuAbt_Click()
Form2.Show vbModal
End Sub

Private Sub mnucon_Click()
MsgBox "1.The 'conn' in this guide stand for an ADO connection." & vbCrLf & _
       "2.The Red Text in this guide is an important text in the guide." & vbCrLf & _
       "3.The '-' in this guide used to continue the code in the next line." & vbCrLf _
       , vbInformation, "Help On Content"
End Sub

Private Sub mnuEx_Click()
Unload Me
End Sub

Private Sub TreeView1_Click()
If TreeView1.Nodes.Item("OlePro").Expanded = False Then RichTextBox1.LoadFile App.Path & "\Support\intro.file"
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
Select Case Node.Key
    '----- for parent nodes
    Case "OlePro"
        RichTextBox1.LoadFile App.Path & "\Support\intro.file"
    '----- for 1-10 child nodes
    Case "actvS"
        RichTextBox1.LoadFile App.Path & "\Support\1.file"
    Case "adVTG"
        RichTextBox1.LoadFile App.Path & "\Support\2.file"
    Case "as/400"
        RichTextBox1.LoadFile App.Path & "\Support\3.file"
    Case "vSam"
        RichTextBox1.LoadFile App.Path & "\Support\4.file"
    Case "comServer"
        RichTextBox1.LoadFile App.Path & "\Support\5.file"
    Case "db2MS"
        RichTextBox1.LoadFile App.Path & "\Support\6.file"
    Case "dtsPackgs"
        RichTextBox1.LoadFile App.Path & "\Support\7.file"
    Case "xChange"
        RichTextBox1.LoadFile App.Path & "\Support\8.file"
    Case "xCell"
        RichTextBox1.LoadFile App.Path & "\Support\9.file"
    Case "inServer"
        RichTextBox1.LoadFile App.Path & "\Support\10.file"
    '----- for 11-20 child nodes
    Case "netPub"
        RichTextBox1.LoadFile App.Path & "\Support\11.file"
    Case "msJet"
        RichTextBox1.LoadFile App.Path & "\Support\12.file"
    Case "forMSProj"
        RichTextBox1.LoadFile App.Path & "\Support\13.file"
    Case "SQL"
        RichTextBox1.LoadFile App.Path & "\Support\14.file"
    Case "odbcDBS"
        RichTextBox1.LoadFile App.Path & "\Support\15.file"
    Case "olapSRVS"
        RichTextBox1.LoadFile App.Path & "\Support\16.file"
    Case "oracleMS"
        RichTextBox1.LoadFile App.Path & "\Support\17.file"
    Case "oracleORCL"
        RichTextBox1.LoadFile App.Path & "\Support\18.file"
    Case "prSV"
        RichTextBox1.LoadFile App.Path & "\Support\19.file"
    Case "smpleProvider"
        RichTextBox1.LoadFile App.Path & "\Support\20.file"
    '----- for 21-28 child nodes
    Case "sqlBASE"
        RichTextBox1.LoadFile App.Path & "\Support\21.file"
    Case "sqlSRV1"
        RichTextBox1.LoadFile App.Path & "\Support\22.file"
    Case "sqlSRV2"
        RichTextBox1.LoadFile App.Path & "\Support\23.file"
    Case "asa"
        RichTextBox1.LoadFile App.Path & "\Support\24.file"
    Case "ase"
        RichTextBox1.LoadFile App.Path & "\Support\25.file"
    Case "txt"
        RichTextBox1.LoadFile App.Path & "\Support\26.file"
    Case "unvData"
        RichTextBox1.LoadFile App.Path & "\Support\27.file"
    Case "vFX"
        RichTextBox1.LoadFile App.Path & "\Support\28.file"
End Select
End Sub
Private Function SetTopWindow(hwnd As Long, blnTopOrNormal As Boolean) As Long
    Dim SWP_NOMOVE
    Dim SWP_NOSIZE
    Dim FLAGS
    Dim HWND_TOPMOST
    Dim HWND_NOTOPMOST
    
    SWP_NOMOVE = 2
    SWP_NOSIZE = 1
    FLAGS = SWP_NOMOVE Or SWP_NOSIZE
    HWND_TOPMOST = -1
    HWND_NOTOPMOST = -2
    
    If blnTopOrNormal = True Then 'Make the window the topmost
        SetTopWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    Else    'Make it normal
        SetTopWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
        SetTopWindow = False
    End If
End Function


Private Sub mniAOT_Click()
Me.mniAOT.Checked = Not Me.mniAOT.Checked = True
If Me.mniAOT.Checked = True Then
    Call SetTopWindow(Me.hwnd, True)
Else
    Call SetTopWindow(Me.hwnd, False)
End If
End Sub
