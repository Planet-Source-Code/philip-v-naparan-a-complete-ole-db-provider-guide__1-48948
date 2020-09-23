VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About The Author"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "*Please Don't Forget To Vote This Application*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   6720
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "SPECIAL THANKS TO: EVANGELOS PETROUSOS AND TO EVERYONE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   6360
      Width           =   5415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "( WEB DEVELOPER/PROGRAMMER )"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   3000
      TabIndex        =   1
      Top             =   6120
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SOFTWARE BY: PHILIP V. NAPARAN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   6120
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   6000
      Left            =   0
      Picture         =   "Form2.frx":06EA
      ToolTipText     =   "Click Me To ExitThis Form"
      Top             =   0
      Width           =   7500
   End
End
Attribute VB_Name = "Form2"
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

Private Sub Form_Load()
Call SetTopWindow(Me.hwnd, True)
End Sub

Private Sub Image1_Click()
Unload Me
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
