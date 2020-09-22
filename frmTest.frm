VERSION 5.00
Begin VB.Form frmTest 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Append to the System Menu"
   ClientHeight    =   3390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6270
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTest.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   6270
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4560
      TabIndex        =   8
      Top             =   2760
      Width           =   1260
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Check the System Menu out!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   2760
      Width           =   4215
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Append to the System Menu right from your app!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Index           =   2
      Left            =   2400
      TabIndex        =   5
      Top             =   1140
      Width           =   2835
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Aerodynamica"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   480
      Index           =   1
      Left            =   1440
      TabIndex        =   4
      Top             =   180
      Width           =   2865
   End
   Begin VB.Label lblDotLine 
      BackStyle       =   0  'Transparent
      Caption         =   "................."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   780
      Width           =   2115
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "AppendToSys"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Index           =   1
      Left            =   1680
      TabIndex        =   2
      Top             =   540
      Width           =   3555
   End
   Begin VB.Label lblDotLine 
      BackStyle       =   0  'Transparent
      Caption         =   "................"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3660
      TabIndex        =   1
      Top             =   120
      Width           =   2475
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmTest.frx":000C
      Height          =   615
      Index           =   0
      Left            =   1380
      TabIndex        =   0
      Top             =   1800
      Width           =   4215
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00000000&
      Height          =   870
      Index           =   3
      Left            =   1380
      TabIndex        =   7
      Top             =   840
      Width           =   3885
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H0080C0FF&
      Height          =   225
      Index           =   4
      Left            =   1380
      TabIndex        =   6
      Top             =   600
      Width           =   3885
   End
   Begin VB.Shape shpRect 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   16
      Height          =   855
      Index           =   1
      Left            =   480
      Top             =   300
      Width           =   855
   End
   Begin VB.Shape shpRect 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   2
      Left            =   5160
      Top             =   240
      Width           =   555
   End
   Begin VB.Shape shpRect 
      BorderColor     =   &H000080FF&
      BorderWidth     =   8
      Height          =   675
      Index           =   0
      Left            =   4800
      Top             =   420
      Width           =   615
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'do not press the stop button on the VB IDE toolbar
'or this example will crash!
Dim MyHeight, MyWidth As Integer

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
    Dim lhSysMenu As Long
    Dim lRet As Long
    
    'for resize stuff (needed for this example only)
    MyHeight = Height
    MyWidth = Width
    
    'add the menu items
    lhSysMenu = GetSystemMenu(hWnd, 0&)
    lRet = AppendMenu(lhSysMenu, MF_SEPARATOR, 0&, vbNullString)
    lRet = AppendMenu(lhSysMenu, MF_STRING, IDM_ITEM1, "Item 1")
    lRet = AppendMenu(lhSysMenu, MF_STRING, IDM_ITEM2, "Item 2")
    lRet = AppendMenu(lhSysMenu, MF_SEPARATOR, 0&, vbNullString)
    lRet = AppendMenu(lhSysMenu, MF_STRING, IDM_ABOUT, "About...")
    
    Show
    
    'save previous windows message handler
    ProcOld = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Private Sub Form_Resize()
    'If minimized, exit sub
    If WindowState = 1 Then Exit Sub
    
    If Height > MyHeight Or Height < MyHeight Then
        Height = MyHeight
    End If
    
    If Width > MyWidth Or Width < MyWidth Then
        Width = MyWidth
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'this is necessary! if you don't add this,
    'VB will crash
    Call SetWindowLong(hWnd, GWL_WNDPROC, ProcOld)
End Sub
