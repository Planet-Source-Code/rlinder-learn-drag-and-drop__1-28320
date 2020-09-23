VERSION 5.00
Begin VB.Form frmTraffic 
   Caption         =   "Online Soft Web LLC Traffic Sign"
   ClientHeight    =   3825
   ClientLeft      =   1260
   ClientTop       =   1545
   ClientWidth     =   6915
   Icon            =   "frmTraffic.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3825
   ScaleWidth      =   6915
   Begin VB.Image imgContainer 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   1
      Left            =   5520
      Top             =   2280
      Width           =   735
   End
   Begin VB.Image imgContainer 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   0
      Left            =   4200
      Top             =   2280
      Width           =   735
   End
   Begin VB.Image imgContainer 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   3
      Left            =   2880
      Top             =   2280
      Width           =   735
   End
   Begin VB.Image imgContainer 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   4
      Left            =   1560
      Top             =   2280
      Width           =   735
   End
   Begin VB.Image imgContainer 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   2
      Left            =   360
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Do Not Enter"
      Height          =   255
      Left            =   5520
      TabIndex        =   5
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Slippery Road"
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Speed Limit"
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Stop"
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Divided Highway"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Drag and Drop the Signs Into the Correct Boxes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
   Begin VB.Image imgSign 
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   4
      Left            =   5760
      Picture         =   "frmTraffic.frx":0442
      Top             =   600
      Width           =   480
   End
   Begin VB.Image imgSign 
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   3
      Left            =   4320
      Picture         =   "frmTraffic.frx":074C
      Top             =   600
      Width           =   480
   End
   Begin VB.Image imgSign 
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   2
      Left            =   3000
      Picture         =   "frmTraffic.frx":0A56
      Top             =   600
      Width           =   480
   End
   Begin VB.Image imgSign 
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   1
      Left            =   1680
      Picture         =   "frmTraffic.frx":0D60
      Top             =   600
      Width           =   480
   End
   Begin VB.Image imgSign 
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   0
      Left            =   600
      Picture         =   "frmTraffic.frx":106A
      Top             =   600
      Width           =   480
   End
End
Attribute VB_Name = "frmTraffic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Visible = True
    
End Sub

Private Sub Form_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Source.Visible = False
    
End Sub

Private Sub Form_Load()
    Dim intX As Integer
    
    frmTraffic.Top = (Screen.Height - frmTraffic.Height) / 2
    frmTraffic.Left = (Screen.Width - frmTraffic.Width) / 2
    
    For intX = 0 To 4
        imgSign(intX).DragIcon = imgSign(intX).Picture
    Next intX
    
End Sub

Private Sub imgContainer_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Dim intRight As Integer
    
    If Source.Index = Index Then
        imgContainer(Index).Picture = Source.Picture
        intRight = intRight + 1
        
        If intRight = 5 Then
            MsgBox "Well Done", vbExclamation, "Drag Drop"
        End If
    Else
        Source.Visible = True
        
    End If
    
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Visible = True
    
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Visible = True
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Visible = True
    
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Visible = True
    
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Visible = True
    
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Visible = True
    
End Sub
