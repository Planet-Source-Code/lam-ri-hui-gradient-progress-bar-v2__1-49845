VERSION 5.00
Begin VB.UserControl GradientProgressBar 
   AutoRedraw      =   -1  'True
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4710
   ScaleHeight     =   390
   ScaleWidth      =   4710
   ToolboxBitmap   =   "Gradient Progress Bar.ctx":0000
   Begin VB.PictureBox MainBox 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.PictureBox Progress 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   15
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   15
         Begin VB.Label Stat2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   2175
            TabIndex        =   3
            Top             =   60
            Width           =   465
         End
      End
      Begin VB.Label Stat1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Waiting to begin..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   60
         Width           =   1350
      End
   End
End
Attribute VB_Name = "GradientProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************
'*****      Gradient Progress Bar v2      *****
'**********************************************
'Release Date : 12th Nov 2003
'Author : Lam Ri Hui

'This is the second version of gradient progress
'bar in its series. This progress bar now comes
'with 12 types of gradient in two catogories :
'1. Horizontal Gradient
'   a)Red
'   b)Green
'   c)Blue
'   d)Grey
'   e)Purple
'   f)Yellow
'2. Vertical Gradient
'   a)Red
'   b)Green
'   c)Blue
'   d)Grey
'   e)Purple
'   f)Yellow

'If you liked Gradient Progress Bar v2, please
'email me at rihui@email.com or vote for me so that
'I know how many of you out there liked this progress
'bar.

Option Explicit
Private ProgVal As Integer
Private MaxNum As Long
    Const m_def_GradientType = 0

Private m_GradientType As Integer
Public Property Let Max(lngNum As Long)

    MaxNum = lngNum

End Property
Public Property Get Max() As Long

    Max = MaxNum

End Property
Public Property Let Value(IntValue As Long)

    On Error Resume Next
    If IntValue = 0 Then
        Progress.Visible = False
        Else
        Progress.Visible = True
    End If
    ProgVal = IntValue
    Progress.Width = MainBox.Width * (ProgVal / MaxNum)
    On Error Resume Next
Dim intLoop As Integer
Dim X As Integer
'change the progress's drawstyle to vbDashDot and see what's different.
    Progress.DrawStyle = vbSolid
    Progress.DrawMode = vbCopyPen
    Progress.ScaleMode = vbPixels
    Progress.DrawWidth = 2
    Progress.ScaleHeight = 256
    Select Case GradientType
        Case 0  'Red Vertical Gradient
        For intLoop = 0 To 255
            Progress.Line (X, 0)-(X - 100, Progress.Height + 100), RGB(intLoop, 0, 0), BF
            X = X + 1.5
        Next
        Case 1  'Green Vertical Gradient
        For intLoop = 0 To 255
            Progress.Line (X, 0)-(X - 100, Progress.Height + 100), RGB(0, intLoop, 0), BF
            X = X + 1.5
        Next
        Case 2  'Blue Vertical Gradient
        For intLoop = 0 To 255
            Progress.Line (X, 0)-(X - 100, Progress.Height + 100), RGB(0, 0, intLoop), BF
            X = X + 1.5
        Next
        Case 3  'Grey Vertical Gradient
        For intLoop = 0 To 255
            Progress.Line (X, 0)-(X - 100, Progress.Height + 100), RGB(intLoop, intLoop, intLoop), BF
            X = X + 1.5
        Next
        Case 4  'Purple Vertical Gradient
        For intLoop = 0 To 255
            Progress.Line (X, 0)-(X - 100, Progress.Height + 100), RGB(intLoop, 0, intLoop), BF
            X = X + 1.5
        Next
        Case 5  'Yellow Vertical Gradient
        For intLoop = 0 To 255
            Progress.Line (X, 0)-(X - 100, Progress.Height + 100), RGB(intLoop, intLoop, 0), BF
            X = X + 1.5
        Next
        Case 6  'Red Horizontal Gradient
        For intLoop = 0 To 255
            Progress.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 0, 0), B
        Next
        Case 7  'Green Horizontal Gradient
        For intLoop = 0 To 255
            Progress.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 255 - intLoop, 0), B
        Next
        Case 8  'Blue Horizontal Gradient
        For intLoop = 0 To 255
            Progress.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 0, 255 - intLoop), B
        Next
        Case 9  'Grey Horizontal Gradient
        For intLoop = 0 To 255
            Progress.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 255 - intLoop, 255 - intLoop), B
        Next
        Case 10 'Purple Horizontal Gradient
        For intLoop = 0 To 255
            Progress.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 0, 255 - intLoop), B
        Next
        Case 11 'Yellow Horizontal Gradient
        For intLoop = 0 To 255
            Progress.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 255 - intLoop, 0), B
        Next
    End Select
    Refresh

End Property
Public Property Get Value() As Long

    ProgVal = Value

End Property
Public Property Let Caption(MyCaption As String)

    Stat1 = MyCaption
    Stat2 = MyCaption

End Property
Public Property Get Caption() As String

    Caption = Stat1

End Property

Private Sub MainBox_Click()

End Sub

Private Sub UserControl_Initialize()

    Progress.Visible = False
    UserControl_Resize

End Sub
Private Sub UserControl_Resize()

    MainBox.Width = UserControl.Width
    MainBox.Height = UserControl.Height
    Stat1.Left = 50
    Stat1.Top = (MainBox.Height / 2) - (Stat1.Height / 2) - 30
    Stat2.Left = 50
    Stat2.Top = Stat1.Top
    Progress.Height = MainBox.Height

End Sub
Public Property Get GradientType() As Integer

    GradientType = m_GradientType

End Property
Public Property Let GradientType(ByVal New_GradientType As Integer)

    Select Case New_GradientType
        Case 0 To 11
        m_GradientType = New_GradientType
        PropertyChanged "GradientType"
        Case Else
        MsgBox "Error setting gradient color." & vbNewLine & "The available colors are : " & vbNewLine & vbTab & "0 - Red Horizontal Gradient" & vbNewLine & vbTab & "1 - Green Horizontal Gradient" & vbNewLine & vbTab & "2 - Blue Horizontal Gradient" & vbNewLine & vbTab & "3 - Grey Horizontal Gradient" & vbNewLine & vbTab & "4 - Purple Horizontal Gradient" & vbNewLine & vbTab & "5 - Yellow Horizontal Gradient" & vbNewLine & vbTab & "6 - Red Vertical Gradient" & vbNewLine & vbTab & "7 - Green Vertical Gradient" & vbNewLine & vbTab & "8 - Blue Vertical Gradient" & vbNewLine & vbTab & "9 - Grey Vertical Gradient" & vbNewLine & vbTab & "10 - Purple Vertical Gradient" & vbNewLine & vbTab & "11 - Yellow Vertical Gradient", vbCritical, "Error"
    End Select

End Property
Private Sub UserControl_InitProperties()

    m_GradientType = m_def_GradientType

End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_GradientType = PropBag.ReadProperty("GradientType", m_def_GradientType)

End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("GradientType", m_GradientType, m_def_GradientType)

End Sub
