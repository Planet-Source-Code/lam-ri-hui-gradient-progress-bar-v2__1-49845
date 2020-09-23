VERSION 5.00
Begin VB.Form frmDemo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gradient Progress Bar v2 "
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10425
   Icon            =   "frmDemo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   10425
   StartUpPosition =   3  'Windows Default
   Begin Gradient_Progress_Bar_v2.GradientProgressBar gpb 
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   873
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4320
      TabIndex        =   0
      Top             =   5160
      Width           =   1815
   End
   Begin Gradient_Progress_Bar_v2.GradientProgressBar gpb 
      Height          =   735
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1296
      GradientType    =   1
   End
   Begin Gradient_Progress_Bar_v2.GradientProgressBar gpb 
      Height          =   735
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1296
      GradientType    =   2
   End
   Begin Gradient_Progress_Bar_v2.GradientProgressBar gpb 
      Height          =   735
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1296
      GradientType    =   3
   End
   Begin Gradient_Progress_Bar_v2.GradientProgressBar gpb 
      Height          =   735
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1296
      GradientType    =   4
   End
   Begin Gradient_Progress_Bar_v2.GradientProgressBar gpb 
      Height          =   735
      Index           =   5
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1296
      GradientType    =   5
   End
   Begin Gradient_Progress_Bar_v2.GradientProgressBar gpb 
      Height          =   735
      Index           =   6
      Left            =   5280
      TabIndex        =   7
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1296
      GradientType    =   6
   End
   Begin Gradient_Progress_Bar_v2.GradientProgressBar gpb 
      Height          =   735
      Index           =   7
      Left            =   5280
      TabIndex        =   8
      Top             =   960
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1296
      GradientType    =   7
   End
   Begin Gradient_Progress_Bar_v2.GradientProgressBar gpb 
      Height          =   735
      Index           =   8
      Left            =   5280
      TabIndex        =   9
      Top             =   1800
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1296
      GradientType    =   8
   End
   Begin Gradient_Progress_Bar_v2.GradientProgressBar gpb 
      Height          =   735
      Index           =   9
      Left            =   5280
      TabIndex        =   10
      Top             =   2640
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1296
      GradientType    =   9
   End
   Begin Gradient_Progress_Bar_v2.GradientProgressBar gpb 
      Height          =   735
      Index           =   10
      Left            =   5280
      TabIndex        =   11
      Top             =   3480
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1296
      GradientType    =   10
   End
   Begin Gradient_Progress_Bar_v2.GradientProgressBar gpb 
      Height          =   735
      Index           =   11
      Left            =   5280
      TabIndex        =   12
      Top             =   4320
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1296
      GradientType    =   11
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Copyright Lam Ri Hui 2003"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3840
      TabIndex        =   13
      Top             =   5760
      Width           =   2820
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Gradient Progress Bar v2 Demo Form
'By Lam Ri Hui

'Any questions or suggestion please email to
'rihui@email.com

'If you downloaded this progress bar from
'Planet Source Code, please vote for me.

Option Explicit
Private Sub Command1_Click()

Dim Work(1 To 5000) As String
Dim j                As Integer
Dim i                As Integer
Dim counter          As Integer

    For i = 0 To 11
        DoEvents
        gpb(i).Max = 100
        For j = 0 To 100
            gpb(i).Value = j
            For counter = LBound(Work) To UBound(Work)
                DoEvents
            Next counter
            gpb(i).Caption = j / 100 * 100 & "%"
            DoEvents
        Next j
    Next i

End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
