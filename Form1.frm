VERSION 5.00
Object = "{9D4122F1-4453-4250-AB9C-67E353839AE6}#8.0#0"; "ActSlider.ocx"
Begin VB.Form frmDemo 
   BackColor       =   &H00FF00FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Slider Demo"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Orientation"
      Height          =   735
      Left            =   120
      TabIndex        =   19
      Top             =   2520
      Width           =   3855
      Begin VB.OptionButton Option3 
         Caption         =   "Bottom To Top"
         Height          =   195
         Index           =   3
         Left            =   2040
         TabIndex        =   23
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Top To Bottom"
         Height          =   195
         Index           =   2
         Left            =   2040
         TabIndex        =   22
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Right To Left"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Left To Right"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1695
      End
   End
   Begin ActSlider.SliderPic SliderPic1 
      Height          =   300
      Left            =   960
      TabIndex        =   18
      Top             =   120
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   529
      ImageSlider     =   "Form1.frx":0000
      ImageLeft       =   "Form1.frx":0F52
      ImagePointer    =   "Form1.frx":1EA4
      BackStyle       =   0
      Orientation     =   3
   End
   Begin VB.Frame Frame2 
      Caption         =   "Enable"
      Height          =   1335
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   1935
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1080
         TabIndex        =   16
         Text            =   "Text4"
         Top             =   960
         Width           =   615
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Darker"
         Height          =   195
         Index           =   1
         Left            =   960
         TabIndex        =   14
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Lighter"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Enabled"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   1800
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   120
         X2              =   1800
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label4 
         Caption         =   "Intensity"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   615
      End
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Transparent"
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   480
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      Caption         =   "Position Set"
      Height          =   1575
      Left            =   2160
      TabIndex        =   2
      Top             =   840
      Width           =   1815
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   840
         TabIndex        =   9
         Text            =   "Text3"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   840
         TabIndex        =   8
         Text            =   "Text2"
         Top             =   540
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Set"
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   840
         TabIndex        =   3
         Text            =   "0"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Position"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Maximum"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Minimum"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   540
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Style"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   1935
      Begin VB.OptionButton Option1 
         Caption         =   "Discrete"
         Height          =   195
         Index           =   1
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   885
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Analog"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
Dim i As Long

If Check1.Value = vbChecked Then
    SliderPic1.ShowPosition = False
    For i = 0 To 3
        Option2(i).Enabled = False
    Next i
Else
    SliderPic1.ShowPosition = True
    For i = 0 To 3
        Option2(i).Enabled = True
    Next i
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = vbChecked Then
    SliderPic1.Enabled = True
Else
    SliderPic1.Enabled = False
End If

End Sub

Private Sub Check3_Click()
If Check3.Value = vbUnchecked Then
    SliderPic1.BackStyle = asOpaque
Else
    SliderPic1.BackStyle = asTransparent
End If

End Sub

Private Sub Command1_Click()
SliderPic1.Min = Text2
SliderPic1.Max = Text3

SliderPic1.CurPosition = Text1.Text
Text1.Text = SliderPic1.CurPosition

End Sub

Private Sub Form_Load()
Text2 = SliderPic1.Min
Text3 = SliderPic1.Max
If SliderPic1.BackStyle = asOpaque Then
    Check3.Value = vbUnchecked
Else
    Check3.Value = vbChecked
End If
Text4 = SliderPic1.DisabledIntense
Option2(SliderPic1.DisabledStyle).Value = True
If SliderPic1.Enabled Then
    Check2.Value = vbChecked
Else
    Check2.Value = vbUnchecked
End If
Option3(SliderPic1.Orientation).Value = True

End Sub

Private Sub Option1_Click(Index As Integer)
SliderPic1.Style = Index

End Sub

Private Sub Option2_Click(Index As Integer)
SliderPic1.DisabledStyle = Index

End Sub

Private Sub Option3_Click(Index As Integer)
SliderPic1.Orientation = Index

End Sub

Private Sub SliderPic1_PositionChanged(oldPosition As Long, newPosition As Long)
Text1 = newPosition
End Sub

Private Sub Text4_Change()
On Error GoTo erh

SliderPic1.DisabledIntense = CLng(Text4.Text)
erh:

End Sub
