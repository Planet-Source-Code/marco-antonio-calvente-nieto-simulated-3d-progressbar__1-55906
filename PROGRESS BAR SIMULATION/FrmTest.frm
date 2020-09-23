VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmTest 
   Caption         =   "Test of Progress bar"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9570
   Icon            =   "FrmTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7860
   ScaleWidth      =   9570
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   675
      TabIndex        =   0
      Top             =   5880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Default"
      Height          =   375
      Left            =   120
      TabIndex        =   54
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Numeric Values"
      Height          =   6615
      Left            =   4800
      TabIndex        =   19
      Top             =   120
      Width           =   4695
      Begin VB.Frame Frame6 
         Caption         =   "Other Numeric Values"
         Height          =   3495
         Left            =   120
         TabIndex        =   35
         Top             =   3000
         Width           =   4455
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   3480
            TabIndex        =   44
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   3480
            TabIndex        =   43
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   3480
            TabIndex        =   42
            Top             =   1320
            Width           =   855
         End
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   3480
            TabIndex        =   41
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   3480
            TabIndex        =   40
            Top             =   2040
            Width           =   855
         End
         Begin VB.TextBox Text9 
            Height          =   285
            Left            =   3480
            TabIndex        =   39
            Top             =   2400
            Width           =   855
         End
         Begin VB.TextBox Text10 
            Height          =   285
            Left            =   3480
            TabIndex        =   38
            Top             =   2760
            Width           =   855
         End
         Begin VB.TextBox Text11 
            Height          =   285
            Left            =   3480
            TabIndex        =   37
            Top             =   3120
            Width           =   855
         End
         Begin VB.TextBox Text3 
            Height          =   315
            Left            =   3480
            TabIndex        =   36
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label12 
            Caption         =   "Picturebox Width"
            Height          =   285
            Left            =   120
            TabIndex        =   53
            Top             =   600
            Width           =   3000
         End
         Begin VB.Label Label13 
            Caption         =   "Picture box Heigth"
            Height          =   285
            Left            =   120
            TabIndex        =   52
            Top             =   960
            Width           =   3000
         End
         Begin VB.Label Label14 
            Caption         =   "Maximum Heigth of the bar"
            Height          =   285
            Left            =   120
            TabIndex        =   51
            Top             =   1320
            Width           =   3000
         End
         Begin VB.Label Label15 
            Caption         =   "Maximum Widht of the bar"
            Height          =   285
            Left            =   120
            TabIndex        =   50
            Top             =   1680
            Width           =   3000
         End
         Begin VB.Label Label16 
            Caption         =   "How up the bar starts"
            Height          =   285
            Left            =   120
            TabIndex        =   49
            Top             =   2040
            Width           =   3000
         End
         Begin VB.Label Label17 
            Caption         =   "Where the Side of the bar starts"
            Height          =   285
            Left            =   120
            TabIndex        =   48
            Top             =   2400
            Width           =   3000
         End
         Begin VB.Label Label18 
            Caption         =   "Where the Top of the bar starts"
            Height          =   285
            Left            =   120
            TabIndex        =   47
            Top             =   2760
            Width           =   3000
         End
         Begin VB.Label Label19 
            Caption         =   "Extra heigth of the bar"
            Height          =   285
            Left            =   120
            TabIndex        =   46
            Top             =   3120
            Width           =   3000
         End
         Begin VB.Label Label9 
            Caption         =   "Introduce Longitude percent:"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Gradient color select"
         Height          =   2655
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   4455
         Begin VB.Frame Frame5 
            Caption         =   "Color 1"
            Height          =   2295
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   2055
            Begin MSComctlLib.Slider V1 
               Height          =   1695
               Left            =   120
               TabIndex        =   29
               Top             =   480
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   2990
               _Version        =   393216
               Orientation     =   1
               Max             =   255
               TickStyle       =   2
               TickFrequency   =   0
            End
            Begin MSComctlLib.Slider Slider1 
               Height          =   1695
               Left            =   720
               TabIndex        =   30
               Top             =   480
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   2990
               _Version        =   393216
               Orientation     =   1
               Max             =   255
               TickStyle       =   2
               TickFrequency   =   0
            End
            Begin MSComctlLib.Slider Slider2 
               Height          =   1695
               Left            =   1320
               TabIndex        =   31
               Top             =   480
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   2990
               _Version        =   393216
               Orientation     =   1
               Max             =   255
               TickStyle       =   2
               TickFrequency   =   0
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Red"
               Height          =   195
               Left            =   277
               TabIndex        =   34
               Top             =   240
               Width           =   300
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Green"
               Height          =   195
               Left            =   810
               TabIndex        =   33
               Top             =   240
               Width           =   435
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Blue"
               Height          =   195
               Left            =   1470
               TabIndex        =   32
               Top             =   240
               Width           =   315
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Color 2"
            Height          =   2295
            Left            =   2280
            TabIndex        =   21
            Top             =   240
            Width           =   2055
            Begin MSComctlLib.Slider Slider3 
               Height          =   1695
               Left            =   120
               TabIndex        =   22
               Top             =   480
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   2990
               _Version        =   393216
               Orientation     =   1
               Min             =   1
               Max             =   255
               SelStart        =   1
               TickStyle       =   2
               TickFrequency   =   0
               Value           =   1
            End
            Begin MSComctlLib.Slider Slider4 
               Height          =   1695
               Left            =   720
               TabIndex        =   23
               Top             =   480
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   2990
               _Version        =   393216
               Orientation     =   1
               Max             =   255
               TickStyle       =   2
               TickFrequency   =   0
            End
            Begin MSComctlLib.Slider Slider5 
               Height          =   1695
               Left            =   1320
               TabIndex        =   24
               Top             =   480
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   2990
               _Version        =   393216
               Orientation     =   1
               Max             =   255
               TickStyle       =   2
               TickFrequency   =   0
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Red"
               Height          =   195
               Left            =   270
               TabIndex        =   27
               Top             =   240
               Width           =   300
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Green"
               Height          =   195
               Left            =   810
               TabIndex        =   26
               Top             =   240
               Width           =   435
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Blue"
               Height          =   195
               Left            =   1470
               TabIndex        =   25
               Top             =   240
               Width           =   315
            End
         End
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Non Numeric values"
      Height          =   4455
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4575
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "FrmTest.frx":08CA
         Left            =   2880
         List            =   "FrmTest.frx":08E3
         TabIndex        =   18
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   2880
         TabIndex        =   16
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   2880
         TabIndex        =   14
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FrmTest.frx":0925
         Left            =   2880
         List            =   "FrmTest.frx":093E
         TabIndex        =   12
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Alternate gradient style"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   10
         Top             =   4080
         Width           =   2500
      End
      Begin VB.CheckBox Check1 
         Caption         =   "No picture box border"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   9
         Top             =   3720
         Width           =   2500
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Do not draw lines"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   3360
         Width           =   2500
      End
      Begin VB.CheckBox Check1 
         Caption         =   "2D Bar"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   3000
         Width           =   2500
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Move the bar to the bottom"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   2640
         Width           =   2500
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Make picturebox invisible"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   2280
         Width           =   2500
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Add a percentage counter"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   1920
         Width           =   2500
      End
      Begin VB.Label Label8 
         Caption         =   "Introduce text:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Introduce Value:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   390
         Width           =   2535
      End
      Begin VB.Label Label11 
         Caption         =   "Choose backcolor for picturebox:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label Label10 
         Caption         =   "Choose font and lines color:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   2535
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Stop!"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Go on!"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   960
      Top             =   5880
   End
End
Attribute VB_Name = "FrmtEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Command1_Click()
Call DefaultValues
End Sub

Private Sub Command2_Click()
Dim valu As Integer
Dim i As Integer
If Text1.Text = "" Then Exit Sub
valu = CInt(Text1.Text)
prgbar.SetUp Picture1, valu, Me, Text2.Text, Determine(0), Determine2(Combo1.Text), Determine(1), _
             Determine(2), Text3.Text, Determine(3), Determine(4), Determine(5), _
             Determine2(Combo2.Text), Text4.Text, Text5.Text, Text6.Text, Text7.Text, _
             Text8.Text, Text9.Text, Text10.Text, Text11.Text, V1.VALUE, _
             Slider1.VALUE, Slider2.VALUE, Slider3.VALUE, Slider4.VALUE, Slider5.VALUE, Determine(6)
Timer1.Enabled = True
End Sub

Private Sub Command4_Click()
prgbar.Reset
DefaultValues
End Sub

Private Sub Form_Load()
DefaultValues
End Sub

Private Sub DefaultValues()
Dim i As Integer
Combo1.Text = "VbBlack"
Combo2.Text = "VbWhite"
For i = 0 To 6
    Check1(i).VALUE = 0
Next i
Check1(0).VALUE = 1
V1.VALUE = 0
Slider1.VALUE = 0
Slider2.VALUE = 0
Slider3.VALUE = 100
Slider4.VALUE = 220
Slider5.VALUE = 255
Text3.Text = ""
Text3.Text = ""
Text3.Text = 99
Text4.Text = 6135
Text5.Text = 600
Text6.Text = 5500
Text7.Text = 1240
Text8.Text = 320
Text9.Text = 200
Text10.Text = 200
Text11.Text = 0

End Sub

Private Sub Timer1_Timer()
prgbar.grow (True)
End Sub
Private Function Determine(index As Integer)
If Check1(index).VALUE = 1 Then Determine = True Else Determine = False
End Function


Private Function Determine2(color As String) As ColorConstants
If color = "VbBlack" Then
    Determine2 = vbBlack
ElseIf color = "VbBlue" Then
    Determine2 = vbBlue
ElseIf color = "VbCyan" Then
    Determine2 = vbCyan
ElseIf color = "VbGreen" Then
    Determine2 = vbGreen
ElseIf color = "VbMagenta" Then
    Determine2 = vbMagenta
ElseIf color = "VbRed" Then
    Determine2 = vbRed
ElseIf color = "VbWhite" Then
    Determine2 = vbWhite
ElseIf color = "VbYellow" Then
    Determine2 = vbYellow
End If









End Function
