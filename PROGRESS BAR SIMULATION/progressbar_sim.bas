Attribute VB_Name = "prgbar"
'################################################################################
'##                           SIMULATED 3D PROGRESSBAR                         ##
'##                                                                            ##
'##   This is an attempt to create a 3D PROGRESSBAR to allow the user to see   ##
'##   the progress of a long process in the application. This uses no OCX,only ##
'##   a Bas module and a standard, invisible picturebox in the form            ##
'##   This bar is used mainly to show the process of the aplication while      ##
'##   executing a repetition structure, like For..next or Do..loop, where you  ##
'##   know the number of times the structure will be executed. This bar is not ##
'##   much useful if you do not know that number, which will be needed for     ##
'##   Value parameter. See below for details                                   ##
'################################################################################

'################################################################################
'  What you need to know to implement this bar in your form
'  1.- Place this code in a bas module
'  2.- Place a picturebox in the form where you wish the progres bar to appear.
'      Is better to set Visible to false and borderstyle to 0 for that picturebox
'  3.- From any place within the form code, call the method SetUp to set up
'      the progress bar.When you call SetUp, you must specify a parameter
'      called VALUE, wich is used to calculate the speed of the bar. See the
'      SetUp method  to see the exact formula
'        Some examples of wich would be the value of VALUE
'       ·If you use For i = X to Y...Next i, VALUE = Y
'       ·If you use do...Loop, to navigate trougth a recodset, VALUE =
'        the recordcount propierty of that recordset
'  4.- From within a repetition structure, like For..next or Do..loop, call
'      the method grow  to make the bar grow. For better results and clarity
'      place the call just after the opening line or just before the end line
'      of the structure.
'      Optional: Insert a timer control and call the grow method from the timer
'      Code. Set that timer interval propierty to 1 for optimum speed
'  5.- In any point you wish the bar to dissappear, call the Reset method to
'      end all the process. Or set autostop parameter within method grow
'      to true for auto stop the bar when it concludes.
'################################################################################

Option Explicit

'Declairations
Dim GraphWidth As Integer
Dim GraphTopHeight As Integer
Dim GraphTop As Integer
Dim GraphStart As Integer 'Back lining start (X)
Dim Graph3DTop As Integer  'Top back lining start (Y)
Dim GraphMax As Integer 'Bar maximun lenght
Dim GraphMaxHeight As Integer 'Bar maximun Height
Dim GraphTopMinimum As Integer 'Bar minimun Height

'Not customisable
Const Graph3DAddWidth As Integer = 30
Const Graph3DReduceWidth As Integer = 10

'Declaration of color-related variables
Type Prg_Values
    vh As Integer
    rv As Integer
    gv As Integer
    bv As Integer
    rv1 As Integer
    gv1 As Integer
    bv1 As Integer
End Type
Dim Valuelist As Prg_Values

'Picture graph information
Dim Picnom As PictureBox
Dim GainAmmount As Double
Dim Gainammount2 As Double
Dim Gainmultiplier As Double
Dim linesout As Boolean
Dim Out3d As Boolean
Dim invertedgradient As Boolean

'Variables for prevent non-function due to higher values
Dim counter As Integer, revcount As Integer

'Variables for text and percentage
Dim Percent As Boolean, Bartext As String

'variables to store the original position of the picturebox
Dim pictop As Long
Dim picleft As Long


'Draw lines for 3D Bar
Private Function DrawGraphLines(PicObject As Object, GraphTop As Integer, GraphLineWidth As Integer, GraphTopHeight As Integer)
    'IF out3D is false, only lines of the 2D part of the bar will be drawed
    GraphWidth = GraphStart + GraphLineWidth
    
    If Out3d = False Then PicObject.Line (GraphWidth - GraphStart, GraphTop)-(GraphWidth, GraphTop - Graph3DTop)
    If Out3d = False Then PicObject.Line (GraphWidth - GraphStart, GraphTop + GraphTopHeight + Graph3DTop)-(GraphWidth, GraphTop + GraphTopHeight)
    PicObject.Line (GraphWidth - GraphStart, GraphTop)-(GraphWidth - GraphStart, GraphTop + GraphTopHeight + Graph3DTop)
    If Out3d = False Then PicObject.Line (GraphWidth - Graph3DReduceWidth, GraphTop - Graph3DTop)-(GraphWidth - Graph3DReduceWidth, GraphTop + GraphTopHeight + Graph3DAddWidth)
    
    If Out3d = False Then PicObject.Line (1, GraphTop)-(GraphStart, GraphTop - Graph3DTop)
    'PicObject.Line (1, GraphTop + GraphTopHeight + Graph3DTop)-(GraphStart, GraphTop + GraphTopHeight)
    'PicObject.Line (Graph3DAddWidth - Graph3DReduceWidth, GraphTop)-(Graph3DAddWidth - Graph3DReduceWidth, GraphTop + GraphTopHeight + Graph3DTop)
    'PicObject.Line (Graph3DAddWidth - Graph3DReduceWidth + GraphStart, GraphTop - Graph3DTop)-(Graph3DAddWidth - Graph3DReduceWidth + GraphStart, GraphTop + GraphTopHeight)
        
    PicObject.Line (1, GraphTop)-(GraphWidth - GraphStart, GraphTop)
    PicObject.Line (1, GraphTop + GraphTopHeight + Graph3DTop)-(GraphWidth - GraphStart, GraphTop + GraphTopHeight + Graph3DTop)
    'PicObject.Line (GraphStart, GraphTop + GraphTopHeight)-(GraphWidth, GraphTop + GraphTopHeight)
    If Out3d = False Then PicObject.Line (GraphStart, GraphTop - Graph3DTop)-(GraphWidth, GraphTop - Graph3DTop)
    
End Function

'Sets up functions that needs to be set
Private Function DrawGraph(PicObject As Object, GraphTop3D As Integer, GraphWidth3D As Integer, GraphHeight3D As Integer)
    'IF out3D is false, Top and end of the bar will not be drawed
    'Draw top of box
    If Out3d = False Then
        If DrawGradient(PicObject, 0, GraphTop3D, GraphWidth3D, GraphHeight3D, 0, 763, Valuelist, invertedgradient) = False Then
            'Do nothing
        End If
    End If
    'Draw side of box
        If DrawGradient(PicObject, 1, GraphTop3D, GraphWidth3D, GraphHeight3D, 0, 763, Valuelist) = False Then
            'Do nothing
        End If
    'Draw end of box
    If Out3d = False Then
        If DrawGradient(PicObject, 2, GraphTop3D, GraphWidth3D, GraphHeight3D, 0, 763, Valuelist) = False Then
            'Do nothing
        End If
    End If
    
    'Draw graph outer lines
    If linesout = False Then DrawGraphLines PicObject, GraphTop3D, GraphWidth3D, GraphHeight3D
        
End Function


'Draws the gradient, which is then set to the object (PicObject)
Private Function DrawGradient(PicObject As Object, SideDraw As Integer, WriteTop As Integer, WriteWidth As Integer, WriteHeight As Integer, St%, H%, Valuelist As Prg_Values, Optional inverted As Boolean = False) As Boolean
    
    On Error GoTo FinaliseError
    
    Dim H2%, H3%, IvR%, IvG%, IvB%
    Dim VR, VG, VB As Single
    Dim Color1, Color2 As Long
    Dim R, G, b, r2, g2, b2 As Integer
    Dim temp As Long
    If inverted = False Then
        Color1 = RGB(Valuelist.rv, Valuelist.gv, Valuelist.bv)
        Color2 = RGB(Valuelist.rv1, Valuelist.gv1, Valuelist.bv1)
    Else
        Color1 = RGB(Valuelist.rv1, Valuelist.gv1, Valuelist.bv1)
        Color2 = RGB(Valuelist.rv, Valuelist.gv, Valuelist.bv)
    End If
    temp = (Color1 And 255): R = temp And 255
    temp = Int(Color1 / 256): G = temp And 255
    temp = Int(Color1 / 65536): b = temp And 255
    temp = (Color2 And 255): r2 = temp And 255
    temp = Int(Color2 / 256): g2 = temp And 255
    temp = Int(Color2 / 65536): b2 = temp And 255
        
    VR = Abs(R - r2) / WriteWidth
    VG = Abs(G - g2) / WriteWidth
    VB = Abs(b - b2) / WriteWidth
    
    If r2 < R Then VR = -VR
    If g2 < G Then VG = -VG
    If b2 < b Then VB = -VB
    
    Dim GraphOccurance As Integer
    GraphWidth = GraphStart + WriteWidth
    GraphTopHeight = WriteHeight
    GraphTop = WriteTop
    
    GraphOccurance = GraphWidth
    
    Dim NextPosition As Integer
    NextPosition = 0
    
    If SideDraw = 2 Then GoTo DrawGraphEnd
    GraphWidth = GraphStart
    
    Do
        GraphWidth = GraphWidth + 15
        NextPosition = NextPosition + 15
        r2 = R + VR * NextPosition: g2 = G + VG * NextPosition: b2 = b + VB * NextPosition
        If SideDraw = 0 Then
            PicObject.Line (GraphWidth - GraphStart, GraphTop)-(GraphWidth, GraphTop - Graph3DTop), RGB(r2, g2, b2)
        ElseIf SideDraw = 1 Then
            PicObject.Line (GraphWidth - GraphStart, GraphTop)-(GraphWidth - GraphStart, GraphTop + GraphTopHeight + Graph3DTop), RGB(r2, g2, b2)
        End If
    Loop Until NextPosition >= GraphOccurance - GraphStart

FinaliseError:
    DrawGradient = True
    Exit Function
    
DrawGraphEnd:

    Do
        NextPosition = NextPosition + 15
        PicObject.Line (GraphWidth - GraphStart, GraphTop + NextPosition)-(GraphWidth, GraphTop - Graph3DTop + NextPosition), RGB(r2, g2, b2)
    Loop Until NextPosition >= GraphTopHeight + Graph3DTop
    DrawGradient = True
    
End Function


'######################################################################################
'                            SET UP AND CUSTOMIZING DETAILS
'
'To set up the progress bar, you need only the picturebox name,the form name and
'Value, which is used to determine the bar speed progress (see grow)
'All the others The optional values is for customization:
'   -Texto is for introducing an optional text above the bar, and percent_pre adds
'    as percentaje counter. Tcolor change text color... but also bar lines color
'    (I had been unable to fix this). Adding text and/or percent increases the
'    heigth of the picturebox and displaces the bar down
'   -Nobox make the bar remove the box in wich the bar appear if set to true
'   -Bottombar, if true, make the bar appear at the bottom of the box, occuping all
'    the longitude of the form. Good to simulate progressbar in a statuebar. Longitude
'    allows you to fix the bar maximum length in a percentaje value of the form
'    width(this is, with values from 1 to 100). A warning. Enabling this option
'    overrides the values of Pwidth and GM
'   -No3D generates a plain 2D bar instead of a 3D bar
'   -NoLines generates a bar without lines, only color
'   -Border, Pcolor, Pwidth and Pheigth controls the picturebox Borderstyle,
'    Backcolor, width and height properties
'   -BHeigth,GM,GMH,and GMT controls the extra heigth, maximum length, maximum heigth,
'    And minimum heigth of simulated progress bar.
'   -GS and G3D controls the starting end of the bar. See this ascii diagram:
'   '-->
'   _________
'   _________|< 3D end controled by GS
'
'   -->
'   ¯¯¯¯¯¯¯¯\ < 3D top controled by G3D
'   ¯¯¯¯¯¯¯¯¯|
'   _________|
'
'   -The Other bunch of values is for controling the two colors forming the
'    Gradient of the bar. Think in this way:
'       To Get First color = RGB(rv,gv,bv)(Black by default)
'       To Get Second color = RGB(rv1,gv1,bv1) (Turquoise by default)
'   -If you want no gradient, make first and second color match
'   -If you wish the gradient of the top of the bar go in the other direction,
'    set inverted to true

' This is all. Play with the values as you wish to obtain full utility from
' this code
'################################################################################

Public Sub SetUp(PicGraph As PictureBox, VALUE As Integer, frmowner As Form, _
                  Optional texto As String = "", Optional percent_pre As Boolean = True, Optional TColor As ColorConstants = vbBlack, _
                  Optional Nobox As Boolean = False, Optional bottombar As Boolean, Optional Longitude As Integer = 99, _
                  Optional No3D As Boolean = False, Optional Nolines As Boolean = False, _
                  Optional border As Boolean = True, Optional PColor As ColorConstants = vbWhite, _
                  Optional Pwidth As Integer = 6135, Optional Pheight As Integer = 600, _
                  Optional GM As Integer = 5500, Optional GMH As Integer = 1240, Optional GTM As Integer = 320, _
                  Optional GS As Integer = 200, Optional G3D As Integer = 200, _
                  Optional Bheight As Integer = 0, Optional rv As Integer = 0, _
                  Optional gv As Integer = 0, Optional bv As Integer = 0, _
                  Optional rv1 As Integer = 100, Optional gv1 As Integer = 220, _
                  Optional bv1 As Integer = 225, Optional inverse As Boolean = True)

'Fix valuelist values
Valuelist.vh = Bheight
Valuelist.rv = rv
Valuelist.gv = gv
Valuelist.bv = bv
Valuelist.rv1 = rv1
Valuelist.gv1 = gv1
Valuelist.bv1 = bv1

'Store the picturebox name
Set Picnom = PicGraph

'store the original position of the picture box in variables
pictop = PicGraph.Top
picleft = PicGraph.Left

'give proper size to the picturebox
If Pwidth < (frmowner.Width * 75 / 100) Then
    PicGraph.Width = Pwidth
    If texto <> "" Or percent_pre = True Then PicGraph.Height = Pheight + 200 Else PicGraph.Height = Pheight
Else
    PicGraph.Width = (frmowner.Width * 75 / 100)
    PicGraph.Height = Pheight
End If

' Fix bar values
If GM >= (frmowner.Width * 75 / 100) Then
    GraphMax = CInt((PicGraph.Width * 90 / 100))
Else
    GraphMax = GM
End If
GraphMaxHeight = GMH
If texto <> "" Or percent_pre = True Then GraphTopMinimum = GTM + 200 Else GraphTopMinimum = GTM
GraphStart = GS
Graph3DTop = G3D
If No3D = True Then
    Out3d = True
    GraphTopMinimum = GraphTopMinimum - 230
    PicGraph.Height = PicGraph.Height - 230
Else
    Out3d = False
End If
If Nolines = True Then linesout = True Else linesout = False
If inverse = True Then invertedgradient = True

'center the picturebox in the form
PicGraph.Top = (frmowner.Height / 2) - (PicGraph.Height / 2)
PicGraph.Left = (frmowner.Width / 2) - (PicGraph.Width / 2)

'set up a propierties for the picturebox
If border = True Then PicGraph.BorderStyle = 1
PicGraph.Appearance = 0
PicGraph.BackColor = PColor
If Nobox = True Then
    PicGraph.BackColor = frmowner.BackColor
    PicGraph.BorderStyle = 0
End If
If bottombar = True Then
    PicGraph.Align = 2
    GraphMax = CInt((frmowner.Width * Longitude / 100))
End If

'set variables for the grow method
If VALUE > GraphMax Then
    counter = CDbl(VALUE / GraphMax)
End If

'set text variables
Percent = percent_pre
Bartext = texto
Picnom.ForeColor = TColor


'This is for calculate the amount the bar grows each time grow is called
If counter <= 1 Then
    GainAmmount = CDbl(GraphMax / VALUE)
Else
    GainAmmount = 1
End If
'This too
Gainmultiplier = 1

'Finally, set autodraw in True and make the picturebox visible
PicGraph.AutoRedraw = True
PicGraph.Visible = True
frmowner.Refresh
End Sub

'this process reset the picturebox to its original state and position.
'This method is used automatically if you set autostops to True when you
'call the grow method, but I have leaved the the reset method public just in
'case you need to abort the process
Public Sub Reset()

On Error Resume Next
'remove picturebox from sight, so the next changes are not visible
Picnom.Visible = False

'return propierties to original values
Picnom.AutoRedraw = False
Picnom.BorderStyle = 0
Picnom.Align = 0

'return the picturebox to a discret size
Picnom.Width = 10
Picnom.Height = 10

'reseting variables
GainAmmount = 0
Gainmultiplier = 0

'clearing picture
Picnom.Cls

'return the picture box to its the original position using the values
'stored during SetUp
Picnom.Top = pictop
Picnom.Left = picleft
End Sub

'this process make the bar grows. Place it into a Do...Loop or For..next
'structure , or inside a timer control, to make the bar progress. The autostop
'pareameter, if set to true, make the bar to autoreset itself after completion.

Public Sub grow(Optional autostop As Boolean = False)
If counter <= 1 Then
    Truegrov
    If autostop = True Then
        If Gainammount2 >= GraphMax Then Reset
    End If
Else
    If revcount = 0 Then
        revcount = counter
        Truegrov
        If autostop = True Then
            If Gainammount2 >= GraphMax Then Reset
        End If
    Else
        revcount = revcount - 1
    End If
End If
End Sub



Private Function Actualicecount() As String
Dim pernum As Double
If Percent = True Then
    If counter <= 1 Then
        Actualicecount = CStr(Auxiliar.Redondear(((GainAmmount * Gainmultiplier) * 100) / GraphMax, 1)) & "%"
    Else
        Actualicecount = CStr(Auxiliar.Redondear((Gainmultiplier * 100) / GraphMax, 1)) & "%"
    End If
End If
End Function

Private Sub Truegrov()
'this variable store the current size of the bar
'clear the picturebox to redraw the bar
Picnom.Cls
'calculate the current size of the bar
If counter <= 1 Then
    Gainammount2 = GainAmmount * Gainmultiplier
Else
    Gainammount2 = Gainmultiplier
End If
'Add text and percent( if activated) and then call DrawGraph to draw the bar
Picnom.Print Bartext & "  " & Actualicecount
DrawGraph Picnom, GraphTopMinimum, Int(Gainammount2), Valuelist.vh 'Make graph
'increase Gainmultiplier in order to keep the bar growing
Gainmultiplier = Gainmultiplier + 1
'along with the setting of Picnom.AutoRedraw to true, this line
'prevent the bar from flickering while being drawed
Picnom.Refresh
End Sub

