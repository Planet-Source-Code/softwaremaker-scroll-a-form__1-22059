VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll1 
      Height          =   2115
      Left            =   4440
      TabIndex        =   2
      Top             =   60
      Width           =   255
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Text            =   "Scroll and You Can see me Clearly"
      Top             =   2580
      Width           =   3735
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   3060
      Width           =   4635
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   4080
      X2              =   7800
      Y1              =   2280
      Y2              =   2580
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Dim VPos As Integer
Dim Hpos As Integer

  'Change the following numbers to the Full height and width of your Form
  intFullHeight = 9000 'Maximized the Form and note the Figures
  intFullWidth = 12000
  'This is the how much of your Form is displayed
  intDisplayHeight = Me.Height
  intDisplayWidth = Me.Width

  With VScroll1
    '.Height = Me.ScaleHeight
    .Min = 0
    .Max = intFullHeight - intDisplayHeight
    .SmallChange = Screen.TwipsPerPixelX * 10
    .LargeChange = .SmallChange
  End With
    
  With HScroll1
    '.Width = Me.ScaleWidth
    .Min = 0
    .Max = intFullWidth - intDisplayWidth
    .SmallChange = Screen.TwipsPerPixelX * 10
    .LargeChange = .SmallChange
  End With
    
End Sub

Sub ScrollForm(Direction As Byte, NewVal As Integer)
  
  Dim CTL As Control
  Static hOldVal As Integer
  Static vOldVal As Integer
  Dim hMoveDiff As Integer 'Diff in the horizontal controls movements
  Dim vMoveDiff As Integer 'Diff in the vertical controls Movements
  
  Select Case Direction
    
  Case 0 'Scroll Vertically
  
    'Check The Direction of the Vertical Scroll & Extract Value Diff
    If NewVal > vOldVal Then 'Scrolled From Top to Bottom
      'Controls MUST move to the TOP, therefore TOP value Decreases
      vMoveDiff = -(NewVal - vOldVal)
    Else 'Scrolled From Bottom to Top
      'Controls MUST move to the Bottom, therefore TOP value Increases
      vMoveDiff = (vOldVal - NewVal)
    End If
  
    For Each CTL In Me.Controls
      'Make sure it's not a ScrollBar
      If Not (TypeOf CTL Is VScrollBar) And Not _
             (TypeOf CTL Is HScrollBar) Then
        'If it's a Line then
        If TypeOf CTL Is Line Then
          CTL.Y1 = CTL.Y1 + vMoveDiff '+ VPos - VScroll1.Value
          CTL.Y2 = CTL.Y2 + vMoveDiff '+ VPos - VScroll1.Value
        Else
          CTL.Top = CTL.Top + vMoveDiff '+ VPos - VScroll1.Value
        End If
      End If
    Next
    
      vOldVal = NewVal 'Reset vOldVal to reflect New Pos of ScrollBar
    
  Case 1 'Scroll Horizontally
  
    'Check The Direction of the Horizontal Scroll & Extract Value Diff
    If NewVal > hOldVal Then 'Scrolled From Left to Right
      'Controls MUST move to the LEFT, therefore LEFT value Decreases
      hMoveDiff = -(NewVal - hOldVal)
    Else 'Scrolled From Right to Left
      'Controls MUST move to the RIGHT, therefore LEFT value Increases
      hMoveDiff = (hOldVal - NewVal)
    End If
  
    For Each CTL In Me.Controls
      'Make sure it's not a ScrollBar
      If Not (TypeOf CTL Is VScrollBar) And Not _
             (TypeOf CTL Is HScrollBar) Then
        'If it's a Line then
        If TypeOf CTL Is Line Then
          CTL.X1 = CTL.X1 + hMoveDiff
          CTL.X2 = CTL.X2 + hMoveDiff
        Else
          CTL.Left = CTL.Left + hMoveDiff
        End If
      End If
    Next
      
      hOldVal = NewVal 'Reset hOldVal to reflect New Pos of ScrollBar
    
  End Select

End Sub

Private Sub HScroll1_Change()
  
  ScrollForm 1, HScroll1.Value

End Sub

Private Sub HScroll1_Scroll()
  
  ScrollForm 1, HScroll1.Value

End Sub

Private Sub VScroll1_Change()
  
  ScrollForm 0, VScroll1.Value

End Sub

Private Sub VScroll1_Scroll()
  
  ScrollForm 0, VScroll1.Value

End Sub
