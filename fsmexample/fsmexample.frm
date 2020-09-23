VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "FSM"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   6450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Try 'the cat chased dog a'"
      Height          =   975
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Try 'the cat chased a dog'"
      Height          =   975
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Determiner(1) As String
Private Noun(9) As String
Private Verb(4) As String
Private Sub Command1_Click()
Dim Inputs(4) As Variant

'Setup the input array
Inputs(0) = Determiner(1)
Inputs(1) = Noun(0)
Inputs(2) = Verb(4)
Inputs(4) = Determiner(1)
Inputs(3) = Noun(1)

Dim Result As Variant

'Ok, try doing it
Result = DoTransitions(CurrentState, Inputs)

' Check the results
Select Case Result
    Case -1
        'error
        Debug.Print "Bad input!"
    Case Else
        Debug.Print "Final state: " + Result
End Select

End Sub

Private Sub Command2_Click()
Dim Inputs(4) As Variant

'Setup the input array
Inputs(0) = Determiner(1)
Inputs(1) = Noun(0)
Inputs(2) = Verb(4)
Inputs(3) = Determiner(1)
Inputs(4) = Noun(1)

Dim Result As Variant

'Ok, try doing it
Result = DoTransitions(CurrentState, Inputs)

' Check the results
Select Case Result
    Case -1
        'error
        Debug.Print "Bad input!"
    Case Else
        Debug.Print "Final state: " + Result
End Select


End Sub


Private Sub Form_Load()

Determiner(0) = "a"
Determiner(1) = "the"

Noun(0) = "cat"
Noun(1) = "dog"
Noun(3) = "boy"
Noun(4) = "girl"

Noun(5) = "cake"
Noun(6) = "apple"
Noun(7) = "water"
Noun(8) = "beer"
Noun(9) = "wine"

Verb(0) = "ate"
Verb(1) = "drank"
Verb(2) = "squashed"
Verb(3) = "fondled"
Verb(4) = "chased"



Transitions = 0

MakeTransition "1", Determiner(0), "2"
MakeTransition "1", Determiner(1), "2"
For ni = 0 To 9
    MakeTransition "2", Noun(ni), "3"
Next ni
For vi = 0 To 4
    MakeTransition "3", Verb(vi), "4"
Next vi
MakeTransition "4", Determiner(0), "5"
MakeTransition "4", Determiner(1), "5"
For ni2 = 0 To 9
    MakeTransition "5", Noun(ni2), "6"
Next ni2




CurrentState = "1"

End Sub


