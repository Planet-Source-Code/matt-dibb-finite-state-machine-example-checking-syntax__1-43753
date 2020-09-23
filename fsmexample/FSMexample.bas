Attribute VB_Name = "FSM"

Type Transition
    StartState As Variant
    Input As Variant
    EndState As Variant
End Type


Public Machine(500) As Transition
Public Transitions As Integer

Public CurrentState As Variant
Public Function DoTransition(s As Variant, i As Variant)
j = 0

While j < Transitions

    If Machine(j).StartState = s And Machine(j).Input = i Then
        CurrentState = Machine(j).EndState
        DoTransition = CurrentState
        Exit Function
    
    End If
    j = j + 1

Wend

DoTransition = -1
End Function

Public Function DoTransitions(s As Variant, i() As Variant)
Dim NoMatch As Boolean
Dim Found As Boolean

j = 0
k = 0
s2 = s


While k <= UBound(i)
    While j < Transitions
        If Machine(j).StartState = s2 And Machine(j).Input = i(k) Then
            s2 = Machine(j).EndState
            Debug.Print "Moving to state " + s2
            Found = True
            j = Transitions
        End If
        j = j + 1
    Wend
    
    If Found = False Then
        NoMatch = True
        DoTransitions = -1
        Exit Function
    End If
    
    Found = False
    
    k = k + 1
    j = 0

Wend
DoTransitions = s2
CurrentState = s2
End Function
Public Function MakeTransition(s As Variant, i As Variant, e As Variant)
If Transitions > 500 Then MsgBox "This machine is too big! (500 transitions!)", vbCritical, "Error": Exit Function

Machine(Transitions).StartState = s
Machine(Transitions).Input = i
Machine(Transitions).EndState = e

Transitions = Transitions + 1
End Function


