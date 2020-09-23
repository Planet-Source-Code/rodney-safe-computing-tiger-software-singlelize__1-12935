Attribute VB_Name = "modSingle"
Option Explicit

Public Sub Singlelize(ListObject As Object)
    Dim I As Integer
    Dim X As Integer
    
    For I = 0 To ListObject.ListCount - 1
        For X = 0 To ListObject.ListCount - 1
            If I <> X Then
                If LCase(ListObject.List(X)) = LCase(ListObject.List(I)) Then
                    ListObject.RemoveItem X
                    X = I
                End If
            End If
        Next X
    Next I
End Sub
