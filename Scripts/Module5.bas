Attribute VB_Name = "Module5"
' Declare a module-level variable to store the calculated sag value
Dim calculatedSagValue As Double

Private Sub saginput_Change()
    Dim originalValue As Double
    Dim result As Double

    ' ... (rest of your code) ...

    result = originalValue * 1.02 ' Adding 2%

    ' Store the result in the module-level variable
    calculatedSagValue = result

    ' ... (rest of your code) ...
End Sub

' Example of how to access the stored value elsewhere in your code:
Private Sub someOtherSubroutine()
    MsgBox "Calculated Sag Value: " & calculatedSagValue
End Sub
