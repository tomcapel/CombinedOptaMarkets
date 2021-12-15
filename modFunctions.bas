Attribute VB_Name = "modFunctions"


Public Function StringMatch(stringOne As String, stringTwo As String) As Boolean

    StringMatch = IIf(InStr(stringOne, stringTwo), True, False)

End Function
