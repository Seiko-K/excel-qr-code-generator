Option Explicit

' Returns True when running on macOS.
' macOSで実行中の場合はTrueを返します。
Public Function IsMac() As Boolean

#If Mac Then
    IsMac = True
#Else
    IsMac = False
#End If

End Function