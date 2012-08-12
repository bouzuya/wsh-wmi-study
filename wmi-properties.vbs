Option Explicit

Private Function GetNamedArguments(ByVal strName, ByVal strDefault)
    If WScript.Arguments.Named.Exists(strName) Then
        GetNamedArguments = WScript.Arguments.Named.Item(strName)
    Else
        GetNamedArguments = strDefault
    End If
End Function

Private Function FormatProperty(ByVal objProperty)
    Dim strMessage
    strMessage = ""
    strMessage = strMessage & PadRight(objProperty.Name, 32, " ")
    strMessage = strMessage & " : "
    If objProperty.IsArray Then
        If IsNull(objArray) Then
            strMessage = strMessage & ""
        Else
            strMessage = strMessage & Join(objArray, ",")
        End If
    Else
        strMessage = strMessage & objProperty.Value
    End If
    FormatProperty = strMessage
End Function

Private Sub EchoArguments(ByVal strServer, ByVal strNamespace, ByVal strClassName, ByVal strPropertyName)
    Dim strArgs
    strArgs = ""
    strArgs = strArgs & "Query info" & vbCrLf
    strArgs = strArgs & "  Server       : " & strServer & vbCrLf
    strArgs = strArgs & "  Namespace    : " & strNamespace & vbCrLf
    strArgs = strArgs & "  ClassName    : " & strClassName & vbCrLf
    strArgs = strArgs & "  PropertyName : " & strPropertyName & vbCrLf
    WScript.Echo(strArgs)
End Sub

Private Sub EchoProperties(ByVal strServer, ByVal strNamespace, ByVal strClassName, ByVal strPropertyName)
    Dim strQuery
    strQuery = "SELECT * FROM " & strClassName
    Dim objLocator, objServices, objObjectSet, objObject, objProperty, intIndex
    Set objLocator = WScript.CreateObject("WbemScripting.SWbemLocator")
    Set objServices = objLocator.ConnectServer(strServer, strNamespace)
    Set objObjectSet = objServices.ExecQuery(strQuery)
    intIndex = 0
    For Each objObject In objObjectSet
        Dim strMessage
        strMessage = "[" & CStr(intIndex) & "]" & vbCrLf
        For Each objProperty In objObject.Properties_
            If strPropertyName = "" Or _
                InStr(1, objProperty.Name, strPropertyName, vbBinaryCompare) > 0 Then
                strMessage = strMessage & FormatProperty(objProperty) & vbCrLf
            End If
        Next
        WScript.Echo(strMessage)
        intIndex = intIndex + 1
    Next
End Sub

Function Main()
    ' Get arguments
    Dim strServer, strNamespace, strClassName, strPropertyName
    strServer = GetNamedArguments("Server", ".")
    strNamespace = GetNamedArguments("Namespace", "root\cimv2")
    strClassName = GetNamedArguments("ClassName", "")
    strPropertyName = GetNamedArguments("PropertyName", "")

    ' Check arguments
    If strClassName = "" Then
        WScript.Echo("ClassName is required.")
        Main = 1
        Exit Function
    End If

    Call EchoArguments(strServer, strNamespace, strClassName, strPropertyName)
    Call EchoProperties(strServer, strNamespace, strClassName, strPropertyName)

    Main = 0
    Exit Function
End Function

WScript.Quit(Main())

