Public Enum ResponseFormat
    Text
    Json
End Enum
Private pResponseText As String
Private pResponseJson
Private pScriptControl As Object
Public Function request(url As String, Optional postParameters As String = "", Optional format As ResponseFormat = ResponseFormat.Json) As String
    Dim xml
    Dim requestType As String
    If postParameters <> "" Then
        requestType = "POST"
    Else
        requestType = "GET"
    End If

    Set xml = CreateObject("MSXML2.XMLHTTP")
    xml.Open requestType, url, False
    xml.setRequestHeader "Content-Type", "application/json"
    xml.setRequestHeader "Accept", "application/json"
    USER = [USUARIO] 'Llama a una celda "USUARIO" que usará como usuario de autenticación
    Password = [CONTRA] 'Llama a una celda "CONTRA" que usará como contraseña
    xml.setRequestHeader "Authorization", "Basic " + USER + ":" + Password
    If postParameters <> "" Then
        xml.send (postParameters)
    Else
        xml.send
    End If
    pResponseText = xml.ResponseText
    request = pResponseText
    Select Case format
        Case Json
            SetJson
    End Select
End Function
Private Sub SetJson()
    Dim qt As String
    qt = """"
    Set pScriptControl = CreateObjectx86("ScriptControl")
    pScriptControl.Language = "JScript"
    pScriptControl.eval "var obj=(" & pResponseText & ")"
    pScriptControl.AddCode "function getObject(){return obj;}"
    pScriptControl.AddCode "function getRootObject(){return rootObj;}"
    pScriptControl.AddCode "function getCount(){ return rootObj.length;}"
    pScriptControl.AddCode "function getBaseValue(){return baseValue;}"
    pScriptControl.AddCode "function getValue(){ return arrayValue;}"
    Set pResponseJson = pScriptControl.Run("getObject")
End Sub
Function CreateObjectx86(Optional sProgID)

    Static oWnd As Object
    Dim bRunning As Boolean

    #If Win64 Then
        bRunning = InStr(TypeName(oWnd), "HTMLWindow") > 0
        Select Case True
            Case IsMissing(sProgID)
                If bRunning Then oWnd.Lost = False
                Exit Function
            Case IsEmpty(sProgID)
                If bRunning Then oWnd.Close
                Exit Function
            Case Not bRunning
                Set oWnd = CreateWindow()
                oWnd.execScript "Function CreateObjectx86(sProgID): Set CreateObjectx86 = CreateObject(sProgID) End Function", "VBScript"
                oWnd.execScript "var Lost, App;": Set oWnd.App = Application
                oWnd.execScript "Sub Check(): On Error Resume Next: Lost = True: App.Run(""CreateObjectx86""): If Lost And (Err.Number = 1004 Or Err.Number = 0) Then close: End If End Sub", "VBScript"
                oWnd.execScript "setInterval('Check();', 500);"
        End Select
        Set CreateObjectx86 = oWnd.CreateObjectx86(sProgID)
    #Else
        Set CreateObjectx86 = CreateObject(sProgID)
    #End If

End Function

Function CreateWindow()

    ' source http://forum.script-coding.com/viewtopic.php?pid=75356#p75356
    Dim sSignature, oShellWnd, oProc

    On Error Resume Next
    sSignature = Left(CreateObject("Scriptlet.TypeLib").GUID, 38)
    CreateObject("WScript.Shell").Run "%systemroot%\syswow64\mshta.exe about:""<head><script>moveTo(-32000,-32000);document.title='x86Host'</script><hta:application showintaskbar=no /><object id='shell' classid='clsid:8856F961-340A-11D0-A96B-00C04FD705A2'><param name=RegisterAsBrowser value=1></object><script>shell.putproperty('" & sSignature & "',document.parentWindow);</script></head>""", 0, False
    Do
        For Each oShellWnd In CreateObject("Shell.Application").Windows
            Set CreateWindow = oShellWnd.GetProperty(sSignature)
            If Err.Number = 0 Then Exit Function
            Err.Clear
        Next
    Loop

End Function
Public Function setJsonRoot(rootPath As String)
    If rootPath = "" Then
        pScriptControl.ExecuteStatement "rootObj = obj"
    Else
        pScriptControl.ExecuteStatement "rootObj = obj." & rootPath
    End If
    Set setJsonRoot = pScriptControl.Run("getRootObject")
End Function
Public Function getJsonObjectCount()
    getJsonObjectCount = pScriptControl.Run("getCount")
End Function
Public Function getJsonArrayValue(index, Key As String)
    Dim qt As String
End Function
Public Function getJsonObjectValue(path As String)
    pScriptControl.ExecuteStatement "baseValue = obj." & path
    qt = """"
    If InStr(Key, ".") > 0 Then
        arr = Split(Key, ".")
        Key = ""
        For Each cKey In arr
            Key = Key + "[" & qt & cKey & qt & "]"
        Next
    Else
        Key = "[" & qt & Key & qt & "]"
    End If
    Dim statement As String
    statement = "arrayValue = rootObj[" & index & "]" & Key

    pScriptControl.ExecuteStatement statement
    getJsonArrayValue = pScriptControl.Run("getValue", index, Key)
End Function
Public Property Get ResponseText() As String
    ResponseText = pResponseText
End Property
Public Property Get ResponseJson()
    ResponseJson = pResponseJson
End Property
Public Property Get ScriptControl() As Object
    ScriptControl = pScriptControl
End Property
