Attribute VB_Name = "Module1"
Option Explicit

Private Function Isx64() As Boolean
    Dim ret As Boolean: ret = False 'èâä˙âª
    If InStr( _
        GetObject("winmgmts:Win32_OperatingSystem=@").OSArchitecture, _
        "64" _
    ) Then ret = True
    Isx64 = ret
End Function

Public Function read_json(ByVal jsonFilePath As String) As Object:
    '#confirm file path exist
    If Dir(jsonFilePath) = "" Then
        Err.Raise _
            Number:=1004, _
            Description:="jsonÉtÉ@ÉCÉãÇ™ë∂ç›ÇµÇ‹ÇπÇÒ: " & jsonFilePath
        
    End If
    '#read json file
    Dim buf As String
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile jsonFilePath
        buf = .ReadText
        .Close
    End With
    
    '#parse json
    Dim obj As Object
    Dim json As Object
    
    '#32bitî≈
    If Isx64 Then
        Set obj = JSFunc.JSFunc("s", "eval('(' + s + ')');")
        
        '#return json object
        Set read_json = obj(buf)
    
    Else
        Set obj = CreateObject("ScriptControl")
        obj.Language = "JScript"
        obj.addcode "function jsonParse(s){ return eval('(' + s + ')'); }"
        
        '#return json object
        Set read_json = obj.CodeObject.jsonParse(buf)
    End If
End Function

Private Sub thisCodeDebug()
    Dim jsonFilePath As String
    Dim config As Object
    
    
    jsonFilePath = ThisWorkbook.Path & "\..\config.json"
    Set config = read_json(jsonFilePath)
    
    Debug.Print CallByName( _
        CallByName(config, "ref", VbGet), _
        "hello", _
        VbGet _
    )
    
    Debug.Print CallByName(config, "ref2", VbGet)
    Dim item
    For Each item In Split(CallByName(config, "ref2", VbGet), ",")
        Debug.Print item & " * 2 = " & Val(item) * 2
    Next
End Sub



