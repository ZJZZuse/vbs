Sub include(file)
    Dim fso, f, str
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(file, 1)
    str = f.ReadAll
    f.Close
    ExecuteGlobal str
End Sub

include "VbsJson.vbs"


Dim fso, json, str, o, i
Set json = New VbsJson
Set fso = createobject("Scripting.Filesystemobject")
str = fso.OpenTextFile("json.txt").ReadAll
Set o = json.Decode(str)
WScript.Echo o("Image")("Width")
WScript.Echo o("Image")("Height")
WScript.Echo o("Image")("Title")
WScript.Echo o("Image")("Thumbnail")("Url")
For Each i In o("Image")("IDs")
    WScript.Echo i
Next