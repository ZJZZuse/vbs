
' MsgBox 1

set dm = createobject("dm.dmsoft")

ver = dm.Ver()

' if ver = null then

' ' 这时，已经确认插件注册失败了。 弹出一些调试信息，以供分析.

' 	MsgBox "请关闭程序,重新打开本程序再尝试"

' else

' 	MsgBox ver

' end if

dim hWnd


' set hWnd = dm.FindWindow("", "记事本")

dim hwnds

' hwnds = dm.EnumWindow(0,"记事本","",1+4+8+16)

MsgBox hwnds

' dm_ret = Dm.SetWindowText(hWnd, "test")

' Dm.SendString(hWnd, "我是来测试的")


' dm.SetWindowTransparent(hWnd,100) 

' Debug.Print hWnd

Sub OnScriptExit()

   dm.UnBindWindow

End Sub



