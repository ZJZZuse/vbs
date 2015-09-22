' set ws=createobject("Wscript.Shell")
' ws.run "regsvr32 dm.dll /s"


set dm = createobject("dm.dmsoft")

dm.SendString 203616, "asdf"

