dim oShell
dim key1, key2

Set oShell = WScript.CreateObject("WScript.Shell")

' 規定値を取得する場合（最後に\をつける）
key1 = oShell.RegRead("HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Access\Security\Trusted Locations\Location2\")

' キーを取得する場合
key2 = oShell.RegRead("HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Access\Security\Trusted Locations\Location2\path")

WScript.Echo "key1: " & key1 & vbCrLf & _ 
             "key2: " & key2
