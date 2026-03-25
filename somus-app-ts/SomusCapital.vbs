Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = "C:\Users\Artur Brito\Documents\APP SOMUS\somus-app-ts"
WshShell.Run "cmd /c start /min ""Somus v2"" node node_modules\vite\bin\vite.js --host", 0, False
WScript.Sleep 4000
WshShell.Run "http://localhost:5173", 1, False
