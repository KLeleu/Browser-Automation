Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.OpenTextFile("C:\Users\Kevin\Documents\Programming Guides, Tips, & Tricks\Browser Automation\functions.vbs", 1)
content = file.ReadAll
Execute content
file.Close

openBrowser()