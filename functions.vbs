Set WShell = CreateObject("WScript.Shell")
Set Shell = CreateObject("Shell.Application")
Set Taskbar = Shell.NameSpace("shell:::{4234d49b-0245-4df3-b780-3893943456e1}") 'DONT FORGET TO RESEARCH HOW THIS WORKS AND WHERE TO FIND LOCATIONS EASIER.

Sub sleep(t) 
  WScript.Sleep(t)
End Sub

Sub openInbox(i) 
  sleep 200
  WShell.Run "https://mail.google.com/mail/u/"&(i)&"/#inbox"
End Sub

Sub openYoutube()
  sleep 200
  WShell.Run "http://youtube.com"
End Sub

Sub openTwitch()
  sleep 200
  wShell.Run "https://twitch.tv/directory/following"
End Sub

Sub openX()
  sleep 200
  WShell.Run "http://x.com/home/"
End Sub

Sub openPandora()
  sleep 200
  WSHell.Run "https://pandora.com/station/play/97778471693990190"
End Sub
	

Sub openEmails() 
  For Each Item in Taskbar.Items
   If Item.Name = "Google Chrome" Then
      Item.InvokeVerb "open"
      openInbox 0
      openInbox 1
      openInbox 2
      openInbox 3
    Exit For
  End If
  Next
End Sub


Sub openBrowser()  
  For Each Item in Taskbar.Items
    If Item.Name = "Google Chrome" Then
      Item.InvokeVerb "Open"
      openPandora
      openX
      openTwitch
      openYoutube
    Exit For
  End If
  Next
End Sub