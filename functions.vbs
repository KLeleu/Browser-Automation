Set WShell = CreateObject("WScript.Shell")
Set Shell = CreateObject("Shell.Application")
Set Taskbar = Shell.NameSpace("shell:::{4234d49b-0245-4df3-b780-3893943456e1}") ' DONT FORGET TO RESEARCH HOW THIS WORKS AND WHERE TO FIND LOCATIONS EASIER.

Sub sleep(t) 
  WScript.Sleep(t)
End Sub

Sub openInbox(i) 
  sleep 200
  WShell.Run "https://mail.google.com/mail/u/"&(i)&"/#inbox"
End Sub

Sub openYoutube()
  sleep 200
  WShell.Run "https://youtube.com"
End Sub

Sub openTwitch()
  sleep 200
  WShell.Run "https://twitch.tv/directory/following"
End Sub

Sub openX()
  sleep 200
  WShell.Run "https://x.com/home/"
End Sub

Sub openPandora()
  sleep 200
  WShell.Run "https://pandora.com/station/play/97778471693990190"
End Sub

Sub openGrok()
  sleep 200
  WShell.Run "https://grok.com/"
End Sub

Sub openChatGPT()
  sleep 200
  WShell.Run "https://chatgpt.com/"
End Sub

  Sub emails() 
   For Each Item in Taskbar.Items
    If Item.Name = "Google Chrome" Then
       openGrok
       openChatGPT
       openInbox 3
       openInbox 2
       openInbox 1
       openInbox 0
    Exit For
   End If
  Next
End Sub

Sub nonWork()  
  For Each Item in Taskbar.Items
    If Item.Name = "Google Chrome" Then
     Item.invokeVerb="open"
      openChatGPT
      openGrok
      openPandora
      openTwitch
      openYoutube
      openX
    Exit For
  End If
  Next
End Sub
