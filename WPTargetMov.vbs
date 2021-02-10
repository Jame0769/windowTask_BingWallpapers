'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'Filename	 : WPTargetDir.vbs
'Description 	 : Opens the target folder of current desktop background in Windows 7
'Author   	 : Ramesh Srinivasan, ex-MSMVP (Windows)
'Website 	 : http://www.winhelponline.com/blog
'Created 	 : Feb 02, 2010
'Modified	 : Jan 10, 2016
'?Ramesh Srinivasan.
'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
Set WshShell = WScript.CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
strMsg = "Completed!" & Chr(10) & Chr(10) & "WPTargetDir.vbs - ?2016 Ramesh Srinivasan" & Chr(10) & Chr(10) & "Visit us at http://www.winhelponline.com/blog"
strCurWP =""

On Error Resume Next
strCurWP = WshShell.RegRead("HKCU\Software\Microsoft\Internet Explorer\Desktop\General\WallpaperSource")
On Error Goto 0

If Trim(strCurWP) = "" Then
	MsgBox "No Wallpaper selected for Desktop Slideshow"
Else
	If fso.FileExists(strCurWP) Then
		'Modified the code here
		'WshShell.run "explorer.exe" & " /select," & Chr(34) & Trim(strCurWP) & Chr(34)
		if msgbox ("Move current wall paper?") = vbOK  then
			'fso.DeleteFile(strCurWP)
			call fso.MoveFile(strCurWP, "C:\Users\JameHe\AppData\Local\Microsoft\Windows\Themes\Collect\")
		end if 
	Else
		MsgBox "Cannot find target for " & strCurWP
	End If
End If

'""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
	'This script was brought to you by "The Winhelponline Blog"
	'Visit us at http://www.winhelponline.com/blog

'""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
