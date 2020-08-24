Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists("cancion") Then
c=C+1
l=1
Else
end if

c=0
Set shell=CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
If Not (fso.FileExists("cancion\1.txt") ) Then
m=1
c=C+1
Else
END If
Set shell=CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
If Not (fso.FileExists("cancion\m.bat") ) Then
v=1
c=C+1
Else
END IF



Set shell=CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
If Not (fso.FileExists("cancion\Musica.vbs") ) Then
y=1
c=C+1
Else
END IF
Set shell=CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
If Not (fso.FileExists("cancion\1.mp3") ) Then
j=1
c=C+1
Else
END IF
Set shell=CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
If Not (fso.FileExists("cancion\2.mp3") ) Then
f=1
c=C+1
Else
END IF
Set shell=CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
If Not (fso.FileExists("cancion\3.mp3") ) Then
d=1
c=C+1
Else
END IF
Set shell=CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
If Not (fso.FileExists("cancion\4.mp3") ) Then
s=1
c=C+1
Else
END IF
Set shell=CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
If Not (fso.FileExists("cancion\5.mp3") ) Then
t=1
c=C+1
Else
END IF

If (c>0) Then
result = MsgBox ("Faltan " & c & " archibos quieres descargarlos? (si no lo haces no funcionara el reproductor)", vbYesNo, "Reproductor de musica por Mrgaton")
Select Case result
Case vbYes

If (l=1) Then
fso.CreateFolder ("cancion")
wscript.sleep 200
End If

If (t=1) Then
dim http_obj
dim stream_obj
dim shell_obj
set http_obj = CreateObject("Microsoft.XMLHTTP")
set stream_obj = CreateObject("ADODB.Stream")
set shell_obj = CreateObject("WScript.Shell")
URL = "https://cdn.discordapp.com/attachments/592374030816772129/747429482192699423/5.mp3"
FILENAME = "cancion\5.mp3"
http_obj.open "GET", URL, False
http_obj.send
stream_obj.type = 1
stream_obj.open
stream_obj.write http_obj.responseBody
stream_obj.savetofile FILENAME, 2
wscript.sleep 200
End If

If (s=1) Then

set http_obj = CreateObject("Microsoft.XMLHTTP")
set stream_obj = CreateObject("ADODB.Stream")
set shell_obj = CreateObject("WScript.Shell")
URL = "https://cdn.discordapp.com/attachments/592374030816772129/747435110164332585/4.mp3"
FILENAME = "cancion\4.mp3"
http_obj.open "GET", URL, False
http_obj.send
stream_obj.type = 1
stream_obj.open
stream_obj.write http_obj.responseBody
stream_obj.savetofile FILENAME, 2
wscript.sleep 200
End If

If (d=1) Then

set http_obj = CreateObject("Microsoft.XMLHTTP")
set stream_obj = CreateObject("ADODB.Stream")
set shell_obj = CreateObject("WScript.Shell")
URL = "https://cdn.discordapp.com/attachments/592374030816772129/747435329597866094/3.mp3"
FILENAME = "cancion\3.mp3"
http_obj.open "GET", URL, False
http_obj.send
stream_obj.type = 1
stream_obj.open
stream_obj.write http_obj.responseBody
stream_obj.savetofile FILENAME, 2
wscript.sleep 200
End If

If (f=1) Then

set http_obj = CreateObject("Microsoft.XMLHTTP")
set stream_obj = CreateObject("ADODB.Stream")
set shell_obj = CreateObject("WScript.Shell")
URL = "https://cdn.discordapp.com/attachments/592374030816772129/747436117456060476/2.mp3"
FILENAME = "cancion\2.mp3"
http_obj.open "GET", URL, False
http_obj.send
stream_obj.type = 1
stream_obj.open
stream_obj.write http_obj.responseBody
stream_obj.savetofile FILENAME, 2
wscript.sleep 200
End If

If (j=1) Then

set http_obj = CreateObject("Microsoft.XMLHTTP")
set stream_obj = CreateObject("ADODB.Stream")
set shell_obj = CreateObject("WScript.Shell")
URL = "https://cdn.discordapp.com/attachments/592374030816772129/747436222409867326/1.mp3"
FILENAME = "cancion\1.mp3"
http_obj.open "GET", URL, False
http_obj.send
stream_obj.type = 1
stream_obj.open
stream_obj.write http_obj.responseBody
stream_obj.savetofile FILENAME, 2
wscript.sleep 200
End If

If (y=1) Then

set http_obj = CreateObject("Microsoft.XMLHTTP")
set stream_obj = CreateObject("ADODB.Stream")
set shell_obj = CreateObject("WScript.Shell")
URL = "https://cdn.discordapp.com/attachments/592374030816772129/747451304711815288/Musica.vbs"
FILENAME = "cancion\Musica.vbs"
http_obj.open "GET", URL, False
http_obj.send
stream_obj.type = 1
stream_obj.open
stream_obj.write http_obj.responseBody
stream_obj.savetofile FILENAME, 2
wscript.sleep 100
End If
Dim f1
If (v=1) Then

Set f1 = fso.CreateTextFile("cancion\m" & ".bat", True)
f1.WriteLine("12")
f1.Close
End If
If (m=1) Then

Set f1 = fso.CreateTextFile("cancion\1" & ".txt", True)
f1.WriteLine("12")
f1.Close
End If


Set objShell = CreateObject("Wscript.Shell")
objShell.Run "cancion\Musica.vbs"
Case vbNo
Set objShell = CreateObject("Wscript.Shell")
objShell.Run "cancion\Musica.vbs"
End Select

End If
If (c=0) Then
Set objShell = CreateObject("Wscript.Shell")
objShell.Run "cancion\Musica.vbs"
End If