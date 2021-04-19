Dim fso, f, f1, fc
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.GetFolder("\\192.168.1.1\Documents")
Set fc = f.Files

For Each f1 in fc
	If DateDiff("d", f1.DateLastModified, Now) > 20 Then
		f1.Delete True
	End If
Next