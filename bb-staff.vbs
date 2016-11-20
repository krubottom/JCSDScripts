'On Error Resume next
Set fso = CreateObject("Scripting.FileSystemObject")

'binding to the object
Set objOU=GetObject ("LDAP://OU=Employees,DC=JCSD1,DC=US")

WScript.Echo "username,lastname,firstname,email,password"

For each OU in objOU
	'On Error Resume Next
	Select Case OU.Class
		Case "user"
			'WScript.Echo OU.name & " - " & objOU.Class
			StudnetInfo(OU.distinguishedName)
		Case "organizationalUnit" , "container"
			'WScript.Echo OU.distinguishedName & " - " & objOU.Class
			ShowSubOU OU.distinguishedName
	End Select
Next




Function ShowSubOU(OUNAME)

'WScript.Echo OUNAME
Set objSubOU=GetObject("LDAP://" & OUNAME)

For each SubOU in objSubOU
	Select Case SubOU.Class
		Case "user"
			'WScript.Echo subOU.name & " - " & SubOU.Class
			StudnetInfo(subOU.distinguishedName)
		Case "organizationalUnit" , "container"
			'WScript.Echo subOU.name & " - " & SubOU.Class
			ShowSubOU subOU.distinguishedName
	End Select
Next
End Function


Function StudnetInfo(UserToBeChecked)

Set objUser = GetObject("LDAP://" & UserToBeChecked)
On Error Resume Next
WScript.Echo objUser.samAccountName & "," & objUser.sn & "," & objUser.givenName & "," & objUser.samAccountName & "@jcsd1.epals.com" & ",password"
End Function