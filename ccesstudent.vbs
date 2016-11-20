On Error Resume Next

'Need to fix code so will work for all buildings, clean up and other crap

Set fso = Wscript.CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Wscript.Shell")
Dim strLine
Dim attribArray
Set txtUsers=fso.OpenTextFile("c:\cces.csv")
Const ID = 7
Const First = 1
Const Last = 2
Const Passwd = 6
Const UserName = 5
Const Grade = 4
Const ADS_SCOPE_SUBTREE = 2

Do While Not txtUsers.AtEndOfStream
strLine = txtUsers.ReadLine
attribArray=split(strLine, ",")

'WScript.Echo attribArray(UserName)

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection

objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 

objCommand.CommandText = "SELECT distinguishedName FROM 'LDAP://dc=jcsd1,dc=us' WHERE objectCategory='user' " & "AND sAMAccountName = '" & attribArray(UserName) & "'" 
Set objRecordSet = objCommand.Execute

objRecordSet.MoveFirst
Do Until objRecordSet.EOF
    strDN = objRecordSet.Fields("distinguishedName").Value
    objRecordSet.MoveNext
Loop

Set objUser = GetObject("LDAP://" & strDN)
'Wscript.Echo objUser.Name




        user.SetPassword attribArray(Passwd) 'Uncomment for file with real psswords, not MDLK


        objUser.SetInfo
	

        WScript.Echo objUser.Name & " - " & attribArray(Passwd)

Loop

