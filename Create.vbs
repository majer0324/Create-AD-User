''''''''''''''''''''''''''''''''''''''''''''
'			
'			roy.chen@sti.com.tw
'

Option Explicit

Dim strPOU,strAD,strDN,strGroup
Dim strCNK
Dim objOU,objUser,objGroup

'�ϥΪ̦W��
Dim strNameK
'�ϥΪ̱K�X
Dim strPwd
'�b�����A
Dim intAccValue
'�Ыرb���ƶq
Dim intUsers
Dim intExist
Dim intErr

Dim objFSO, objInput
Dim strPathAndFile

Const ForReading = 1
Const ForWriting = 2

'�w�q MAC Addresss ���ɮצW�٤θ��|
strPathAndFile = "C:\Temp\mac.csv"

' Open the input file for read access
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objInput = objFSO.OpenTextFile(strPathAndFile, ForReading)

'�w�q�ϥΪ̩Ҧb�� OU
strPOU = "OU=MAC_Group,OU=Xinpu,OU=ITEQ,DC=iteq,DC=corp"
'�w�q�s��
strGroup = "CN=MAC-List,OU=MAC_Group,OU=Xinpu,OU=ITEQ,DC=iteq,DC=corp"
'�]�w�T�w���ϥΪ̱K�X
'strPwd = "Ytrewq12345^"
'�]�w�ϥΪ̪��A
intAccValue = 544

intUsers = 0
intExist = 0
intErr = 0

'�ϥ� LDAP �s���� OU Object
strDN = "LDAP://"&strPOU
Set objOU=GetObject(strDN)

'�ϥ� LDAP �s���� Group Object
strDN = "LDAP://"&strGroup
Set objGroup=GetObject(strDN)

Do Until objInput.AtEndOfStream

'�ϥΪ̪� Distinguished Name
	strNameK = objInput.ReadLine

	If (Trim(strNameK) <> "") Then
		strCNK = "CN="&strNameK
		Set objUser=objOU.Create("user",strCNK)
		objUser.Put "sAMAccountName",strNameK
		objUser.Put "userAccountControl", intAccValue

		On Error Resume Next
		objUser.SetInfo()
		intErr = Err.Number
		
		If intErr <> 0 Then
   		WScript.Echo "Error:   Create User "&strCNK&" failed."
   		intExist = intExist + 1
		Else
			Wscript.Echo "Create User = "&strCNK
'�N�ϥΪ̥[�J�s��
			objGroup.add(objUser.ADsPath)
'�]�w�ϥΪ̱K�X�A�������[�J�s�ի�~�i�H�]�w�K�X
			objUser.SetPassword strNameK
			intUsers = intUsers + 1
		End If
		On Error Goto 0
	 End If
Loop

' Clean up.
objInput.Close
Wscript.Echo	"fail to create "&intExist&" users."
Wscript.Echo	"Total create "&intUsers&" users."
