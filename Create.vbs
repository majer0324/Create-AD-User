''''''''''''''''''''''''''''''''''''''''''''
'			
'			roy.chen@sti.com.tw
'

Option Explicit

Dim strPOU,strAD,strDN,strGroup
Dim strCNK
Dim objOU,objUser,objGroup

'使用者名稱
Dim strNameK
'使用者密碼
Dim strPwd
'帳號狀態
Dim intAccValue
'創建帳號數量
Dim intUsers
Dim intExist
Dim intErr

Dim objFSO, objInput
Dim strPathAndFile

Const ForReading = 1
Const ForWriting = 2

'定義 MAC Addresss 的檔案名稱及路徑
strPathAndFile = "C:\Temp\mac.csv"

' Open the input file for read access
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objInput = objFSO.OpenTextFile(strPathAndFile, ForReading)

'定義使用者所在的 OU
strPOU = "OU=MAC_Group,OU=Xinpu,OU=ITEQ,DC=iteq,DC=corp"
'定義群組
strGroup = "CN=MAC-List,OU=MAC_Group,OU=Xinpu,OU=ITEQ,DC=iteq,DC=corp"
'設定固定的使用者密碼
'strPwd = "Ytrewq12345^"
'設定使用者狀態
intAccValue = 544

intUsers = 0
intExist = 0
intErr = 0

'使用 LDAP 連結到 OU Object
strDN = "LDAP://"&strPOU
Set objOU=GetObject(strDN)

'使用 LDAP 連結到 Group Object
strDN = "LDAP://"&strGroup
Set objGroup=GetObject(strDN)

Do Until objInput.AtEndOfStream

'使用者的 Distinguished Name
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
'將使用者加入群組
			objGroup.add(objUser.ADsPath)
'設定使用者密碼，必須先加入群組後才可以設定密碼
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
