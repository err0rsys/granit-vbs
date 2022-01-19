Option Explicit 'must declare variables

Dim fso : Set fso = CreateObject("scripting.filesystemobject")
ExecuteGlobal fso.opentextfile("md5.vbs", 1).ReadAll
Set fso = Nothing

Dim remote
remote = InputBox("Enter server ip or name:","APPLICATION SERVER","localhost")

Dim login, password
login = InputBox("Enter username:", "GRANIT LOGON", "ADMINISTRATOR")
password = InputBox("Enter password for " & LOGIN & ":", "GRANIT LOGON", "???")

'hash function
password = MD5(password)

Dim db
db = InputBox("Enter database name:","GRANIT LOGON","GRANIT")

Dim obj
Set obj = CreateObject("DBC.DBCOM",remote)

obj.AssignAccessCode("LOGON")
obj.SQL("SELECT BazaDanychRzeczywistaId FROM BazaDanychRzeczywista WHERE DatabaseName='" & db &"'")
obj.TimeOut(-1)
obj.OpenA("List")

Dim i
Dim ovData
i = -1
i = obj.GetPacketNumber(i, ovData)

If i > 0 Then
  Dim id
  id = ovData(1,0)

  Dim wsh
  Set wsh = WScript.CreateObject("WScript.Shell")

  Dim Params(18,1)

  Params(0,0)=1 'accesscode
  Params(1,0)=2 'clientkey
  Params(2,0)=3 'enterprise
  Params(3,0)=4 'dbindex
  Params(4,0)=5 'languages
  Params(5,0)=6 'login
  Params(5,1)=login
  Params(6,0)=7 'password
  Params(6,1)=password
  Params(7,0)=8 'deploymentid
  Params(7,1)=0
  Params(8,0)=9 'dbid
  Params(8,1)=id
  Params(9,0)=10 'language
  Params(9,1)=3  '1-polish,2-german,3-english
  Params(10,0)=11 'lifetime
  Params(11,0)=12 'computername
  Params(11,1)=wsh.ExpandEnvironmentStrings("%COMPUTERNAME%")
  Params(12,0)=13 'computerusername
  Params(12,1)=wsh.ExpandEnvironmentStrings("%USERNAME%")
  Params(13,0)=14 'messagepack
  Params(14,0)=15 'passchange
  Params(15,0)=16 'message
  Params(16,0)=17 'worktype
  Params(16,1)=0
  Params(17,0)=19 'appname
  Params(17,1)="vbs"
  Params(18,0)=20 'systemlogin
  Params(18,1)=1

  Set obj = CreateObject("GateKeeperComEx.Manager",remote)
  obj.AssignAccessCode("XXX")

  i = obj.LoginClient(2, Params)

  If i > 0 Then
  	 Dim AC
  	 AC = Params(0,1)
  	 MsgBox "Logged to GRANIT with AccessCode: " & Chr(13) & Chr(10) & AC

  	 Set obj = CreateObject("Script.ScriptCOM",remote)
         obj.AssignAccessCode(AC)

     Dim Par
     Par = Array("/ResultTypeId=5", Nothing, Nothing)
  	 i = obj.RunMethod("ScriptHelp", Par)

  	 MsgBox "Method result: " & CStr(i) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & CStr(Par(2))

  ElseIf i = -12020 Then
     MsgBox "Wrong login or password!"
  Else
     MsgBox "Cannot login to GRANIT, errorcode: " & CStr(i)
  End If
Else
	MsgBox "Database " & db & " not found!", 64, "Information"
End If