<%
'기본 세팅 설정

'//////////////////////////////////////////////////////////////////////////////////
'============= DB Connection Basic Info ======================
Dim strDBServerIP
Dim strDBName
Dim strDBUserID
Dim strDBPassword
Dim strDBConnString
Dim strDBConnString2

'============= DB Connection ===============================
'strDBServerIP = "192.168.1.163"
'strDBServerIP = "210.102.151.105"
'//레알
'strDBServerIP = "210.102.151.72,4508"
'//테스트

'//사내 데이터베이스 사용 시
'strDBServerIP = "MMSMISS"
'strDBName = "TEST"
'strDBUserID = "xteb"
'strDBPassword	= "east12!@"

'//개인 데이터베이스 사용 시
strDBServerIP = "DESKTOP-V7SJLLF"
strDBName = "DailySupport_DB"
strDBUserID = "DailySupport"
strDBPassword = "Harang508!"

strDBConnString = "Provider=SQLOLEDB;Data Source=" & strDBServerIP & ";Initial Catalog=" & strDBName & ";user ID=" & strDBUserID & ";password=" & strDBPassword & ";"
'=======================================================
%>