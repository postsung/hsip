<% @CodePage = 65001 %>
<!-- #include file="include/dbcon.asp" -->
<!-- #include file="include/Session_chk.asp" -->
<!-- #include file="include/function.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
Session.CodePage = 65001
Server.ScriptTimeout = 7600

Dim Sql, Rs, i
Dim sPaper, sMemberID
Dim StartYear, StartMonth, EndYear, EndMonth, EndDay
Dim StartDate, EndDate

Dim Fs, objFile

StartYear = Left(Request("StartYear"), 10)
StartMonth = Left(Request("StartMonth"), 10)
EndYear = Left(Request("EndYear"), 10)
EndMonth = Left(Request("EndMonth"), 10)

EndDay = Day(DateSerial(EndYear, EndMonth+1, 0))
StartDate = StartYear & "-" & StartMonth & "-1"
EndDate = DateSerial(EndYear, EndMonth+1, 0)


Set Fs = CreateObject("Scripting.FileSystemObject")

Call Fs.CopyFile(Server.Mappath("./") & "\xml\result_part_exam.xml", Server.Mappath("./") & "\Temp\" & Session.SessionID & ".xml")
Set objFile = Fs.OpenTextFile(Server.Mappath("./") & "\Temp\" & Session.SessionID & ".xml", 8) 

objFile.WriteLine "<w:body>"

Dim ListCnt
Dim ApplNo, YourRef, OurRef, FileDate, PCTFileDate, DDRExam


'Call sbTopTitleLine()

Sql = "SELECT Field42, Field6, Field5, Field41, Field128, Field49 "
Sql = Sql & "FROM LeftMenu0001 "
Sql = Sql & "WHERE ( '" & StartDate & "' <= Field49 AND Field49 <= '" & EndDate & "' ) "


Set Rs = oConn.Execute(Sql)
ListCnt = 1
Do Until Rs.EOF '사건수만큼 루프
	
	ApplNo = Rs.Fields(0)
	YourRef = Rs.Fields(1)
	OurRef = Rs.Fields(2)
	FileDate = Rs.Fields(3)
	PCTFileDate = Rs.Fields(4)
	DDRExam = Rs.Fields(5)
	Rs.MoveNext

	If ListCnt > 25 Then '리스트 25건최과시 페이지이동
		objFile.WriteLine "</w:tbl>"
'		Call sbNextPage()
'		Call sbTopTitleLine()
		ListCnt = 1
	End If

'	Call sbCommissionLine(arMemberPriceResult)
	
	ListCnt = ListCnt + 1

Loop

Rs.Close
Set Rs = Nothing	

'  Call sbTopBottomLine()

objFile.Close

Dim Cks
Set Cks = Server.CreateObject("CkString.CkString")

Cks.LoadFile Server.Mappath("./") & "\temp\" & Session.SessionID & ".xml", "ANSI"
Cks.SaveToFile Server.Mappath("./") & "\Temp\" & Session.SessionID & Session.SessionID & ".doc","utf-8"

Set Cks = Nothing
Set Fs = Nothing


Set objFile = Nothing
Set Fs = Nothing
oConn.Close
Set oConn = Nothing




%>
