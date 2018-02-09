<% @CodePage = 949 %>
<!-- #include file="include/dbcon.asp" -->
<!-- #include file="include/Session_chk.asp" -->
<!-- #include file="include/function.asp" -->
<%
Response.ContentType = "application/msword" 
Response.AddHeader "Content-Disposition","attachment;filename=Æ¯ÇãÁý°èÇ¥.doc"

Dim Sql, Rs, Rs2, i, k
Dim TodayEng, PageLoop, PageCnt
Dim CustNum

CustNum = Int(Request("CustNum"))

TodayEng = fnDateReplaceEnd(Date)

'°í°´Á¤º¸
Dim CustName, Email, LetterAdd1, LetterAdd2, LetterAdd3, LetterAdd4, LetterAdd5
Sql = "SELECT mName, Field9, ISNULL(Field46,''), ISNULL(Field47,''), ISNULL(Field48,''), ISNULL(Field49,''), ISNULL(Field50,'') "
Sql = Sql & "FROM Customer WHERE Num = " & CustNum
Set Rs = oConn.Execute(Sql)
CustName = Rs.Fields(0)
Email = Rs.Fields(1)
LetterAdd1 = Rs.Fields(2)
LetterAdd2 = Rs.Fields(3)
LetterAdd3 = Rs.Fields(4)
LetterAdd4 = Rs.Fields(5)
LetterAdd5 = Rs.Fields(6)
Rs.Close
Set Rs = Nothing

Dim nDate, sDate
Dim sDateS, sDateE
nDate = Year(Date) '³¡³âµµ
sDate = nDate-9 '½ÃÀÛ³âµµ

'´ëÇ¥Ãâ¿øÀÎ È®ÀÎ
Dim appCustNum(1), appCustName(1), appCustRef(1)
Dim appCustNumCnt
i = 0
Sql = "SELECT TOP 2 CustomerCode, (SELECT mName FROM Customer WHERE Num = D.CustomerCode), (SELECT Field51 FROM Customer WHERE Num = D.CustomerCode) " &_
		"FROM CustomerCode D WHERE TableNum IN( " &_
		"SELECT num FROM LeftMenu0001 WHERE Field41 <= '" & nDate & "-12-31' AND (PatCode = 'A01' OR PatCode = 'A05') " &_
		"AND Num IN(SELECT TableNum FROM CustomerCode WHERE CustomerCode = " & CustNum & " AND (FieldCount = 7 OR FieldCount = 34) ) "&_
		") AND FieldCount = 37 GROUP BY CustomerCode ORDER BY COUNT(*) DESC "
Set Rs = oConn.Execute(Sql)
Do Until Rs.EOF
	appCustNum(i) = Rs.Fields(0)
	appCustName(i) = Rs.Fields(1)
	appCustRef(i) = Rs.Fields(2)
	i = i + 1
	Rs.MoveNext
Loop
Rs.Close
Set Rs = Nothing

appCustNumCnt = i

Dim rsYear, rsCount

'´ëÇ¥Ãâ¿øÀÎ1
Dim arStatisApplName1(9)
Dim sumStatisApplName1

Call sbArrayInitial(arStatisApplName1)
sumStatisApplName1 = 0
Sql = "SELECT YEAR(Field41), COUNT(*) FROM LeftMenu0001 WHERE Field41 <= '" & nDate & "-12-31' "
Sql = Sql & "AND (PatCode = 'A01' OR PatCode = 'A05') "
Sql = Sql & "AND Num IN(	 "
Sql = Sql & "SELECT TableNum FROM CustomerCode WHERE TableNum IN( " &_		
					"SELECT TableNum FROM CustomerCode WHERE CustomerCode = " & CustNum & " AND (FieldCount = 7 OR FieldCount = 34) " &_
					") " &_
				"AND CustomerCode IN(SELECT Num FROM Customer WHERE mName = '" & appCustName(0) & "') AND FieldCount = 37 "
Sql = Sql & ") GROUP BY YEAR(Field41) "
Set Rs = oConn.Execute(Sql)
Do Until Rs.EOF 
	rsYear = Rs.Fields(0)
	rsCount = Rs.Fields(1)

	k = rsYear-sDate '¹è¿­¹øÈ£
	If rsYear <= sDate Then
		arStatisApplName1(0) = arStatisApplName1(0) + rsCount
	Else
		arStatisApplName1(k) = arStatisApplName1(k) + rsCount
	End If
	sumStatisApplName1 = sumStatisApplName1 + rsCount

	Rs.MoveNext
Loop
Rs.Close
Set Rs = Nothing

'´ëÇ¥Ãâ¿øÀÎ2
Dim arStatisApplName2(9)
Dim sumStatisApplName2

Call sbArrayInitial(arStatisApplName2)
sumStatisApplName2 = 0
If appCustNumCnt = 2 Then

	Sql = "SELECT YEAR(Field41), COUNT(*) FROM LeftMenu0001 WHERE Field41 <= '" & nDate & "-12-31' "
	Sql = Sql & "AND (PatCode = 'A01' OR PatCode = 'A05') "
	Sql = Sql & "AND Num IN(	 "
	Sql = Sql & "SELECT TableNum FROM CustomerCode WHERE TableNum IN( " &_
							"SELECT TableNum FROM CustomerCode WHERE (CustomerCode = " & CustNum & " AND (FieldCount = 7 OR FieldCount = 34)) OR (CustomerCode IN(SELECT Num FROM Customer WHERE mName = '" & appCustName(0) & "') AND FieldCount = 37) " &_
						") AND CustomerCode IN(SELECT Num FROM Customer WHERE mName = '" & appCustName(1) & "') AND FieldCount = 37 "
	Sql = Sql & ") GROUP BY YEAR(Field41) "
	Set Rs = oConn.Execute(Sql)
	Do Until Rs.EOF 
		rsYear = Rs.Fields(0)
		rsCount = Rs.Fields(1)

		k = rsYear-sDate '¹è¿­¹øÈ£
		If rsYear <= sDate Then
			arStatisApplName2(0) = arStatisApplName2(0) + rsCount
		Else
			arStatisApplName2(k) = arStatisApplName2(k) + rsCount
		End If
		sumStatisApplName2 = sumStatisApplName2 + rsCount

		Rs.MoveNext
	Loop
	Rs.Close
	Set Rs = Nothing

End If

'±âÅ¸Ãâ¿øÀÎ
Dim arStatisApplEtc(9)
Dim sumStatisApplEtc

Call sbArrayInitial(arStatisApplEtc)
sumStatisApplEtc = 0

If appCustNumCnt = 2 Then

	Sql = "SELECT YEAR(Field41), COUNT(*) FROM LeftMenu0001 WHERE Field41 <= '" & nDate & "-12-31' "
	Sql = Sql & "AND (PatCode = 'A01' OR PatCode = 'A05') "
	Sql = Sql & "AND Num IN(	 " &_		
						"SELECT TableNum FROM CustomerCode WHERE CustomerCode = " & CustNum & " AND (FieldCount = 7 OR FieldCount = 34) " &_
						"AND TableNum NOT IN( " &_
							"SELECT TableNum FROM CustomerCode WHERE (FieldCount = 37 AND CustomerCode IN(SELECT Num FROM Customer WHERE mName = '" & appCustName(0) & "')) OR (FieldCount = 37 AND CustomerCode IN(SELECT Num FROM Customer WHERE mName = '" & appCustName(1) & "')) " &_
						") "
	Sql = Sql & ") GROUP BY YEAR(Field41) "
	Set Rs = oConn.Execute(Sql)
	Do Until Rs.EOF 
		rsYear = Rs.Fields(0)
		rsCount = Rs.Fields(1)

		k = rsYear-sDate '¹è¿­¹øÈ£
		If rsYear <= sDate Then
			arStatisApplEtc(0) = arStatisApplEtc(0) + rsCount
		Else
			arStatisApplEtc(k) = arStatisApplEtc(k) + rsCount
		End If
		sumStatisApplEtc = sumStatisApplEtc + rsCount

		Rs.MoveNext
	Loop
	Rs.Close
	Set Rs = Nothing

End If

'1. µî·Ï°Ç
Dim arStatisReg(9)
Dim sumStatisReg

Call sbArrayInitial(arStatisReg)
sumStatisReg = 0
Sql = "SELECT YEAR(Field41), COUNT(*) FROM LeftMenu0001 WHERE Field41 <= '" & nDate & "-12-31' AND Field73 IS NOT NULL "
Sql = Sql & "AND (PatCode = 'A01' OR PatCode = 'A05') "
Sql = Sql & "AND Num IN(	SELECT TableNum FROM CustomerCode WHERE CustomerCode = " & CustNum & " AND (FieldCount = 7 OR FieldCount = 34) ) GROUP BY YEAR(Field41) "
Set Rs = oConn.Execute(Sql)
Do Until Rs.EOF 
	rsYear = Rs.Fields(0)
	rsCount = Rs.Fields(1)

	k = rsYear-sDate '¹è¿­¹øÈ£
	If rsYear <= sDate Then
		arStatisReg(0) = arStatisReg(0) + rsCount
	Else
		arStatisReg(k) = arStatisReg(k) + rsCount
	End If
	sumStatisReg = sumStatisReg + rsCount

	Rs.MoveNext
Loop
Rs.Close
Set Rs = Nothing

'1-1. µî·Ï°Ç(ÇÑ¼º°ü¸®°Ç)
Dim Sql_1_1, DataCnt_1_1
Dim OurRef, YourRef, RegDate, RegNo, AnnEndDate, AnnYear
Sql_1_1 = "SELECT Field5, CASE WHEN Field221 IS NULL OR Field221 = '' THEN Field6 ELSE Field221 END, Field73, LEFT(ISNULL(Field74,''),10), MIN(EndDate), MIN(YearCount) " &_
				"FROM LeftMenu0001 L INNER JOIN YearlyPay Y ON L.Num = Y.LeftMenuTableNo " &_
				"WHERE Field41 <= '" & nDate & "-12-31' AND Field73 IS NOT NULL " &_
				"AND (PatCode = 'A01' OR PatCode = 'A05') " &_
				"AND L.Num IN(	SELECT TableNum FROM CustomerCode WHERE CustomerCode = " & CustNum & " AND (FieldCount = 7 OR FieldCount = 34) ) " &_
				"AND Y.State = '°ü¸®Áß' " &_
				"GROUP BY Field5, CASE WHEN Field221 IS NULL OR Field221 = '' THEN Field6 ELSE Field221 END, Field73, Field74 ORDER BY Field74 "
Set Rs = Server.CreateObject("ADODB.RECORDSET")
Rs.Open Sql_1_1, oConn, 1
DataCnt_1_1 = Rs.RecordCount
Rs.Close
Set Rs = Nothing

'1-2. µî·Ï°Ç(ÇÑ¼º¹Ì°ü¸®°Ç)
Dim Sql_1_2, DataCnt_1_2
Sql_1_2 = "SELECT LEFT(ISNULL(Field74,''),10) FROM LeftMenu0001 WHERE Field41 <= '" & nDate & "-12-31' AND Field73 IS NOT NULL " &_
				"AND (PatCode = 'A01' OR PatCode = 'A05') AND Field74 IS NOT NULL AND Field74 <> '' " &_
				"AND Num IN(	SELECT TableNum FROM CustomerCode WHERE CustomerCode = " & CustNum & " AND (FieldCount = 7 OR FieldCount = 34) ) " &_ 
				"AND Num NOT IN( " &_
					"SELECT DISTINCT L.Num " &_
					"FROM LeftMenu0001 L INNER JOIN YearlyPay Y ON L.Num = Y.LeftMenuTableNo " &_ 
					"WHERE Field41 <= '" & nDate & "-12-31' AND Field73 IS NOT NULL  " &_
					"AND (PatCode = 'A01' OR PatCode = 'A05') " &_ 
					"AND L.Num IN(	SELECT TableNum FROM CustomerCode WHERE CustomerCode = " & CustNum & " AND (FieldCount = 7 OR FieldCount = 34) ) " &_ 
					"AND Y.State = '°ü¸®Áß' " &_
				") ORDER BY Field74 "
Set Rs = Server.CreateObject("ADODB.RECORDSET")
Rs.Open Sql_1_2, oConn, 1
DataCnt_1_2 = Rs.RecordCount
Rs.Close
Set Rs = Nothing

'1-2. µî·Ï°Ç(ÇÑ¼º¹Ì°ü¸®°Ç)-µî·ÏÀÏ°ø¹é
Dim Sql_1_2_app, DataCnt_1_2_app
Sql_1_2_app = "SELECT ISNULL(Field42,'') FROM LeftMenu0001 WHERE Field41 <= '" & nDate & "-12-31' AND Field73 IS NOT NULL " &_
				"AND (PatCode = 'A01' OR PatCode = 'A05') AND (Field74 IS NULL OR Field74 = '') " &_
				"AND Num IN(	SELECT TableNum FROM CustomerCode WHERE CustomerCode = " & CustNum & " AND (FieldCount = 7 OR FieldCount = 34) ) " &_ 
				"AND Num NOT IN( " &_
					"SELECT DISTINCT L.Num " &_
					"FROM LeftMenu0001 L INNER JOIN YearlyPay Y ON L.Num = Y.LeftMenuTableNo " &_ 
					"WHERE Field41 <= '" & nDate & "-12-31' AND Field73 IS NOT NULL  " &_
					"AND (PatCode = 'A01' OR PatCode = 'A05') " &_ 
					"AND L.Num IN(	SELECT TableNum FROM CustomerCode WHERE CustomerCode = " & CustNum & " AND (FieldCount = 7 OR FieldCount = 34) ) " &_ 
					"AND Y.State = '°ü¸®Áß' " &_
				") ORDER BY Field74 "
Set Rs = Server.CreateObject("ADODB.RECORDSET")
Rs.Open Sql_1_2_app, oConn, 1
DataCnt_1_2_app = Rs.RecordCount
Rs.Close
Set Rs = Nothing

'2. Æ÷±â°Ç
Dim arStatisAban(9)
Dim sumStatisAban
Call sbArrayInitial(arStatisAban)
sumStatisAban = 0
Sql = "SELECT YEAR(Field41), COUNT(*) FROM LeftMenu0001 WHERE Field41 <= '" & nDate & "-12-31' AND Field73 IS NULL AND Field85 IS NOT NULL "
Sql = Sql & "AND (PatCode = 'A01' OR PatCode = 'A05') "
Sql = Sql & "AND Num IN(	SELECT TableNum FROM CustomerCode WHERE CustomerCode = " & CustNum & " AND (FieldCount = 7 OR FieldCount = 34) ) GROUP BY YEAR(Field41) "
Set Rs = oConn.Execute(Sql)
Do Until Rs.EOF 
	rsYear = Rs.Fields(0)
	rsCount = Rs.Fields(1)

	k = rsYear-sDate '¹è¿­¹øÈ£
	If rsYear <= sDate Then
		arStatisAban(0) = arStatisAban(0) + rsCount
	Else
		arStatisAban(k) = arStatisAban(k) + rsCount
	End If
	sumStatisAban = sumStatisAban + rsCount

	Rs.MoveNext
Loop
Rs.Close
Set Rs = Nothing

'2-1. Æ÷±â°Ç(ÃÖ±Ù2³â)
Dim Sql_2_1, DataCnt_2_1
Dim ApplNo, AbanDate, AbanDeDate, AbanMethod
Sql_2_1 = "SELECT Field5, CASE WHEN Field221 IS NULL OR Field221 = '' THEN Field6 ELSE Field221 END, RIGHT(ISNULL(Field42,''), 12), Field85, Field342, dbo.fnAbandonType_HS(Num) " &_
				"FROM LeftMenu0001 WHERE Field41 <= '" & nDate & "-12-31' AND Field73 IS NULL AND Field85 >= '" & nDate-1 & "-01-01' " &_
				"AND (PatCode = 'A01' OR PatCode = 'A05') " &_ 
				"AND Num IN(	SELECT TableNum FROM CustomerCode WHERE CustomerCode = " & CustNum & " AND (FieldCount = 7 OR FieldCount = 34) ) " &_
				"ORDER BY Field42 "
Set Rs = Server.CreateObject("ADODB.RECORDSET")
Rs.Open Sql_2_1, oConn, 1
DataCnt_2_1 = Rs.RecordCount
Rs.Close
Set Rs = Nothing

'2-1. Æ÷±â°Ç(2³â¹Ì¸¸)
Dim Sql_2_2, DataCnt_2_2
Sql_2_2 = "SELECT RIGHT(ISNULL(Field42,''), 12) " &_
				"FROM LeftMenu0001 WHERE Field41 <= '" & nDate & "-12-31' AND Field73 IS NULL AND Field85 <= '" & nDate-2 & "-12-31' " &_
				"AND (PatCode = 'A01' OR PatCode = 'A05') " &_ 
				"AND Num IN(	SELECT TableNum FROM CustomerCode WHERE CustomerCode = " & CustNum & " AND (FieldCount = 7 OR FieldCount = 34) ) " &_
				"ORDER BY Field42 "
Set Rs = Server.CreateObject("ADODB.RECORDSET")
Rs.Open Sql_2_2, oConn, 1
DataCnt_2_2 = Rs.RecordCount
Rs.Close
Set Rs = Nothing

'3. °è·ù°Ç
Dim arStatisPen(9)
Dim sumStatisPen
Call sbArrayInitial(arStatisPen)
sumStatisPen = 0
Sql = "SELECT YEAR(Field41), COUNT(*) FROM LeftMenu0001 WHERE Field41 <= '" & nDate & "-12-31' AND Field73 IS NULL AND Field85 IS NULL "
Sql = Sql & "AND (PatCode = 'A01' OR PatCode = 'A05') "
Sql = Sql & "AND Num IN(	SELECT TableNum FROM CustomerCode WHERE CustomerCode = " & CustNum & " AND (FieldCount = 7 OR FieldCount = 34) ) GROUP BY YEAR(Field41) "
Set Rs = oConn.Execute(Sql)
Do Until Rs.EOF 
	rsYear = Rs.Fields(0)
	rsCount = Rs.Fields(1)

	k = rsYear-sDate '¹è¿­¹øÈ£
	If rsYear <= sDate Then
		arStatisPen(0) = arStatisPen(0) + rsCount
	Else
		arStatisPen(k) = arStatisPen(k) + rsCount
	End If
	sumStatisPen = sumStatisPen + rsCount

	Rs.MoveNext
Loop
Rs.Close
Set Rs = Nothing

'3. °è·ù°Ç(¸®½ºÆ®)
Dim Sql_3, DataCnt_3
Dim FilingDate, PresentStatus
Sql_3 = "SELECT CASE WHEN Field221 IS NULL OR Field221 = '' THEN Field6 ELSE Field221 END, Field5, Field42, Field128, dbo.fnPendingType_HS(Num) " &_
			"FROM LeftMenu0001 WHERE Field41 <= '" & nDate & "-12-31' AND Field73 IS NULL AND Field85 IS NULL " &_
			"AND (PatCode = 'A01' OR PatCode = 'A05') " &_
			"AND Num IN(	SELECT TableNum FROM CustomerCode WHERE CustomerCode = " & CustNum & " AND (FieldCount = 7 OR FieldCount = 34) ) " &_
			"ORDER BY Field42 "
Set Rs = Server.CreateObject("ADODB.RECORDSET")
Rs.Open Sql_3, oConn, 1
DataCnt_3 = Rs.RecordCount
Rs.Close
Set Rs = Nothing

'3-a. ½É»ç¹ÌÃ»±¸
Dim arStatisNoExam(9)
Dim sumStatisNoExam
Call sbArrayInitial(arStatisNoExam)
sumStatisNoExam = 0
Sql = "SELECT YEAR(Field41), COUNT(*) FROM LeftMenu0001 WHERE Field41 <= '" & nDate & "-12-31' AND Field73 IS NULL AND Field85 IS NULL "
Sql = Sql & "AND Field48 IS NULL "
Sql = Sql & "AND (PatCode = 'A01' OR PatCode = 'A05') "
Sql = Sql & "AND Num IN(	SELECT TableNum FROM CustomerCode WHERE CustomerCode = " & CustNum & " AND (FieldCount = 7 OR FieldCount = 34) ) GROUP BY YEAR(Field41) "
Set Rs = oConn.Execute(Sql)
Do Until Rs.EOF 
	rsYear = Rs.Fields(0)
	rsCount = Rs.Fields(1)

	k = rsYear-sDate '¹è¿­¹øÈ£
	If rsYear <= sDate Then
		arStatisNoExam(0) = arStatisNoExam(0) + rsCount
	Else
		arStatisNoExam(k) = arStatisNoExam(k) + rsCount
	End If
	sumStatisNoExam = sumStatisNoExam + rsCount

	Rs.MoveNext
Loop
Rs.Close
Set Rs = Nothing

'3-b. ½É»çÃ»±¸
Dim arStatisExam(9)
Dim sumStatisExam
Call sbArrayInitial(arStatisExam)
sumStatisExam = 0
Sql = "SELECT YEAR(Field41), COUNT(*) FROM LeftMenu0001 WHERE Field41 <= '" & nDate & "-12-31' AND Field73 IS NULL AND Field85 IS NULL "
Sql = Sql & "AND Field48 IS NOT NULL "
Sql = Sql & "AND (PatCode = 'A01' OR PatCode = 'A05') "
Sql = Sql & "AND Num IN(	SELECT TableNum FROM CustomerCode WHERE CustomerCode = " & CustNum & " AND (FieldCount = 7 OR FieldCount = 34) ) GROUP BY YEAR(Field41) "
Set Rs = oConn.Execute(Sql)
Do Until Rs.EOF 
	rsYear = Rs.Fields(0)
	rsCount = Rs.Fields(1)

	k = rsYear-sDate '¹è¿­¹øÈ£
	If rsYear <= sDate Then
		arStatisExam(0) = arStatisExam(0) + rsCount
	Else
		arStatisExam(k) = arStatisExam(k) + rsCount
	End If
	sumStatisExam = sumStatisExam + rsCount

	Rs.MoveNext
Loop
Rs.Close
Set Rs = Nothing

'3-b-1. ½É»çÁß
Dim arStatisWaitExam(9)
Dim sumStatisWaitExam
Call sbArrayInitial(arStatisWaitExam)
sumStatisWaitExam = 0
Sql = "SELECT YEAR(Field41), COUNT(*) FROM LeftMenu0001 WHERE Field41 <= '" & nDate & "-12-31' AND Field73 IS NULL AND Field85 IS NULL "
Sql = Sql & "AND Field48 IS NOT NULL "
Sql = Sql & "AND Field71 IS NULL AND dbo.fnOfficeActionCnt(Num) = 0 AND dbo.fnAppealType(Field42, Num) = '' "
Sql = Sql & "AND (PatCode = 'A01' OR PatCode = 'A05') "
Sql = Sql & "AND Num IN(	SELECT TableNum FROM CustomerCode WHERE CustomerCode = " & CustNum & " AND (FieldCount = 7 OR FieldCount = 34) ) GROUP BY YEAR(Field41) "
Set Rs = oConn.Execute(Sql)
Do Until Rs.EOF 
	rsYear = Rs.Fields(0)
	rsCount = Rs.Fields(1)

	k = rsYear-sDate '¹è¿­¹øÈ£
	If rsYear <= sDate Then
		arStatisWaitExam(0) = arStatisWaitExam(0) + rsCount
	Else
		arStatisWaitExam(k) = arStatisWaitExam(k) + rsCount
	End If
	sumStatisWaitExam = sumStatisWaitExam + rsCount

	Rs.MoveNext
Loop
Rs.Close
Set Rs = Nothing

'3-b-2. ½É»ç¿Ï·á
Dim arStatisUnderExam(9)
Dim sumStatisUnderExam
Call sbArrayInitial(arStatisUnderExam)
sumStatisUnderExam = 0
Sql = "SELECT YEAR(Field41), COUNT(*) FROM LeftMenu0001 WHERE Field41 <= '" & nDate & "-12-31' AND Field73 IS NULL AND Field85 IS NULL "
Sql = Sql & "AND Field48 IS NOT NULL "
Sql = Sql & "AND (Field71 IS NOT NULL OR dbo.fnOfficeActionCnt(Num) > 0 OR dbo.fnAppealType(Field42, Num) <> '') "
Sql = Sql & "AND (PatCode = 'A01' OR PatCode = 'A05') "
Sql = Sql & "AND Num IN(	SELECT TableNum FROM CustomerCode WHERE CustomerCode = " & CustNum & " AND (FieldCount = 7 OR FieldCount = 34) ) GROUP BY YEAR(Field41) "
'OA, ½ÉÆÇÁøÇà ¿©ºÎ È®ÀÎ
Set Rs = oConn.Execute(Sql)
Do Until Rs.EOF 
	rsYear = Rs.Fields(0)
	rsCount = Rs.Fields(1)

	k = rsYear-sDate '¹è¿­¹øÈ£
	If rsYear <= sDate Then
		arStatisUnderExam(0) = arStatisUnderExam(0) + rsCount
	Else
		arStatisUnderExam(k) = arStatisUnderExam(k) + rsCount
	End If
	sumStatisUnderExam = sumStatisUnderExam + rsCount

	Rs.MoveNext
Loop
Rs.Close
Set Rs = Nothing

'3-b-2- a. µî·Ï°áÁ¤
Dim arStatisNoticeExam(9)
Dim sumStatisNoticeExam
Call sbArrayInitial(arStatisNoticeExam)
sumStatisNoticeExam = 0
Sql = "SELECT YEAR(Field41), COUNT(*) FROM LeftMenu0001 WHERE Field41 <= '" & nDate & "-12-31' AND Field73 IS NULL AND Field85 IS NULL "
Sql = Sql & "AND Field48 IS NOT NULL "
Sql = Sql & "AND Field71 IS NOT NULL "
Sql = Sql & "AND (PatCode = 'A01' OR PatCode = 'A05') "
Sql = Sql & "AND Num IN(	SELECT TableNum FROM CustomerCode WHERE CustomerCode = " & CustNum & " AND (FieldCount = 7 OR FieldCount = 34) ) GROUP BY YEAR(Field41) "
'OA, ½ÉÆÇÁøÇà ¿©ºÎ È®ÀÎ
Set Rs = oConn.Execute(Sql)
Do Until Rs.EOF 
	rsYear = Rs.Fields(0)
	rsCount = Rs.Fields(1)

	k = rsYear-sDate '¹è¿­¹øÈ£
	If rsYear <= sDate Then
		arStatisNoticeExam(0) = arStatisNoticeExam(0) + rsCount
	Else
		arStatisNoticeExam(k) = arStatisNoticeExam(k) + rsCount
	End If
	sumStatisNoticeExam = sumStatisNoticeExam + rsCount

	Rs.MoveNext
Loop
Rs.Close
Set Rs = Nothing

'3-b-2- b. OAÁøÇà
Dim arStatisOAExam(9)
Dim sumStatisOAExam
Call sbArrayInitial(arStatisOAExam)
sumStatisOAExam = 0
Sql = "SELECT YEAR(Field41), COUNT(*) FROM LeftMenu0001 WHERE Field41 <= '" & nDate & "-12-31' AND Field73 IS NULL AND Field85 IS NULL "
Sql = Sql & "AND Field48 IS NOT NULL AND Field71 IS NULL "
Sql = Sql & "AND dbo.fnOfficeActionCnt(Num) > 0 AND dbo.fnAppealType(Field42, Num) = '' "
Sql = Sql & "AND (PatCode = 'A01' OR PatCode = 'A05') "
Sql = Sql & "AND Num IN(	SELECT TableNum FROM CustomerCode WHERE CustomerCode = " & CustNum & " AND (FieldCount = 7 OR FieldCount = 34) ) GROUP BY YEAR(Field41) "
'OA, ½ÉÆÇÁøÇà ¿©ºÎ È®ÀÎ
Set Rs = oConn.Execute(Sql)
Do Until Rs.EOF 
	rsYear = Rs.Fields(0)
	rsCount = Rs.Fields(1)

	k = rsYear-sDate '¹è¿­¹øÈ£
	If rsYear <= sDate Then
		arStatisOAExam(0) = arStatisOAExam(0) + rsCount
	Else
		arStatisOAExam(k) = arStatisOAExam(k) + rsCount
	End If
	sumStatisOAExam = sumStatisOAExam + rsCount

	Rs.MoveNext
Loop
Rs.Close
Set Rs = Nothing

'3-b-2- c. Æ¯Çã½ÉÆÇ¿ø
Dim arStatisIPTExam(9)
Dim sumStatisIPTExam
Call sbArrayInitial(arStatisIPTExam)
sumStatisIPTExam = 0
Sql = "SELECT YEAR(Field41), COUNT(*) FROM LeftMenu0001 WHERE Field41 <= '" & nDate & "-12-31' AND Field73 IS NULL AND Field85 IS NULL "
Sql = Sql & "AND Field48 IS NOT NULL "
Sql = Sql & "AND dbo.fnAppealType(Field42, Num) = 'A' AND Field71 IS NULL "
Sql = Sql & "AND (PatCode = 'A01' OR PatCode = 'A05') "
Sql = Sql & "AND Num IN(	SELECT TableNum FROM CustomerCode WHERE CustomerCode = " & CustNum & " AND (FieldCount = 7 OR FieldCount = 34) ) GROUP BY YEAR(Field41) "
'OA, ½ÉÆÇÁøÇà ¿©ºÎ È®ÀÎ
Set Rs = oConn.Execute(Sql)
Do Until Rs.EOF 
	rsYear = Rs.Fields(0)
	rsCount = Rs.Fields(1)

	k = rsYear-sDate '¹è¿­¹øÈ£
	If rsYear <= sDate Then
		arStatisIPTExam(0) = arStatisIPTExam(0) + rsCount
	Else
		arStatisIPTExam(k) = arStatisIPTExam(k) + rsCount
	End If
	sumStatisIPTExam = sumStatisIPTExam + rsCount

	Rs.MoveNext
Loop
Rs.Close
Set Rs = Nothing

'3-b-2- d. Æ¯Çã¹ý¿ø
Dim arStatisPCExam(9)
Dim sumStatisPCExam
Call sbArrayInitial(arStatisPCExam)
sumStatisPCExam = 0
Sql = "SELECT YEAR(Field41), COUNT(*) FROM LeftMenu0001 WHERE Field41 <= '" & nDate & "-12-31' AND Field73 IS NULL AND Field85 IS NULL "
Sql = Sql & "AND Field48 IS NOT NULL "
Sql = Sql & "AND dbo.fnAppealType(Field42, Num) = 'B' AND Field71 IS NULL "
Sql = Sql & "AND (PatCode = 'A01' OR PatCode = 'A05') "
Sql = Sql & "AND Num IN(	SELECT TableNum FROM CustomerCode WHERE CustomerCode = " & CustNum & " AND (FieldCount = 7 OR FieldCount = 34) ) GROUP BY YEAR(Field41) "
'OA, ½ÉÆÇÁøÇà ¿©ºÎ È®ÀÎ
Set Rs = oConn.Execute(Sql)
Do Until Rs.EOF 
	rsYear = Rs.Fields(0)
	rsCount = Rs.Fields(1)

	k = rsYear-sDate '¹è¿­¹øÈ£
	If rsYear <= sDate Then
		arStatisPCExam(0) = arStatisPCExam(0) + rsCount
	Else
		arStatisPCExam(k) = arStatisPCExam(k) + rsCount
	End If
	sumStatisPCExam = sumStatisPCExam + rsCount

	Rs.MoveNext
Loop
Rs.Close
Set Rs = Nothing

'3-b-2- d. ´ë¹ý¿ø
Dim arStatisSCExam(9)
Dim sumStatisSCExam
Call sbArrayInitial(arStatisSCExam)
sumStatisSCExam = 0
Sql = "SELECT YEAR(Field41), COUNT(*) FROM LeftMenu0001 WHERE Field41 <= '" & nDate & "-12-31' AND Field73 IS NULL AND Field85 IS NULL "
Sql = Sql & "AND Field48 IS NOT NULL "
Sql = Sql & "AND dbo.fnAppealType(Field42, Num) = 'C' AND Field71 IS NULL "
Sql = Sql & "AND (PatCode = 'A01' OR PatCode = 'A05') "
Sql = Sql & "AND Num IN(	SELECT TableNum FROM CustomerCode WHERE CustomerCode = " & CustNum & " AND (FieldCount = 7 OR FieldCount = 34) ) GROUP BY YEAR(Field41) "
'OA, ½ÉÆÇÁøÇà ¿©ºÎ È®ÀÎ
Set Rs = oConn.Execute(Sql)
Do Until Rs.EOF 
	rsYear = Rs.Fields(0)
	rsCount = Rs.Fields(1)

	k = rsYear-sDate '¹è¿­¹øÈ£
	If rsYear <= sDate Then
		arStatisSCExam(0) = arStatisSCExam(0) + rsCount
	Else
		arStatisSCExam(k) = arStatisSCExam(k) + rsCount
	End If
	sumStatisSCExam = sumStatisSCExam + rsCount

	Rs.MoveNext
Loop
Rs.Close
Set Rs = Nothing

%>
<html xmlns:v="urn:schemas-microsoft-com:vml"
xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:w="urn:schemas-microsoft-com:office:word"
xmlns:m="http://schemas.microsoft.com/office/2004/12/omml"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=Content-Type content="text/html; charset=ks_c_5601-1987">
<meta name=ProgId content=Word.Document>
<meta name=Generator content="Microsoft Word 15">
<meta name=Originator content="Microsoft Word 15">

<!--[if !mso]>
<style>
v\:* {behavior:url(#default#VML);}
o\:* {behavior:url(#default#VML);}
w\:* {behavior:url(#default#VML);}
.shape {behavior:url(#default#VML);}
</style>
<![endif]--><!--[if gte mso 9]>
	<xml>
    <w:WordDocument>
			<w:View>Print</w:View>
			<w:Zoom>90</w:Zoom>
			<w:DoNotOptimizeForBrowser/>
		</w:WordDocument>
	</xml>
<![endif]-->

<style>
<!--
 /* Font Definitions */
 @font-face
	{font-family:Wingdings;
	panose-1:5 0 0 0 0 0 0 0 0 0;
	mso-font-charset:2;
	mso-generic-font-family:auto;
	mso-font-pitch:variable;
	mso-font-signature:0 268435456 0 0 -2147483648 0;}
@font-face
	{font-family:¹ÙÅÁ;
	panose-1:2 3 6 0 0 1 1 1 1 1;
	mso-font-alt:Batang;
	mso-font-charset:129;
	mso-generic-font-family:roman;
	mso-font-pitch:variable;
	mso-font-signature:-1342176593 1775729915 48 0 524447 0;}
@font-face
	{font-family:"Cambria Math";
	panose-1:2 4 5 3 5 4 6 3 2 4;
	mso-font-charset:0;
	mso-generic-font-family:roman;
	mso-font-pitch:variable;
	mso-font-signature:-536870145 1107305727 0 0 415 0;}
@font-face
	{font-family:"¸¼Àº °íµñ";
	panose-1:2 11 5 3 2 0 0 2 0 4;
	mso-font-charset:129;
	mso-generic-font-family:modern;
	mso-font-pitch:variable;
	mso-font-signature:-1879047505 165117179 18 0 524289 0;}
@font-face
	{font-family:ÇÑÄÄ¹ÙÅÁ;
	mso-font-alt:"Arial Unicode MS";
	mso-font-charset:129;
	mso-generic-font-family:roman;
	mso-font-pitch:variable;
	mso-font-signature:0 -69206017 16777215 0 -2143354369 0;}
@font-face
	{font-family:"Microsoft YaHei";
	panose-1:2 11 5 3 2 2 4 2 2 4;
	mso-font-charset:134;
	mso-generic-font-family:swiss;
	mso-font-pitch:variable;
	mso-font-signature:-2147483001 672087122 22 0 262175 0;}
@font-face
	{font-family:Times-Bold;
	panose-1:0 0 0 0 0 0 0 0 0 0;
	mso-font-alt:"Times New Roman";
	mso-font-charset:0;
	mso-generic-font-family:roman;
	mso-font-format:other;
	mso-font-pitch:auto;
	mso-font-signature:3 0 0 0 1 0;}
@font-face
	{font-family:"\@¹ÙÅÁ";
	panose-1:2 3 6 0 0 1 1 1 1 1;
	mso-font-charset:129;
	mso-generic-font-family:roman;
	mso-font-pitch:variable;
	mso-font-signature:-1342176593 1775729915 48 0 524447 0;}
@font-face
	{font-family:"\@¸¼Àº °íµñ";
	panose-1:2 11 5 3 2 0 0 2 0 4;
	mso-font-charset:129;
	mso-generic-font-family:modern;
	mso-font-pitch:variable;
	mso-font-signature:-1879047505 165117179 18 0 524289 0;}
@font-face
	{font-family:"\@Microsoft YaHei";
	panose-1:2 11 5 3 2 2 4 2 2 4;
	mso-font-charset:134;
	mso-generic-font-family:swiss;
	mso-font-pitch:variable;
	mso-font-signature:-2147483001 672087122 22 0 262175 0;}
@font-face
	{font-family:"\@ÇÑÄÄ¹ÙÅÁ";
	mso-font-charset:129;
	mso-generic-font-family:roman;
	mso-font-pitch:variable;
	mso-font-signature:0 -69206017 16777215 0 -2143354369 0;}
 /* Style Definitions */
 p.MsoNormal, li.MsoNormal, div.MsoNormal
	{mso-style-unhide:no;
	mso-style-qformat:yes;
	mso-style-parent:"";
	margin-top:0cm;
	margin-right:0cm;
	margin-bottom:8.0pt;
	margin-left:0cm;
	text-align:justify;
	text-justify:inter-ideograph;
	line-height:107%;
	mso-pagination:none;
	text-autospace:none;
	word-break:break-hangul;
	font-size:10.0pt;
	mso-bidi-font-size:11.0pt;
	font-family:"¸¼Àº °íµñ";
	mso-bidi-font-family:"Times New Roman";
	mso-font-kerning:1.0pt;}
p.MsoHeader, li.MsoHeader, div.MsoHeader
	{mso-style-priority:99;
	mso-style-link:"¸Ó¸®±Û Char";
	margin-top:0cm;
	margin-right:0cm;
	margin-bottom:8.0pt;
	margin-left:0cm;
	text-align:justify;
	text-justify:inter-ideograph;
	line-height:107%;
	mso-pagination:none;
	tab-stops:center 225.65pt right 451.3pt;
	layout-grid-mode:char;
	text-autospace:none;
	word-break:break-hangul;
	font-size:10.0pt;
	mso-bidi-font-size:11.0pt;
	font-family:"¸¼Àº °íµñ";
	mso-bidi-font-family:"Times New Roman";
	mso-font-kerning:1.0pt;}
p.MsoFooter, li.MsoFooter, div.MsoFooter
	{mso-style-priority:99;
	mso-style-link:"¹Ù´Ú±Û Char";
	margin-top:0cm;
	margin-right:0cm;
	margin-bottom:8.0pt;
	margin-left:0cm;
	text-align:justify;
	text-justify:inter-ideograph;
	line-height:107%;
	mso-pagination:none;
	tab-stops:center 225.65pt right 451.3pt;
	layout-grid-mode:char;
	text-autospace:none;
	word-break:break-hangul;
	font-size:10.0pt;
	mso-bidi-font-size:11.0pt;
	font-family:"¸¼Àº °íµñ";
	mso-bidi-font-family:"Times New Roman";
	mso-font-kerning:1.0pt;}
span.Char
	{mso-style-name:"¸Ó¸®±Û Char";
	mso-style-priority:99;
	mso-style-unhide:no;
	mso-style-locked:yes;
	mso-style-link:¸Ó¸®±Û;}
span.Char0
	{mso-style-name:"¹Ù´Ú±Û Char";
	mso-style-priority:99;
	mso-style-unhide:no;
	mso-style-locked:yes;
	mso-style-link:¹Ù´Ú±Û;}
span.GramE
	{mso-style-name:"";
	mso-gram-e:yes;}
.MsoChpDefault
	{mso-style-type:export-only;
	mso-default-props:yes;
	font-family:"¸¼Àº °íµñ";
	mso-ascii-font-family:"¸¼Àº °íµñ";
	mso-fareast-font-family:"¸¼Àº °íµñ";
	mso-hansi-font-family:"¸¼Àº °íµñ";}
 /* Page Definitions */
 @page
	{mso-page-border-surround-header:no;
	mso-page-border-surround-footer:no;
	mso-gutter-position:top;
	mso-footnote-separator:url("") fs;
	mso-footnote-continuation-separator:url("") fcs;
	mso-endnote-separator:url("") es;
	mso-endnote-continuation-separator:url("") ecs;}
@page WordSection1
	{size:595.3pt 841.9pt;
	margin:70.9pt 42.55pt 2.0cm 42.55pt;
	mso-header-margin:42.55pt;
	mso-footer-margin:49.6pt;
	mso-footer:url("") f1;
	mso-paper-source:0;}
div.WordSection1
	{page:WordSection1;}
 /* List Definitions */
 @list l0
	{mso-list-id:1280986654;
	mso-list-type:hybrid;
	mso-list-template-ids:-2107184744 1052524888 67698691 67698693 67698689 67698691 67698693 67698689 67698691 67698693;}
@list l0:level1
	{mso-level-start-at:0;
	mso-level-number-format:bullet;
	mso-level-text:¡Ø;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:38.0pt;
	text-indent:-18.0pt;
	font-family:"¸¼Àº °íµñ";
	mso-bidi-font-family:"Times New Roman";}
@list l0:level2
	{mso-level-number-format:bullet;
	mso-level-text:\F06E;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:60.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l0:level3
	{mso-level-number-format:bullet;
	mso-level-text:\F075;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:80.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l0:level4
	{mso-level-number-format:bullet;
	mso-level-text:\F06C;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:100.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l0:level5
	{mso-level-number-format:bullet;
	mso-level-text:\F06E;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:120.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l0:level6
	{mso-level-number-format:bullet;
	mso-level-text:\F075;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:140.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l0:level7
	{mso-level-number-format:bullet;
	mso-level-text:\F06C;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:160.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l0:level8
	{mso-level-number-format:bullet;
	mso-level-text:\F06E;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:180.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l0:level9
	{mso-level-number-format:bullet;
	mso-level-text:\F075;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:200.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l1
	{mso-list-id:1464889382;
	mso-list-type:hybrid;
	mso-list-template-ids:1339734894 -1668618764 67698691 67698693 67698689 67698691 67698693 67698689 67698691 67698693;}
@list l1:level1
	{mso-level-start-at:0;
	mso-level-number-format:bullet;
	mso-level-text:¡Ø;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:38.0pt;
	text-indent:-18.0pt;
	font-family:"¸¼Àº °íµñ";
	mso-bidi-font-family:"Times New Roman";}
@list l1:level2
	{mso-level-number-format:bullet;
	mso-level-text:\F06E;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:60.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l1:level3
	{mso-level-number-format:bullet;
	mso-level-text:\F075;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:80.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l1:level4
	{mso-level-number-format:bullet;
	mso-level-text:\F06C;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:100.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l1:level5
	{mso-level-number-format:bullet;
	mso-level-text:\F06E;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:120.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l1:level6
	{mso-level-number-format:bullet;
	mso-level-text:\F075;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:140.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l1:level7
	{mso-level-number-format:bullet;
	mso-level-text:\F06C;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:160.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l1:level8
	{mso-level-number-format:bullet;
	mso-level-text:\F06E;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:180.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l1:level9
	{mso-level-number-format:bullet;
	mso-level-text:\F075;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:200.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l2
	{mso-list-id:1727533308;
	mso-list-type:hybrid;
	mso-list-template-ids:1322166890 1052524888 67698691 67698693 67698689 67698691 67698693 67698689 67698691 67698693;}
@list l2:level1
	{mso-level-start-at:6201;
	mso-level-number-format:bullet;
	mso-level-text:¡Ø;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:38.0pt;
	text-indent:-18.0pt;
	font-family:"¸¼Àº °íµñ";
	mso-bidi-font-family:"Times New Roman";}
@list l2:level2
	{mso-level-number-format:bullet;
	mso-level-text:\F06E;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:60.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l2:level3
	{mso-level-number-format:bullet;
	mso-level-text:\F075;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:80.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l2:level4
	{mso-level-number-format:bullet;
	mso-level-text:\F06C;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:100.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l2:level5
	{mso-level-number-format:bullet;
	mso-level-text:\F06E;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:120.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l2:level6
	{mso-level-number-format:bullet;
	mso-level-text:\F075;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:140.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l2:level7
	{mso-level-number-format:bullet;
	mso-level-text:\F06C;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:160.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l2:level8
	{mso-level-number-format:bullet;
	mso-level-text:\F06E;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:180.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l2:level9
	{mso-level-number-format:bullet;
	mso-level-text:\F075;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:200.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l3
	{mso-list-id:1772973136;
	mso-list-type:hybrid;
	mso-list-template-ids:181571736 67698691 67698691 67698693 67698689 67698691 67698693 67698689 67698691 67698693;}
@list l3:level1
	{mso-level-number-format:bullet;
	mso-level-text:\F06E;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:40.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l3:level2
	{mso-level-number-format:bullet;
	mso-level-text:\F06E;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:60.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l3:level3
	{mso-level-number-format:bullet;
	mso-level-text:\F075;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:80.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l3:level4
	{mso-level-number-format:bullet;
	mso-level-text:\F06C;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:100.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l3:level5
	{mso-level-number-format:bullet;
	mso-level-text:\F06E;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:120.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l3:level6
	{mso-level-number-format:bullet;
	mso-level-text:\F075;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:140.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l3:level7
	{mso-level-number-format:bullet;
	mso-level-text:\F06C;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:160.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l3:level8
	{mso-level-number-format:bullet;
	mso-level-text:\F06E;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:180.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l3:level9
	{mso-level-number-format:bullet;
	mso-level-text:\F075;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:200.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l4
	{mso-list-id:2099906583;
	mso-list-type:hybrid;
	mso-list-template-ids:-1123756488 -1191037282 67698691 67698693 67698689 67698691 67698693 67698689 67698691 67698693;}
@list l4:level1
	{mso-level-start-at:0;
	mso-level-number-format:bullet;
	mso-level-text:¡Ø;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:38.0pt;
	text-indent:-18.0pt;
	font-family:"¸¼Àº °íµñ";
	mso-bidi-font-family:"Times New Roman";}
@list l4:level2
	{mso-level-number-format:bullet;
	mso-level-text:\F06E;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:60.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l4:level3
	{mso-level-number-format:bullet;
	mso-level-text:\F075;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:80.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l4:level4
	{mso-level-number-format:bullet;
	mso-level-text:\F06C;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:100.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l4:level5
	{mso-level-number-format:bullet;
	mso-level-text:\F06E;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:120.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l4:level6
	{mso-level-number-format:bullet;
	mso-level-text:\F075;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:140.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l4:level7
	{mso-level-number-format:bullet;
	mso-level-text:\F06C;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:160.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l4:level8
	{mso-level-number-format:bullet;
	mso-level-text:\F06E;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:180.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
@list l4:level9
	{mso-level-number-format:bullet;
	mso-level-text:\F075;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	margin-left:200.0pt;
	text-indent:-20.0pt;
	font-family:Wingdings;}
ol
	{margin-bottom:0cm;}
ul
	{margin-bottom:0cm;}
-->
</style>
<!--[if gte mso 10]>
<style>
 /* Style Definitions */
 table.MsoNormalTable
	{mso-style-name:"Ç¥ÁØ Ç¥";
	mso-tstyle-rowband-size:0;
	mso-tstyle-colband-size:0;
	mso-style-noshow:yes;
	mso-style-priority:99;
	mso-style-parent:"";
	mso-padding-alt:0cm 5.4pt 0cm 5.4pt;
	mso-para-margin:0cm;
	mso-para-margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:10.0pt;
	font-family:"¸¼Àº °íµñ";}
table.MsoTableGrid
	{mso-style-name:"Ç¥ ±¸ºÐ¼±";
	mso-tstyle-rowband-size:0;
	mso-tstyle-colband-size:0;
	mso-style-priority:39;
	mso-style-unhide:no;
	border:solid windowtext 1.0pt;
	mso-border-alt:solid windowtext .5pt;
	mso-padding-alt:0cm 5.4pt 0cm 5.4pt;
	mso-border-insideh:.5pt solid windowtext;
	mso-border-insidev:.5pt solid windowtext;
	mso-para-margin:0cm;
	mso-para-margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:10.0pt;
	font-family:"¸¼Àº °íµñ";}
</style>
<![endif]--><!--[if gte mso 9]><xml>
 <o:shapedefaults v:ext="edit" spidmax="2049"/>
</xml><![endif]--><!--[if gte mso 9]><xml>
 <o:shapelayout v:ext="edit">
  <o:idmap v:ext="edit" data="1"/>
 </o:shapelayout></xml><![endif]-->
</head>

<body lang=KO style='tab-interval:40.0pt'>

<div class=WordSection1>

<p class=MsoNormal align=center style='text-align:center;mso-line-height-alt:
0pt;mso-pagination:widow-orphan'><span lang=EN-US style='font-size:12.0pt;
mso-bidi-font-size:11.0pt;font-family:"Times New Roman",serif;mso-fareast-font-family:
ÇÑÄÄ¹ÙÅÁ;mso-hansi-font-family:"¸¼Àº °íµñ";color:black;mso-no-proof:yes'><!--[if gte vml 1]><v:shapetype
 id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t"
 path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">
 <v:stroke joinstyle="miter"/>
 <v:formulas>
  <v:f eqn="if lineDrawn pixelLineWidth 0"/>
  <v:f eqn="sum @0 1 0"/>
  <v:f eqn="sum 0 0 @1"/>
  <v:f eqn="prod @2 1 2"/>
  <v:f eqn="prod @3 21600 pixelWidth"/>
  <v:f eqn="prod @3 21600 pixelHeight"/>
  <v:f eqn="sum @0 0 1"/>
  <v:f eqn="prod @6 1 2"/>
  <v:f eqn="prod @7 21600 pixelWidth"/>
  <v:f eqn="sum @8 21600 0"/>
  <v:f eqn="prod @7 21600 pixelHeight"/>
  <v:f eqn="sum @10 21600 0"/>
 </v:formulas>
 <v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>
 <o:lock v:ext="edit" aspectratio="t"/>
</v:shapetype><v:shape id="±×¸²_x0020_1" o:spid="_x0000_i1026" type="#_x0000_t75"
 alt="¿µ¹®Çìµå-2015" style='width:506.25pt;height:60.75pt;visibility:visible;
 mso-wrap-style:square'>
 <v:imagedata src="http://<%=Request.ServerVariables("SERVER_NAME") & ":" & Request.ServerVariables("SERVER_PORT")%>/images_letter/RPABE0084E.jpg" o:title="¿µ¹®Çìµå-2015"/>
</v:shape><![endif]--><![if !vml]><img width=675 height=81
src="http://<%=Request.ServerVariables("SERVER_NAME") & ":" & Request.ServerVariables("SERVER_PORT")%>/images_letter/RPABE0084E.jpg" alt=¿µ¹®Çìµå-2015 v:shapes="±×¸²_x0020_1"><![endif]></span><u><span
lang=NL style='font-size:12.0pt;mso-bidi-font-size:11.0pt;font-family:"Times New Roman",serif;
mso-fareast-font-family:ÇÑÄÄ¹ÙÅÁ;mso-hansi-font-family:"¸¼Àº °íµñ";color:black;
mso-ansi-language:NL'><o:p></o:p></span></u></p>

<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
12.0pt;mso-line-height-rule:exactly;mso-pagination:widow-orphan'><u><span
lang=EN-US style='font-size:12.0pt;mso-bidi-font-size:11.0pt;font-family:"Times New Roman",serif;
mso-fareast-font-family:"Microsoft YaHei";mso-hansi-font-family:"¸¼Àº °íµñ";
color:black'><o:p><span style='text-decoration:none'>&nbsp;</span></o:p></span></u></p>

<table class=MsoTableGrid border=0 cellspacing=0 cellpadding=0
 style='border-collapse:collapse;border:none;mso-yfti-tbllook:1184;mso-padding-alt:
 0cm 0cm 0cm 0cm;mso-border-insidev:none'>
 <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;mso-yfti-lastrow:yes'>
  <td width=340 valign=top style='width:254.85pt;padding:0cm 0cm 0cm 0cm'>
  <p class=MsoNormal align=left style='margin-bottom:0cm;margin-bottom:.0001pt;
  text-align:left;line-height:14.0pt;mso-line-height-rule:exactly;mso-pagination:
  widow-orphan'><u><span lang=EN-US style='font-size:12.0pt;mso-bidi-font-size:
  11.0pt;font-family:"Times New Roman",serif;mso-fareast-font-family:ÇÑÄÄ¹ÙÅÁ;
  mso-hansi-font-family:"¸¼Àº °íµñ";color:black'>Via E-mail<o:p></o:p></span></u></p>
  </td>
  <td width=340 valign=top style='width:254.85pt;padding:0cm 0cm 0cm 0cm'>
  <p class=MsoNormal align=right style='margin-bottom:0cm;margin-bottom:.0001pt;
  text-align:right;line-height:14.0pt;mso-line-height-rule:exactly;mso-pagination:
  widow-orphan'><span lang=EN-US style='font-size:12.0pt;mso-bidi-font-size:
  11.0pt;font-family:"Times New Roman",serif;mso-fareast-font-family:ÇÑÄÄ¹ÙÅÁ;
  mso-hansi-font-family:"¸¼Àº °íµñ";color:black'><%=TodayEng%><u><o:p></o:p></u></span></p>
  </td>
 </tr>
</table>

<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
14.0pt;mso-line-height-rule:exactly;tab-stops:-33.75pt 19.8pt 202.65pt 238.65pt 274.65pt 310.65pt 346.65pt 382.65pt 418.65pt 454.65pt'><span
lang=NL style='font-size:12.0pt;mso-bidi-font-size:11.0pt;font-family:"Times New Roman",serif;
mso-fareast-font-family:ÇÑÄÄ¹ÙÅÁ;mso-hansi-font-family:"¸¼Àº °íµñ";color:black;
mso-ansi-language:NL'><%=Email%><o:p></o:p></span></p>

<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;tab-stops:
-33.75pt 19.8pt 202.65pt 238.65pt 274.65pt 310.65pt 346.65pt 382.65pt 418.65pt 454.65pt'><span
lang=NL style='font-size:12.0pt;mso-bidi-font-size:11.0pt;line-height:107%;
font-family:"Times New Roman",serif;mso-fareast-font-family:ÇÑÄÄ¹ÙÅÁ;mso-hansi-font-family:
"¸¼Àº °íµñ";color:black;mso-ansi-language:NL'><o:p>&nbsp;</o:p></span></p>

<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
14.0pt;mso-line-height-rule:exactly;mso-pagination:widow-orphan;word-break:
keep-all'><b style='mso-bidi-font-weight:normal'><span lang=EN-US
style='font-size:12.0pt;mso-bidi-font-size:11.0pt;font-family:"Times New Roman",serif;
mso-fareast-font-family:ÇÑÄÄ¹ÙÅÁ;mso-hansi-font-family:"¸¼Àº °íµñ";color:black'><%=CustName%><o:p></o:p></span></b></p>

<% If LetterAdd1 <> "" Then %>
	<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
	14.0pt;mso-line-height-rule:exactly;mso-pagination:widow-orphan;word-break:
	keep-all'><span lang=EN-US style='font-size:12.0pt;mso-bidi-font-size:11.0pt;
	font-family:"Times New Roman",serif;mso-fareast-font-family:ÇÑÄÄ¹ÙÅÁ;mso-hansi-font-family:
	"¸¼Àº °íµñ";color:black'><%=LetterAdd1%><o:p></o:p></span></p>
<% End If %>

<% If LetterAdd2 <> "" Then %>
	<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
	14.0pt;mso-line-height-rule:exactly;mso-pagination:widow-orphan;word-break:
	keep-all'><span lang=EN-US style='font-size:12.0pt;mso-bidi-font-size:11.0pt;
	font-family:"Times New Roman",serif;mso-fareast-font-family:ÇÑÄÄ¹ÙÅÁ;mso-hansi-font-family:
	"¸¼Àº °íµñ";color:black'><%=LetterAdd2%><o:p></o:p></span></p>
<% End If %>

<% If LetterAdd3 <> "" Then %>
	<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
	14.0pt;mso-line-height-rule:exactly;mso-pagination:widow-orphan;word-break:
	keep-all'><span lang=EN-US style='font-size:12.0pt;mso-bidi-font-size:11.0pt;
	font-family:"Times New Roman",serif;mso-fareast-font-family:ÇÑÄÄ¹ÙÅÁ;mso-hansi-font-family:
	"¸¼Àº °íµñ";color:black'><%=LetterAdd3%><o:p></o:p></span></p>
<% End If %>

<% If LetterAdd4 <> "" Then %>
	<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
	14.0pt;mso-line-height-rule:exactly;mso-pagination:widow-orphan;word-break:
	keep-all'><span lang=EN-US style='font-size:12.0pt;mso-bidi-font-size:11.0pt;
	font-family:"Times New Roman",serif;mso-fareast-font-family:ÇÑÄÄ¹ÙÅÁ;mso-hansi-font-family:
	"¸¼Àº °íµñ";color:black'><%=LetterAdd4%><o:p></o:p></span></p>
<% End If %>

<% If LetterAdd5 <> "" Then %>
	<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
	14.0pt;mso-line-height-rule:exactly;mso-pagination:widow-orphan;word-break:
	keep-all'><span lang=EN-US style='font-size:12.0pt;mso-bidi-font-size:11.0pt;
	font-family:"Times New Roman",serif;mso-fareast-font-family:ÇÑÄÄ¹ÙÅÁ;mso-hansi-font-family:
	"¸¼Àº °íµñ";color:black'><%=LetterAdd5%><o:p></o:p></span></p>
<% End If %>

<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
14.0pt;mso-line-height-rule:exactly;mso-pagination:widow-orphan;word-break:
keep-all'><span lang=EN-US style='font-size:12.0pt;mso-bidi-font-size:11.0pt;
font-family:"Times New Roman",serif;mso-fareast-font-family:ÇÑÄÄ¹ÙÅÁ;mso-hansi-font-family:
"¸¼Àº °íµñ";color:black'><o:p>&nbsp;</o:p></span></p>

<table class=MsoNormalTable border=1 cellspacing=0 cellpadding=0
 style='margin-left:60.0pt;border-collapse:collapse;border:none;mso-border-bottom-alt:
 solid windowtext .5pt;mso-yfti-tbllook:480;mso-padding-alt:0cm 5.4pt 0cm 5.4pt'>
 <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;mso-yfti-lastrow:yes;
  height:21.75pt'>
  <td width=28 valign=top style='width:19.7pt;border:none;border-bottom:solid windowtext 1.0pt;
  mso-border-bottom-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt;
  height:21.75pt'>
  <p class=MsoNormal style='margin-top:0cm;margin-right:-4.0pt;margin-bottom:
  0cm;margin-left:-6.1pt;margin-bottom:.0001pt;mso-para-margin-top:0cm;
  mso-para-margin-right:-.4gd;mso-para-margin-bottom:0cm;mso-para-margin-left:
  -.61gd;mso-para-margin-bottom:.0001pt;line-height:14.0pt;mso-line-height-rule:
  exactly;word-break:keep-all'><span lang=EN-US style='font-size:12.0pt;
  mso-bidi-font-size:11.0pt;font-family:"Times New Roman",serif;mso-hansi-font-family:
  "¸¼Àº °íµñ"'>Re:<o:p></o:p></span></p>
  </td>
  <td width=109 valign=top style='width:81.65pt;border:none;border-bottom:solid windowtext 1.0pt;
  mso-border-bottom-alt:solid windowtext .5pt;padding:0cm 0cm 0cm 1.4pt;
  height:21.75pt'>
  <p class=MsoNormal style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;
  margin-left:-.1pt;margin-bottom:.0001pt;mso-para-margin-top:0cm;mso-para-margin-right:
  0cm;mso-para-margin-bottom:0cm;mso-para-margin-left:-.07gd;mso-para-margin-bottom:
  .0001pt;text-indent:-.6pt;mso-char-indent-count:-.05;line-height:14.0pt;
  mso-line-height-rule:exactly;word-break:keep-all'><b style='mso-bidi-font-weight:
  normal'><span lang=EN-US style='font-size:12.0pt;mso-bidi-font-size:11.0pt;
  font-family:"Times New Roman",serif;mso-hansi-font-family:"¸¼Àº °íµñ"'>Annual
  Report<o:p></o:p></span></b></p>
  </td>
 </tr>
</table>

<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
14.0pt;mso-line-height-rule:exactly;mso-pagination:widow-orphan;word-break:
keep-all'><span lang=EN-US style='font-size:12.0pt;mso-bidi-font-size:11.0pt;
font-family:"Times New Roman",serif;mso-fareast-font-family:ÇÑÄÄ¹ÙÅÁ;mso-hansi-font-family:
"¸¼Àº °íµñ";color:black'><o:p>&nbsp;</o:p></span></p>

<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
14.0pt;mso-line-height-rule:exactly;mso-pagination:widow-orphan;word-break:
keep-all'><span lang=EN-US style='font-size:12.0pt;mso-bidi-font-size:11.0pt;
font-family:"Times New Roman",serif;mso-hansi-font-family:"¸¼Àº °íµñ";mso-bidi-font-weight:
bold'>Dear Sirs: </span><span lang=EN-US style='font-size:12.0pt;mso-bidi-font-size:
11.0pt;font-family:"Times New Roman",serif;mso-fareast-font-family:ÇÑÄÄ¹ÙÅÁ;
mso-hansi-font-family:"¸¼Àº °íµñ";color:black'></span><span lang=EN-US
style='font-size:12.0pt;mso-bidi-font-size:11.0pt;font-family:"Times New Roman",serif;
mso-hansi-font-family:"¸¼Àº °íµñ";mso-bidi-font-weight:bold'></span><span
lang=EN-US style='font-size:12.0pt;mso-bidi-font-size:11.0pt;font-family:"Times New Roman",serif;
mso-fareast-font-family:ÇÑÄÄ¹ÙÅÁ;mso-hansi-font-family:"¸¼Àº °íµñ";color:black'><o:p></o:p></span></p>

<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
14.0pt;mso-line-height-rule:exactly;word-break:keep-all'><span lang=EN-US
style='font-size:12.0pt;mso-bidi-font-size:11.0pt;font-family:"Times New Roman",serif;
mso-hansi-font-family:"¸¼Àº °íµñ"'><o:p>&nbsp;</o:p></span></p>

<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
14.0pt;mso-line-height-rule:exactly;mso-pagination:widow-orphan;word-break:
keep-all'><span lang=EN-US style='font-size:12.0pt;mso-bidi-font-size:11.0pt;
font-family:"Times New Roman",serif;mso-hansi-font-family:"¸¼Àº °íµñ"'>Please find
attached hereto our annual report on the status of the registered patents /
applications which </span><b style='mso-bidi-font-weight:normal'><span
lang=EN-US style='font-size:12.0pt;mso-bidi-font-size:11.0pt;font-family:"Times New Roman",serif;
mso-fareast-font-family:ÇÑÄÄ¹ÙÅÁ;mso-hansi-font-family:"¸¼Àº °íµñ";color:black'><%=CustName%></span></b><span lang=EN-US style='font-size:
12.0pt;mso-bidi-font-size:11.0pt;font-family:"Times New Roman",serif;
mso-hansi-font-family:"¸¼Àº °íµñ"'> has entrusted us to file with the Korean
Intellectual Property Office.<br>
<br>
After reviewing the report, if you find any discrepancies between it and your
records, please let us know.</span><b style='mso-bidi-font-weight:normal'><span
lang=EN-US style='font-size:12.0pt;mso-bidi-font-size:11.0pt;font-family:"Times New Roman",serif;
mso-fareast-font-family:ÇÑÄÄ¹ÙÅÁ;mso-hansi-font-family:"¸¼Àº °íµñ";color:black'><o:p></o:p></span></b></p>

<p class=MsoNormal style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;
margin-left:4.2pt;margin-bottom:.0001pt;line-height:14.0pt;mso-line-height-rule:
exactly;word-break:keep-all'><span lang=EN-US style='font-size:12.0pt;
mso-bidi-font-size:11.0pt;font-family:"Times New Roman",serif;mso-hansi-font-family:
"¸¼Àº °íµñ"'><o:p>&nbsp;</o:p></span></p>

<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
14.0pt;mso-line-height-rule:exactly;word-break:keep-all'><span
style='font-size:12.0pt;mso-bidi-font-size:11.0pt;mso-hansi-font-family:¹ÙÅÁ;
mso-bidi-font-family:¹ÙÅÁ'>¡Ø</span><span lang=EN-US style='font-size:12.0pt;
mso-bidi-font-size:11.0pt;font-family:"Times New Roman",serif;mso-hansi-font-family:
"¸¼Àº °íµñ"'> The report includes the following lists.<o:p></o:p></span></p>

<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0
 style='margin-left:22.7pt;border-collapse:collapse;mso-yfti-tbllook:480;
 mso-padding-alt:0cm 5.4pt 0cm 5.4pt'>
 <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;height:17.3pt'>
  <td width=30 style='width:14.2pt;padding:0cm 5.4pt 0cm 5.4pt;height:17.3pt'>
  <p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
  14.0pt;mso-line-height-rule:exactly;word-break:keep-all'><b style='mso-bidi-font-weight:
  normal'><span lang=EN-US style='font-size:12.0pt;mso-bidi-font-size:11.0pt;
  mso-ascii-font-family:"Times New Roman"'>¡á</span></b><b style='mso-bidi-font-weight:
  normal'><span lang=EN-US style='font-size:12.0pt;mso-bidi-font-size:11.0pt;
  font-family:"Times New Roman",serif;mso-hansi-font-family:"¸¼Àº °íµñ"'><o:p></o:p></span></b></p>
  </td>
  <td width=601 style='width:450.45pt;padding:0cm 5.4pt 0cm 5.4pt;height:17.3pt'>
  <p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
  14.0pt;mso-line-height-rule:exactly;word-break:keep-all'><b style='mso-bidi-font-weight:
  normal'><span lang=EN-US style='font-size:12.0pt;mso-bidi-font-size:11.0pt;
  font-family:"Times New Roman",serif;mso-hansi-font-family:"¸¼Àº °íµñ"'>Statistics of Patents / Applications<o:p></o:p></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:1;height:17.3pt'>
  <td width=30 style='width:14.2pt;padding:0cm 5.4pt 0cm 5.4pt;height:17.3pt'>
  <p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
  14.0pt;mso-line-height-rule:exactly;word-break:keep-all'><b style='mso-bidi-font-weight:
  normal'><span lang=EN-US style='font-size:12.0pt;mso-bidi-font-size:11.0pt;
  mso-ascii-font-family:"Times New Roman"'><%If DataCnt_1_1 = 0 Then Response.Write "¡à" Else Response.Write "¡á" End If%></span></b><span lang=EN-US
  style='font-size:12.0pt;mso-bidi-font-size:11.0pt;font-family:"Times New Roman",serif;
  mso-hansi-font-family:"¸¼Àº °íµñ"'><o:p></o:p></span></p>
  </td>
  <td width=601 style='width:450.45pt;padding:0cm 5.4pt 0cm 5.4pt;height:17.3pt'>
  <p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
  14.0pt;mso-line-height-rule:exactly;word-break:keep-all'><%If DataCnt_1_1 > 0 Then Response.Write "<b style='mso-bidi-font-weight:normal'>" End If%><span lang=EN-US style='font-size:12.0pt;mso-bidi-font-size:11.0pt;
  font-family:"Times New Roman",serif;mso-hansi-font-family:"¸¼Àº °íµñ"'>1-1.
  Registered Patents Administered by HANSUNG<o:p></o:p></span><%If DataCnt_1_1 > 0 Then Response.Write "</b>" End If%></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:3;height:17.3pt'>
  <td width=30 style='width:14.2pt;padding:0cm 5.4pt 0cm 5.4pt;height:17.3pt'>
  <p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
  14.0pt;mso-line-height-rule:exactly;word-break:keep-all'><b style='mso-bidi-font-weight:
  normal'><span lang=EN-US style='font-size:12.0pt;mso-bidi-font-size:11.0pt;
  mso-ascii-font-family:"Times New Roman"'><%If DataCnt_1_2 = 0 Then Response.Write "¡à" Else Response.Write "¡á" End If%></span></b><b style='mso-bidi-font-weight:
  normal'><span lang=EN-US style='font-size:12.0pt;mso-bidi-font-size:11.0pt;
  font-family:"Times New Roman",serif;mso-hansi-font-family:"¸¼Àº °íµñ"'><o:p></o:p></span></b></p>
  </td>
  <td width=601 style='width:450.45pt;padding:0cm 5.4pt 0cm 5.4pt;height:17.3pt'>
  <p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
  14.0pt;mso-line-height-rule:exactly;word-break:keep-all'><%If DataCnt_1_2 > 0 Then Response.Write "<b style='mso-bidi-font-weight:normal'>" End If%><span lang=EN-US style='font-size:12.0pt;mso-bidi-font-size:11.0pt;
  font-family:"Times New Roman",serif;mso-hansi-font-family:"¸¼Àº °íµñ"'>1-2.
  Registered Patents Not Administered by HANSUNG<o:p></o:p></span><%If DataCnt_1_2 > 0 Then Response.Write "</b>" End If%></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:1;height:17.3pt'>
  <td width=30 style='width:14.2pt;padding:0cm 5.4pt 0cm 5.4pt;height:17.3pt'>
  <p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
  14.0pt;mso-line-height-rule:exactly;word-break:keep-all'><b style='mso-bidi-font-weight:
  normal'><span lang=EN-US style='font-size:12.0pt;mso-bidi-font-size:11.0pt;
  mso-ascii-font-family:"Times New Roman"'><%If DataCnt_2_1 = 0 Then Response.Write "¡à" Else Response.Write "¡á" End If%></span></b><span lang=EN-US
  style='font-size:12.0pt;mso-bidi-font-size:11.0pt;font-family:"Times New Roman",serif;
  mso-hansi-font-family:"¸¼Àº °íµñ"'><o:p></o:p></span></p>
  </td>
  <td width=601 style='width:450.45pt;padding:0cm 5.4pt 0cm 5.4pt;height:17.3pt'>
  <p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
  14.0pt;mso-line-height-rule:exactly;word-break:keep-all'><%If DataCnt_2_1 > 0 Then Response.Write "<b style='mso-bidi-font-weight:normal'>" End If%><span lang=EN-US style='font-size:12.0pt;mso-bidi-font-size:11.0pt;
  font-family:"Times New Roman",serif;mso-hansi-font-family:"¸¼Àº °íµñ"'>2-1.
  Applications Abandoned or Instructed to be Abandoned in <%=nDate-1%> (<%=nDate%>)<o:p></o:p></span><%If DataCnt_2_1 > 0 Then Response.Write "</b>" End If%></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:4;height:17.3pt'>
  <td width=30 style='width:14.2pt;padding:0cm 5.4pt 0cm 5.4pt;height:17.3pt'>
  <p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
  14.0pt;mso-line-height-rule:exactly;word-break:keep-all'><b style='mso-bidi-font-weight:
  normal'><span lang=EN-US style='font-size:12.0pt;mso-bidi-font-size:11.0pt;
  mso-ascii-font-family:"Times New Roman"'><%If DataCnt_2_2 = 0 Then Response.Write "¡à" Else Response.Write "¡á" End If%></span></b><b style='mso-bidi-font-weight:
  normal'><span lang=EN-US style='font-size:12.0pt;mso-bidi-font-size:11.0pt;
  font-family:"Times New Roman",serif;mso-hansi-font-family:"¸¼Àº °íµñ"'><o:p></o:p></span></b></p>
  </td>
  <td width=601 style='width:450.45pt;padding:0cm 5.4pt 0cm 5.4pt;height:17.3pt'>
  <p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
  14.0pt;mso-line-height-rule:exactly;word-break:keep-all'><%If DataCnt_2_2 > 0 Then Response.Write "<b style='mso-bidi-font-weight:normal'>" End If%><span lang=EN-US style='font-size:12.0pt;mso-bidi-font-size:11.0pt;
  font-family:"Times New Roman",serif;mso-hansi-font-family:"¸¼Àº °íµñ"'>2-2.
  Applications Abandoned by December 31, <%=nDate-2%><o:p></o:p></span><%If DataCnt_2_2 > 0 Then Response.Write "</b>" End If%></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:5;mso-yfti-lastrow:yes;height:17.3pt'>
  <td width=30 style='width:14.2pt;padding:0cm 5.4pt 0cm 5.4pt;height:17.3pt'>
  <p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
  14.0pt;mso-line-height-rule:exactly;word-break:keep-all'><b style='mso-bidi-font-weight:
  normal'><span lang=EN-US style='font-size:12.0pt;mso-bidi-font-size:11.0pt;
  mso-ascii-font-family:"Times New Roman"'><%If DataCnt_3 = 0 Then Response.Write "¡à" Else Response.Write "¡á" End If%></span></b><b style='mso-bidi-font-weight:
  normal'><span lang=EN-US style='font-size:12.0pt;mso-bidi-font-size:11.0pt;
  font-family:"Times New Roman",serif;mso-hansi-font-family:"¸¼Àº °íµñ"'><o:p></o:p></span></b></p>
  </td>
  <td width=601 style='width:450.45pt;padding:0cm 5.4pt 0cm 5.4pt;height:17.3pt'>
  <p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
  14.0pt;mso-line-height-rule:exactly;word-break:keep-all'><%If DataCnt_3 > 0 Then Response.Write "<b style='mso-bidi-font-weight:normal'>" End If%><span lang=EN-US style='font-size:12.0pt;mso-bidi-font-size:11.0pt;
  font-family:"Times New Roman",serif;mso-hansi-font-family:"¸¼Àº °íµñ"'>3. Pending
  Applications<o:p></o:p></span><%If DataCnt_3 > 0 Then Response.Write "</b>" End If%></p>
  </td>
 </tr>
</table>

<p class=MsoNormal align=right style='margin-left:-.1pt;mso-para-margin-left:
-.01gd;text-align:right;text-indent:.05pt;tab-stops:-33.75pt 19.8pt 202.65pt 238.65pt 274.65pt 310.65pt 346.65pt 382.65pt 418.65pt 454.65pt'><span
lang=EN-US style='font-size:12.0pt;mso-bidi-font-size:11.0pt;line-height:107%;
font-family:"Times New Roman",serif;mso-hansi-font-family:"¸¼Àº °íµñ"'><o:p>&nbsp;</o:p></span></p>


<p class=MsoNormal align=right style='margin-left:-.1pt;mso-para-margin-left:
-.01gd;text-align:right;text-indent:.05pt;tab-stops:-33.75pt 19.8pt 202.65pt 238.65pt 274.65pt 310.65pt 346.65pt 382.65pt 418.65pt 454.65pt'><span
lang=EN-US style='font-size:12.0pt;mso-bidi-font-size:11.0pt;line-height:106%;
font-family:"Times New Roman",serif;mso-hansi-font-family:"¸¼Àº °íµñ";mso-no-proof:
yes'><!--[if gte vml 1]><v:shape id="±×¸²_x0020_2" o:spid="_x0000_i1025" type="#_x0000_t75"
 alt="Untitled" style='width:192pt;height:99.75pt;visibility:visible;
 mso-wrap-style:square'>
 <v:imagedata src="http://<%=Request.ServerVariables("SERVER_NAME") & ":" & Request.ServerVariables("SERVER_PORT")%>/images_letter/RPABE0084E.png" o:title="Untitled"/>
</v:shape><![endif]--><![if !vml]><img width=256 height=133
src="http://<%=Request.ServerVariables("SERVER_NAME") & ":" & Request.ServerVariables("SERVER_PORT")%>/images_letter/RPABE0084E.png" alt=Untitled v:shapes="±×¸²_x0020_2"><![endif]></span><span
lang=EN-US style='font-size:12.0pt;mso-bidi-font-size:11.0pt;line-height:106%;
font-family:"Times New Roman",serif;mso-hansi-font-family:"¸¼Àº °íµñ"'><o:p></o:p></span></p>

<p class=MsoNormal style='tab-stops:-33.75pt 19.8pt 202.65pt 238.65pt 274.65pt 310.65pt 346.65pt 382.65pt 418.65pt 454.65pt'><span
lang=EN-US style='font-size:12.0pt;mso-bidi-font-size:11.0pt;line-height:107%;
font-family:"Times New Roman",serif;mso-hansi-font-family:"¸¼Àº °íµñ"'>KPC/JWY/JSC<o:p></o:p></span></p>

<b style='mso-bidi-font-weight:normal'><span lang=EN-US style='font-size:16.0pt;
line-height:107%;font-family:"Times New Roman",serif;mso-fareast-font-family:
"¸¼Àº °íµñ";mso-font-kerning:1.0pt;mso-ansi-language:EN-US;mso-fareast-language:
KO;mso-bidi-language:AR-SA'><br clear=all style='page-break-before:always'>
</span></b>

<p class=MsoNormal align=center style='margin-bottom:0cm;margin-bottom:.0001pt;
text-align:center'><b style='mso-bidi-font-weight:normal'><span lang=EN-US
style='font-size:16.0pt;line-height:107%;font-family:"Times New Roman",serif'>Statistics of Patents / <span class=GramE>Applications</span></span></b><span lang=EN-US
style='font-size:16.0pt;line-height:107%;font-family:"Times New Roman",serif'><br>
</span><b style='mso-bidi-font-weight:normal'><span lang=EN-US
style='mso-bidi-font-size:10.0pt;line-height:107%;font-family:"Times New Roman",serif'>(<%=CustName%>)</span></b><span lang=EN-US style='mso-bidi-font-size:10.0pt;
line-height:107%;font-family:"Times New Roman",serif'><o:p></o:p></span></p>

<p class=MsoNormal align=center style='margin-bottom:0cm;margin-bottom:.0001pt;
text-align:center'><span lang=EN-US style='mso-bidi-font-size:10.0pt;
line-height:107%;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p>

<p class=MsoNormal align=center style='margin-bottom:0cm;margin-bottom:.0001pt;
text-align:center'><span lang=EN-US style='mso-bidi-font-size:10.0pt;
line-height:107%;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p>


<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0
 style='border-collapse:collapse;mso-yfti-tbllook:1184;mso-padding-alt:0cm 5.4pt 0cm 5.4pt;
 mso-border-insideh:.5pt solid windowtext'>
 <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'>
  <td width=525 colspan=9 valign=top style='width:393.85pt;border:none;
  border-bottom:solid windowtext 1.0pt;mso-border-bottom-alt:solid windowtext .5pt;
  padding:0cm 0cm 0cm 0cm'>
  <p class=MsoNormal align=left style='margin-bottom:0cm;margin-bottom:.0001pt;
  text-align:left;line-height:12.0pt'><b style='mso-bidi-font-weight:normal'><span
  lang=EN-US style='mso-bidi-font-size:10.0pt;font-family:"Times New Roman",serif'>HANSUNG Intellectual Property<o:p></o:p></span></b></p>
  </td>
  <td width=162 colspan=4 valign=top style='width:116.35pt;border:none;
  border-bottom:solid windowtext 1.0pt;mso-border-bottom-alt:solid windowtext .5pt;
  padding:0cm 0cm 0cm 0cm'>
  <p class=MsoNormal align=right style='margin-bottom:0cm;margin-bottom:.0001pt;
  text-align:right;line-height:12.0pt'><b style='mso-bidi-font-weight:normal'><span
  lang=EN-US style='mso-bidi-font-size:10.0pt;font-family:"Times New Roman",serif'><%=TodayEng%><o:p></o:p></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:1;height:22.7pt'>
  <td width=237 colspan=2 style='width:177.75pt;border-top:none;border-left:
  solid windowtext 1.0pt;border-bottom:solid windowtext 1.5pt;border-right:
  solid windowtext 1.0pt;mso-border-top-alt:solid windowtext .5pt;mso-border-alt:
  solid windowtext .5pt;mso-border-bottom-alt:solid windowtext 1.5pt;
  padding:0cm 5.4pt 0cm 5.4pt;height:22.7pt'>
  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
  center;line-height:150%;word-break:keep-all'><b style='mso-bidi-font-weight:
  normal'><span lang=EN-US style='font-size:8.0pt;line-height:150%;font-family:
  "Times New Roman",serif'>YEAR<o:p></o:p></span></b></p>
  </td>
  <%
  For i = sDate To nDate
%>
	<td width=40 style='width:30.3pt;border-top:none;border-left:none;border-bottom:
	  solid windowtext 1.5pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:
	  solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:
	  solid windowtext .5pt;mso-border-bottom-alt:solid windowtext 1.5pt;
	  padding:0cm 1.4pt 0cm 1.4pt;height:22.7pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:150%;word-break:keep-all'><b style='mso-bidi-font-weight:
	  normal'><span lang=EN-US style='font-size:8.0pt;line-height:150%;font-family:
	  "Times New Roman",serif'>
<%
	If i = sDate Then
		Response.Write "~" & i
	Else
		Response.Write i
	End If
%><o:p></o:p></span></b></p>
	  </td>
<%
  Next
  %>
  <td width=41 style='width:30.45pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.5pt;border-right:solid windowtext 1.0pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-alt:solid windowtext .5pt;mso-border-bottom-alt:solid windowtext 1.5pt;
  padding:0cm 1.4pt 0cm 1.4pt;height:22.7pt'>
  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
  center;line-height:150%;word-break:keep-all'><b style='mso-bidi-font-weight:
  normal'><span lang=EN-US style='font-size:7.0pt;line-height:150%;font-family:
  "Times New Roman",serif'>TOTAL<o:p></o:p></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:2;height:22.7pt'>
  <td width=174 rowspan=4 style='width:130.55pt;border-top:none;border-left:
  solid windowtext 1.0pt;border-bottom:double windowtext 1.5pt;border-right:
  solid windowtext 1.0pt;mso-border-top-alt:solid windowtext 1.5pt;mso-border-top-alt:
  solid 1.5pt;mso-border-left-alt:solid .5pt;mso-border-bottom-alt:double 1.5pt;
  mso-border-right-alt:solid .5pt;mso-border-color-alt:windowtext;padding:0cm 5.4pt 0cm 5.4pt;
  height:22.7pt'>
  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
  center;line-height:150%;word-break:keep-all'><b style='mso-bidi-font-weight:
  normal'><span lang=EN-US style='font-size:8.0pt;line-height:150%;font-family:
  "Times New Roman",serif'>Number of applications<o:p></o:p></span></b></p>
  </td>
  <td width=63 style='width:47.2pt;border-top:none;border-left:none;border-bottom:
  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:
  solid windowtext 1.5pt;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:
  solid windowtext .5pt;mso-border-top-alt:solid windowtext 1.5pt;padding:0cm 5.4pt 0cm 5.4pt;
  height:22.7pt'>
  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
  center;line-height:150%;word-break:keep-all'><b style='mso-bidi-font-weight:
  normal'><span lang=EN-US style='font-size:8.0pt;line-height:150%;font-family:
  "Times New Roman",serif'><%=appCustRef(0)%><o:p></o:p></span></b></p>
  </td>
  <%
For i = 0 To 9 '´ëÇ¥Ãâ¿øÀÎ1
%>
	<td width=40 style='width:30.3pt;border-top:none;border-left:none;border-bottom:
	  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:
	  solid windowtext 1.5pt;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:
	  solid windowtext .5pt;mso-border-top-alt:solid windowtext 1.5pt;padding:0cm 1.4pt 0cm 1.4pt;
	  height:22.7pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:150%;word-break:keep-all'><span lang=EN-US
	  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=arStatisApplName1(i)%><o:p></o:p></span></p>
	  </td>
 <%
 Next
 %>   
  <td width=41 style='width:30.45pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-top-alt:solid windowtext 1.5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-alt:solid windowtext .5pt;mso-border-top-alt:solid windowtext 1.5pt;
  padding:0cm 1.4pt 0cm 1.4pt;height:22.7pt'>
  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
  center;line-height:150%;word-break:keep-all'><span lang=EN-US
  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=sumStatisApplName1%><o:p></o:p></span></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:3;height:22.7pt'>
  <td width=63 style='width:47.2pt;border-top:none;border-left:none;border-bottom:
  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:
  solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:
  solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt;height:22.7pt'>
  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
  center;line-height:150%;word-break:keep-all'><b style='mso-bidi-font-weight:
  normal'><span lang=EN-US style='font-size:8.0pt;line-height:150%;font-family:
  "Times New Roman",serif'><%=appCustRef(1)%><o:p></o:p></span></b></p></td>
<%
For i = 0 To 9 '´ëÇ¥Ãâ¿øÀÎ2
%>
	<td width=40 style='width:30.3pt;border-top:none;border-left:none;border-bottom:
	  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:
	  solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:
	  solid windowtext .5pt;padding:0cm 1.4pt 0cm 1.4pt;height:22.7pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:150%;word-break:keep-all'><span lang=EN-US
	  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=arStatisApplName2(i)%><o:p></o:p></span></p>
	  </td>
 <%
 Next
 %>   
  <td width=41 style='width:30.45pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-alt:solid windowtext .5pt;padding:0cm 1.4pt 0cm 1.4pt;height:22.7pt'>
  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
  center;line-height:150%;word-break:keep-all'><span lang=EN-US
  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=sumStatisApplName2%><o:p></o:p></span></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:4;height:22.7pt'>
  <td width=63 style='width:47.2pt;border-top:none;border-left:none;border-bottom:
  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:
  solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:
  solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt;height:22.7pt'>
  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
  center;line-height:150%;word-break:keep-all'><b style='mso-bidi-font-weight:
  normal'><span lang=EN-US style='font-size:8.0pt;line-height:150%;font-family:
  "Times New Roman",serif'>etc.<o:p></o:p></span></b></p>
  </td>
 <%
For i = 0 To 9 '±âÅ¸Ãâ¿øÀÎ
%>
	 <td width=40 style='width:30.3pt;border-top:none;border-left:none;border-bottom:
	  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:
	  solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:
	  solid windowtext .5pt;padding:0cm 1.4pt 0cm 1.4pt;height:22.7pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:150%;word-break:keep-all'><span lang=EN-US
	  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=arStatisApplEtc(i)%><o:p></o:p></span></p>
	  </td>
 <%
 Next
 %> 
  <td width=41 style='width:30.45pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-alt:solid windowtext .5pt;padding:0cm 1.4pt 0cm 1.4pt;height:22.7pt'>
  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
  center;line-height:150%;word-break:keep-all'><span lang=EN-US
  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=sumStatisApplEtc%><o:p></o:p></span></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:5;height:22.7pt'>
  <td width=63 style='width:47.2pt;border-top:none;border-left:none;border-bottom:
  double windowtext 1.5pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:
  solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:
  solid windowtext .5pt;mso-border-bottom-alt:double windowtext 1.5pt;
  padding:0cm 5.4pt 0cm 5.4pt;height:22.7pt'>
  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
  center;line-height:150%;word-break:keep-all'><b style='mso-bidi-font-weight:
  normal'><span lang=EN-US style='font-size:8.0pt;line-height:150%;font-family:
  "Times New Roman",serif'>Total<o:p></o:p></span></b></p>
  </td>
  <%
For i = 0 To 9 'Total
%>
	 <td width=40 style='width:30.3pt;border-top:none;border-left:none;border-bottom:
	  double windowtext 1.5pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:
	  solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:
	  solid windowtext .5pt;mso-border-bottom-alt:double windowtext 1.5pt;
	  padding:0cm 1.4pt 0cm 1.4pt;height:22.7pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:150%;word-break:keep-all'><span lang=EN-US
	  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=arStatisApplName1(i) + arStatisApplName2(i) + arStatisApplEtc(i)%><o:p></o:p></span></p>
	  </td>
 <%
 Next
 %> 
 
  <td width=41 style='width:30.45pt;border-top:none;border-left:none;
  border-bottom:double windowtext 1.5pt;border-right:solid windowtext 1.0pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-alt:solid windowtext .5pt;mso-border-bottom-alt:double windowtext 1.5pt;
  padding:0cm 1.4pt 0cm 1.4pt;height:22.7pt'>
  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
  center;line-height:150%;word-break:keep-all'><span lang=EN-US
  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=sumStatisApplName1 + sumStatisApplName2 + sumStatisApplEtc%><o:p></o:p></span></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:6;height:22.7pt'>
  <td width=237 colspan=2 style='width:177.75pt;border:solid windowtext 1.0pt;
  border-top:none;mso-border-top-alt:double windowtext 1.5pt;mso-border-alt:
  solid windowtext .5pt;mso-border-top-alt:double windowtext 1.5pt;padding:
  0cm 5.4pt 0cm 5.4pt;height:22.7pt'>
  <p class=MsoNormal align=left style='mso-margin-bottom-alt:auto;text-align:
  left;line-height:150%;word-break:keep-all'><b style='mso-bidi-font-weight:
  normal'><span lang=EN-US style='font-size:8.0pt;line-height:150%;font-family:
  "Times New Roman",serif'>1. Patents Registered<o:p></o:p></span></b></p>
  </td>
  <%
  For i = 0 To 9 '1.µî·Ï°Ç
  %>
	<td width=40 style='width:30.3pt;border-top:none;border-left:none;border-bottom:
	  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:
	  double windowtext 1.5pt;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:
	  solid windowtext .5pt;mso-border-top-alt:double windowtext 1.5pt;padding:
	  0cm 1.4pt 0cm 1.4pt;height:22.7pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:150%;word-break:keep-all'><span lang=EN-US
	  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=arStatisReg(i)%><o:p></o:p></span></p>
	  </td>
  <%
  Next
  %>
  <td width=41 style='width:30.45pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-top-alt:double windowtext 1.5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-alt:solid windowtext .5pt;mso-border-top-alt:double windowtext 1.5pt;
  padding:0cm 1.4pt 0cm 1.4pt;height:22.7pt'>
  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
  center;line-height:150%;word-break:keep-all'><span lang=EN-US
  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=sumStatisReg%><o:p></o:p></span></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:7;height:22.7pt'>
  <td width=237 colspan=2 style='width:177.75pt;border:solid windowtext 1.0pt;
  border-top:none;mso-border-top-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;
  padding:0cm 5.4pt 0cm 5.4pt;height:22.7pt'>
  <p class=MsoNormal align=left style='mso-margin-bottom-alt:auto;text-align:
  left;line-height:150%;word-break:keep-all'><b style='mso-bidi-font-weight:
  normal'><span lang=EN-US style='font-size:8.0pt;line-height:150%;font-family:
  "Times New Roman",serif'>2. (To be) Abandoned<o:p></o:p></span></b></p>
  </td>
   <%
  For i = 0 To 9 '2.Æ÷±â°Ç
  %>
	<td width=40 style='width:30.3pt;border-top:none;border-left:none;border-bottom:
	  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:
	  solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:
	  solid windowtext .5pt;padding:0cm 1.4pt 0cm 1.4pt;height:22.7pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:150%;word-break:keep-all'><span lang=EN-US
	  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=arStatisAban(i)%><o:p></o:p></span></p>
	  </td>
  <%
  Next
  %> 
  <td width=41 style='width:30.45pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-alt:solid windowtext .5pt;padding:0cm 1.4pt 0cm 1.4pt;height:22.7pt'>
  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
  center;line-height:150%;word-break:keep-all'><span lang=EN-US
  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=sumStatisAban%><o:p></o:p></span></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:8;height:22.7pt'>
  <td width=237 colspan=2 style='width:177.75pt;border-top:none;border-left:
  solid windowtext 1.0pt;border-bottom:solid windowtext 1.5pt;border-right:
  solid windowtext 1.0pt;mso-border-top-alt:solid windowtext .5pt;mso-border-alt:
  solid windowtext .5pt;mso-border-bottom-alt:solid windowtext 1.5pt;
  padding:0cm 5.4pt 0cm 5.4pt;height:22.7pt'>
  <p class=MsoNormal align=left style='mso-margin-bottom-alt:auto;text-align:
  left;line-height:150%;word-break:keep-all'><b style='mso-bidi-font-weight:
  normal'><span lang=EN-US style='font-size:8.0pt;line-height:150%;font-family:
  "Times New Roman",serif'>3. Pending<o:p></o:p></span></b></p>
  </td>
  <%
  For i = 0 To 9 '3.°è·ù°Ç
  %>
	<td width=40 style='width:30.3pt;border-top:none;border-left:none;border-bottom:
	  solid windowtext 1.5pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:
	  solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:
	  solid windowtext .5pt;mso-border-bottom-alt:solid windowtext 1.5pt;
	  padding:0cm 1.4pt 0cm 1.4pt;height:22.7pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:150%;word-break:keep-all'><span lang=EN-US
	  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=arStatisPen(i)%><o:p></o:p></span></p>
	  </td>
  <%
  Next
  %>
  <td width=41 style='width:30.45pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.5pt;border-right:solid windowtext 1.0pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-alt:solid windowtext .5pt;mso-border-bottom-alt:solid windowtext 1.5pt;
  padding:0cm 1.4pt 0cm 1.4pt;height:22.7pt'>
  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
  center;line-height:150%;word-break:keep-all'><span lang=EN-US
  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=sumStatisPen%><o:p></o:p></span></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:9;height:22.7pt'>
  <td width=237 colspan=2 style='width:177.75pt;border:solid windowtext 1.0pt;
  border-top:none;mso-border-top-alt:solid windowtext 1.5pt;mso-border-alt:
  solid windowtext .5pt;mso-border-top-alt:solid windowtext 1.5pt;padding:0cm 5.4pt 0cm 5.4pt;
  height:22.7pt'>
  <p class=MsoNormal align=left style='mso-margin-bottom-alt:auto;text-align:
  left;line-height:150%;word-break:keep-all'><b style='mso-bidi-font-weight:
  normal'><span lang=EN-US style='font-size:8.0pt;line-height:150%;font-family:
  "Times New Roman",serif'><span style='mso-spacerun:yes'>&nbsp; </span><span
  style='mso-spacerun:yes'>&nbsp;&nbsp;</span>a. No Request for Examination<o:p></o:p></span></b></p>
  </td>
  <%
  For i = 0 To 9 '3-a. ½É»ç¹ÌÃ»±¸
  %>
	<td width=40 style='width:30.3pt;border-top:none;border-left:none;border-bottom:
	  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:
	  solid windowtext 1.5pt;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:
	  solid windowtext .5pt;mso-border-top-alt:solid windowtext 1.5pt;padding:0cm 1.4pt 0cm 1.4pt;
	  height:22.7pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:150%;word-break:keep-all'><span lang=EN-US
	  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=arStatisNoExam(i)%><o:p></o:p></span></p>
	  </td>
  <%
  Next
  %> 
  <td width=41 style='width:30.45pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-top-alt:solid windowtext 1.5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-alt:solid windowtext .5pt;mso-border-top-alt:solid windowtext 1.5pt;
  padding:0cm 1.4pt 0cm 1.4pt;height:22.7pt'>
  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
  center;line-height:150%;word-break:keep-all'><span lang=EN-US
  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=sumStatisNoExam%><o:p></o:p></span></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:10;height:22.7pt'>
  <td width=237 colspan=2 style='width:177.75pt;border-top:none;border-left:
  solid windowtext 1.0pt;border-bottom:solid windowtext 1.5pt;border-right:
  solid windowtext 1.0pt;mso-border-top-alt:solid windowtext .5pt;mso-border-alt:
  solid windowtext .5pt;mso-border-bottom-alt:solid windowtext 1.5pt;
  padding:0cm 5.4pt 0cm 5.4pt;height:22.7pt'>
  <p class=MsoNormal align=left style='mso-margin-bottom-alt:auto;text-align:
  left;line-height:150%;word-break:keep-all'><b style='mso-bidi-font-weight:
  normal'><span lang=EN-US style='font-size:8.0pt;line-height:150%;font-family:
  "Times New Roman",serif'><span style='mso-spacerun:yes'>&nbsp;</span><span
  style='mso-spacerun:yes'>&nbsp; </span><span
  style='mso-spacerun:yes'>&nbsp;</span>b. Requested for Examination<o:p></o:p></span></b></p>
  </td>
  <%
  For i = 0 To 9 '3-b. ½É»çÃ»±¸
  %>
	 <td width=40 style='width:30.3pt;border-top:none;border-left:none;border-bottom:
	  solid windowtext 1.5pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:
	  solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:
	  solid windowtext .5pt;mso-border-bottom-alt:solid windowtext 1.5pt;
	  padding:0cm 1.4pt 0cm 1.4pt;height:22.7pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:150%;word-break:keep-all'><span lang=EN-US
	  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=arStatisExam(i)%><o:p></o:p></span></p>
	  </td>
  <%
  Next
  %>
  <td width=41 style='width:30.45pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.5pt;border-right:solid windowtext 1.0pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-alt:solid windowtext .5pt;mso-border-bottom-alt:solid windowtext 1.5pt;
  padding:0cm 1.4pt 0cm 1.4pt;height:22.7pt'>
  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
  center;line-height:150%;word-break:keep-all'><span lang=EN-US
  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=sumStatisExam%><o:p></o:p></span></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:11;height:22.7pt'>
  <td width=237 colspan=2 style='width:177.75pt;border:solid windowtext 1.0pt;
  border-top:none;mso-border-top-alt:solid windowtext 1.5pt;mso-border-alt:
  solid windowtext .5pt;mso-border-top-alt:solid windowtext 1.5pt;padding:0cm 5.4pt 0cm 5.4pt;
  height:22.7pt'>
  <p class=MsoNormal align=left style='mso-margin-bottom-alt:auto;text-align:
  left;line-height:150%;word-break:keep-all'><b style='mso-bidi-font-weight:
  normal'><span lang=EN-US style='font-size:8.0pt;line-height:150%;font-family:
  "Times New Roman",serif'><span style='mso-spacerun:yes'>&nbsp;&nbsp; </span><span
  style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;</span><span
  style='mso-spacerun:yes'>&nbsp;</span>1) Waiting for Examination<o:p></o:p></span></b></p>
  </td>
  <%
  For i = 0 To 9 '3-b-1. ½É»çÁß
  %>
	 <td width=40 style='width:30.3pt;border-top:none;border-left:none;border-bottom:
	  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:
	  solid windowtext 1.5pt;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:
	  solid windowtext .5pt;mso-border-top-alt:solid windowtext 1.5pt;padding:0cm 1.4pt 0cm 1.4pt;
	  height:22.7pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:150%;word-break:keep-all'><span lang=EN-US
	  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=arStatisWaitExam(i)%><o:p></o:p></span></p>
	  </td>
  <%
  Next
  %>
  <td width=41 style='width:30.45pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-top-alt:solid windowtext 1.5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-alt:solid windowtext .5pt;mso-border-top-alt:solid windowtext 1.5pt;
  padding:0cm 1.4pt 0cm 1.4pt;height:22.7pt'>
  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
  center;line-height:150%;word-break:keep-all'><span lang=EN-US
  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=sumStatisWaitExam%><o:p></o:p></span></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:12;height:22.7pt'>
  <td width=237 colspan=2 style='width:177.75pt;border-top:none;border-left:
  solid windowtext 1.0pt;border-bottom:solid windowtext 1.5pt;border-right:
  solid windowtext 1.0pt;mso-border-top-alt:solid windowtext .5pt;mso-border-alt:
  solid windowtext .5pt;mso-border-bottom-alt:solid windowtext 1.5pt;
  padding:0cm 5.4pt 0cm 5.4pt;height:22.7pt'>
  <p class=MsoNormal align=left style='mso-margin-bottom-alt:auto;text-align:
  left;line-height:150%;word-break:keep-all'><b style='mso-bidi-font-weight:
  normal'><span lang=EN-US style='font-size:8.0pt;line-height:150%;font-family:
  "Times New Roman",serif'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;
  </span><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;</span>2) Under
  Examination<o:p></o:p></span></b></p>
  </td>
  <%
  For i = 0 To 9 '3-b-2. ½É»ç¿Ï·á
  %>
	 <td width=40 style='width:30.3pt;border-top:none;border-left:none;border-bottom:
	  solid windowtext 1.5pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:
	  solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:
	  solid windowtext .5pt;mso-border-bottom-alt:solid windowtext 1.5pt;
	  padding:0cm 1.4pt 0cm 1.4pt;height:22.7pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:150%;word-break:keep-all'><span lang=EN-US
	  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=arStatisUnderExam(i)%><o:p></o:p></span></p>
	  </td>
  <%
  Next
  %>
  <td width=41 style='width:30.45pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.5pt;border-right:solid windowtext 1.0pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-alt:solid windowtext .5pt;mso-border-bottom-alt:solid windowtext 1.5pt;
  padding:0cm 1.4pt 0cm 1.4pt;height:22.7pt'>
  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
  center;line-height:150%;word-break:keep-all'><span lang=EN-US
  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=sumStatisUnderExam%><o:p></o:p></span></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:13;height:22.7pt'>
  <td width=237 colspan=2 style='width:177.75pt;border:solid windowtext 1.0pt;
  border-top:none;mso-border-top-alt:solid windowtext 1.5pt;mso-border-alt:
  solid windowtext .5pt;mso-border-top-alt:solid windowtext 1.5pt;padding:0cm 5.4pt 0cm 5.4pt;
  height:22.7pt'>
  <p class=MsoNormal align=left style='mso-margin-bottom-alt:auto;text-align:
  left;line-height:150%;word-break:keep-all'><b style='mso-bidi-font-weight:
  normal'><span lang=EN-US style='font-size:8.0pt;line-height:150%;font-family:
  "Times New Roman",serif'><span
  style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
  style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>a)
  Notice of Allowance<o:p></o:p></span></b></p>
  </td>
  <%
  For i = 0 To 9 '3-b-2- a. µî·Ï°áÁ¤
  %>
	<td width=40 style='width:30.3pt;border-top:none;border-left:none;border-bottom:
	  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:
	  solid windowtext 1.5pt;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:
	  solid windowtext .5pt;mso-border-top-alt:solid windowtext 1.5pt;padding:0cm 1.4pt 0cm 1.4pt;
	  height:22.7pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:150%;word-break:keep-all'><span lang=EN-US
	  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=arStatisNoticeExam(i)%><o:p></o:p></span></p>
	  </td>
  <%
  Next
  %> 
  <td width=41 style='width:30.45pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-top-alt:solid windowtext 1.5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-alt:solid windowtext .5pt;mso-border-top-alt:solid windowtext 1.5pt;
  padding:0cm 1.4pt 0cm 1.4pt;height:22.7pt'>
  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
  center;line-height:150%;word-break:keep-all'><span lang=EN-US
  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=sumStatisNoticeExam%><o:p></o:p></span></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:14;height:22.7pt'>
  <td width=237 colspan=2 style='width:177.75pt;border:solid windowtext 1.0pt;
  border-top:none;mso-border-top-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;
  padding:0cm 5.4pt 0cm 5.4pt;height:22.7pt'>
  <p class=MsoNormal align=left style='mso-margin-bottom-alt:auto;text-align:
  left;line-height:150%;word-break:keep-all'><b style='mso-bidi-font-weight:
  normal'><span lang=EN-US style='font-size:8.0pt;line-height:150%;font-family:
  "Times New Roman",serif'><span
  style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><span
  style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>b)
  Office Action<o:p></o:p></span></b></p>
  </td>
  <%
  For i = 0 To 9 '3-b-2- b. OAÁøÇà
  %>
	 <td width=40 style='width:30.3pt;border-top:none;border-left:none;border-bottom:
	  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:
	  solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:
	  solid windowtext .5pt;padding:0cm 1.4pt 0cm 1.4pt;height:22.7pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:150%;word-break:keep-all'><span lang=EN-US
	  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=arStatisOAExam(i)%><o:p></o:p></span></p>
	  </td>
  <%
  Next
  %> 
  <td width=41 style='width:30.45pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-alt:solid windowtext .5pt;padding:0cm 1.4pt 0cm 1.4pt;height:22.7pt'>
  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
  center;line-height:150%;word-break:keep-all'><span lang=EN-US
  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=sumStatisOAExam%><o:p></o:p></span></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:15;height:22.7pt'>
  <td width=237 colspan=2 style='width:177.75pt;border:solid windowtext 1.0pt;
  border-top:none;mso-border-top-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;
  padding:0cm 5.4pt 0cm 5.4pt;height:22.7pt'>
  <p class=MsoNormal align=left style='mso-margin-bottom-alt:auto;text-align:
  left;line-height:150%;word-break:keep-all'><b style='mso-bidi-font-weight:
  normal'><span lang=EN-US style='font-size:8.0pt;line-height:150%;font-family:
  "Times New Roman",serif'><span
  style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span><span
  style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span><span
  style='mso-spacerun:yes'>&nbsp;</span>c) Appealed to IPT<o:p></o:p></span></b></p>
  </td>
   <%
  For i = 0 To 9 '3-b-2- c. Æ¯Çã½ÉÆÇ¿ø
  %>
	<td width=40 style='width:30.3pt;border-top:none;border-left:none;border-bottom:
	  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:
	  solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:
	  solid windowtext .5pt;padding:0cm 1.4pt 0cm 1.4pt;height:22.7pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:150%;word-break:keep-all'><span lang=EN-US
	  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=arStatisIPTExam(i)%><o:p></o:p></span></p>
	  </td>
  <%
  Next
  %> 
  <td width=41 style='width:30.45pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-alt:solid windowtext .5pt;padding:0cm 1.4pt 0cm 1.4pt;height:22.7pt'>
  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
  center;line-height:150%;word-break:keep-all'><span lang=EN-US
  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=sumStatisIPTExam%><o:p></o:p></span></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:16;height:22.7pt'>
  <td width=237 colspan=2 style='width:177.75pt;border:solid windowtext 1.0pt;
  border-top:none;mso-border-top-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;
  padding:0cm 5.4pt 0cm 5.4pt;height:22.7pt'>
  <p class=MsoNormal align=left style='mso-margin-bottom-alt:auto;text-align:
  left;line-height:150%;word-break:keep-all'><b style='mso-bidi-font-weight:
  normal'><span lang=EN-US style='font-size:8.0pt;line-height:150%;font-family:
  "Times New Roman",serif'><span style='mso-spacerun:yes'>&nbsp;</span><span
  style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp; </span><span
  style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span><span
  style='mso-spacerun:yes'>&nbsp;</span>d) Appealed to P.C.<o:p></o:p></span></b></p>
  </td>
   <%
  For i = 0 To 9 '3-b-2- d. Æ¯Çã¹ý¿ø
  %>
	<td width=40 style='width:30.3pt;border-top:none;border-left:none;border-bottom:
	  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:
	  solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:
	  solid windowtext .5pt;padding:0cm 1.4pt 0cm 1.4pt;height:22.7pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:150%;word-break:keep-all'><span lang=EN-US
	  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=arStatisPCExam(i)%><o:p></o:p></span></p>
	  </td>
  <%
  Next
  %>
  <td width=41 style='width:30.45pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-alt:solid windowtext .5pt;padding:0cm 1.4pt 0cm 1.4pt;height:22.7pt'>
  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
  center;line-height:150%;word-break:keep-all'><span lang=EN-US
  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=sumStatisPCExam%><o:p></o:p></span></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:17;mso-yfti-lastrow:yes;height:22.7pt'>
  <td width=237 colspan=2 style='width:177.75pt;border:solid windowtext 1.0pt;
  border-top:none;mso-border-top-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;
  padding:0cm 5.4pt 0cm 5.4pt;height:22.7pt'>
  <p class=MsoNormal align=left style='mso-margin-bottom-alt:auto;text-align:
  left;line-height:150%;word-break:keep-all'><b style='mso-bidi-font-weight:
  normal'><span lang=EN-US style='font-size:8.0pt;line-height:150%;font-family:
  "Times New Roman",serif'><span
  style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp; </span><span
  style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span><span
  style='mso-spacerun:yes'>&nbsp;</span>e) Appealed to S.C.<o:p></o:p></span></b></p>
  </td>
   <%
  For i = 0 To 9 '3-b-2- e. ´ë¹ý¿ø
  %>
	<td width=40 style='width:30.3pt;border-top:none;border-left:none;border-bottom:
	  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:
	  solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:
	  solid windowtext .5pt;padding:0cm 1.4pt 0cm 1.4pt;height:22.7pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:150%;word-break:keep-all'><span lang=EN-US
	  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=arStatisSCExam(i)%><o:p></o:p></span></p>
	  </td>
  <%
  Next
  %> 
  <td width=41 style='width:30.45pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-alt:solid windowtext .5pt;padding:0cm 1.4pt 0cm 1.4pt;height:22.7pt'>
  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
  center;line-height:150%;word-break:keep-all'><span lang=EN-US
  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=sumStatisSCExam%><o:p></o:p></span></p>
  </td>
 </tr>
</table>


<p class=MsoNormal align=left style='margin-left:15.0pt;text-align:left;
text-indent:-15.0pt;mso-char-indent-count:-1.5'><span lang=EN-US
style='mso-bidi-font-size:10.0pt;line-height:107%;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p>

<p class=MsoNormal align=left style='margin-left:15.0pt;text-align:left;
text-indent:-15.0pt;mso-char-indent-count:-1.5'><span lang=EN-US
style='mso-bidi-font-size:10.0pt;line-height:107%;font-family:"Times New Roman",serif'>¡Ø
Number of applications is based on the filing date in Korea<br>
- <%=appCustRef(0)%>: <%=appCustName(0)%><span class=GramE>.</span><br>
- <%=appCustRef(1)%>: <%=appCustName(1)%><o:p></o:p></span></p>

<p class=MsoNormal align=left style='text-align:left'><span lang=EN-US
style='mso-bidi-font-size:10.0pt;line-height:107%;font-family:"Times New Roman",serif'>¡Ø
IPT: Intellectual Property Tribunal<br>
¡Ø P.C.: Patent Court<br>
¡Ø S.C.: Supreme Court<o:p></o:p></span></p>


<span lang=EN-US style='font-size:10.0pt;line-height:107%;font-family:"Times New Roman",serif;
mso-fareast-font-family:"¸¼Àº °íµñ";mso-font-kerning:1.0pt;mso-ansi-language:EN-US;
mso-fareast-language:KO;mso-bidi-language:AR-SA'><br clear=all style='page-break-before:always'>
</span>


<% If DataCnt_1_1 > 0 Then %>
 
 <%
 i = 0
 Set Rs = oConn.Execute(Sql_1_1)
 If Rs.EOF Then
	Call sbTitle_1_1("A")
 Else
	 Do Until Rs.EOF
		OurRef = Rs.Fields(0)
		YourRef = Rs.Fields(1)
		RegDate = Rs.Fields(2)
		RegNo = Rs.Fields(3)
		AnnEndDate = Rs.Fields(4)
		AnnYear = Rs.Fields(5)

		If i = 0 Then
			Call sbTitle_1_1("A")
		End If
		
		If i > 0 And i Mod 30 = 0 Then 'ÆäÀÌÂ¡Ã³¸®
			Call sbTableClose_1_1()
			Call sbNextLine()
			Call sbTitle_1_1("B")
		End If
	%>
		<tr style='mso-yfti-irow:3;height:19.85pt'>
		  <td width=35 style='width:25.9pt;border:solid windowtext 1.0pt;border-top:
		  none;mso-border-top-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;
		  padding:0cm 5.4pt 0cm 5.4pt;height:19.85pt'>
		  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
		  center;line-height:150%;word-break:keep-all'><span lang=EN-US
		  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=i+1%><o:p></o:p></span></p>
		  </td>
		  <td width=158 style='width:118.35pt;border-top:none;border-left:none;
		  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
		  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
		  mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt;height:19.85pt'>
		  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
		  center;line-height:150%;word-break:keep-all'><span lang=EN-US
		  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=YourRef%><o:p></o:p></span></p>
		  </td>
		  <td width=112 style='width:83.65pt;border-top:none;border-left:none;
		  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
		  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
		  mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt;height:19.85pt'>
		  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
		  center;line-height:150%;word-break:keep-all'><span lang=EN-US
		  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=OurRef%><o:p></o:p></span></p>
		  </td>
		  <td width=116 style='width:87.2pt;border-top:none;border-left:none;
		  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
		  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
		  mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt;height:19.85pt'>
		  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
		  center;line-height:150%;word-break:keep-all'><span lang=EN-US
		  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=fnDateReplaceEnd(RegDate)%><o:p></o:p></span></p>
		  </td>
		  <td width=85 colspan=2 style='width:63.6pt;border-top:none;border-left:none;
		  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
		  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
		  mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt;height:19.85pt'>
		  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
		  center;line-height:150%;word-break:keep-all'><span lang=EN-US
		  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=RegNo%><o:p></o:p></span></p>
		  </td>
		  <td width=47 style='width:35.4pt;border-top:none;border-left:none;border-bottom:
		  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:
		  solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:
		  solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt;height:19.85pt'>
		  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
		  center;line-height:150%;word-break:keep-all'><span lang=EN-US
		  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=AnnYear%>th<o:p></o:p></span></p>
		  </td>
		  <td width=135 colspan=2 style='width:101.05pt;border-top:none;border-left:
		  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
		  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
		  mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt;height:19.85pt'>
		  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
		  center;line-height:150%;word-break:keep-all'><span lang=EN-US
		  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=fnDateReplaceEnd(AnnEndDate)%><o:p></o:p></span></p>
		  </td>
		 </tr>
	<%
		i = i + 1
		Rs.MoveNext
	 Loop
 End If
 Rs.Close
 Set Rs = Nothing
 %>
 
 <![if !supportMisalignedColumns]>
 <tr height=0>
  <td width=35 style='border:none'></td>
  <td width=158 style='border:none'></td>
  <td width=112 style='border:none'></td>
  <td width=116 style='border:none'></td>
  <td width=81 style='border:none'></td>
  <td width=3 style='border:none'></td>
  <td width=47 style='border:none'></td>
  <td width=133 style='border:none'></td>
  <td width=2 style='border:none'></td>
 </tr>
 <![endif]>
</table>

<p class=MsoNormal align=left style='text-align:left'><span lang=EN-US
style='mso-bidi-font-size:10.0pt;line-height:107%;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p>

<span lang=EN-US style='font-size:10.0pt;line-height:107%;font-family:"Times New Roman",serif;
mso-fareast-font-family:"¸¼Àº °íµñ";mso-font-kerning:1.0pt;mso-ansi-language:EN-US;
mso-fareast-language:KO;mso-bidi-language:AR-SA'><br clear=all
style='page-break-before:always'>
</span>

<% End If %>

<% If DataCnt_1_2 > 0 Then %>

<p class=MsoNormal align=center style='margin-bottom:0cm;margin-bottom:.0001pt;
text-align:center'><b style='mso-bidi-font-weight:normal'><span lang=EN-US
style='font-size:16.0pt;line-height:107%;font-family:"Times New Roman",serif'>1-2.
Registered Patents Not Administered By HANSUNG<o:p></o:p></span></b></p>

<p class=MsoNormal style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;
margin-left:40.0pt;margin-bottom:.0001pt;text-indent:50.0pt;mso-char-indent-count:
5.0'><b style='mso-bidi-font-weight:normal'><span lang=EN-US style='mso-bidi-font-size:
10.0pt;line-height:107%;font-family:"Times New Roman",serif'>Client: <%=CustName%><o:p></o:p></span></b></p>

<p class=MsoNormal align=center style='margin-bottom:0cm;margin-bottom:.0001pt;
text-align:center'><b style='mso-bidi-font-weight:normal'><span lang=EN-US
style='mso-bidi-font-size:10.0pt;line-height:107%;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></b></p>

<p class=MsoNormal align=center style='margin-bottom:0cm;margin-bottom:.0001pt;
text-align:center'><b style='mso-bidi-font-weight:normal'><span lang=EN-US
style='mso-bidi-font-size:10.0pt;line-height:107%;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></b></p>

<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width=0
 style='width:513.5pt;border-collapse:collapse;mso-yfti-tbllook:1184;
 mso-padding-alt:0cm 5.4pt 0cm 5.4pt;mso-border-insideh:.5pt solid windowtext'>
 <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;mso-yfti-lastrow:yes'>
  <td width=502 valign=top style='width:376.2pt;padding:0cm 0cm 0cm 0cm'>
  <p class=MsoNormal align=left style='margin-bottom:0cm;margin-bottom:.0001pt;
  text-align:left;line-height:12.0pt'><b style='mso-bidi-font-weight:normal'><span
  lang=EN-US style='mso-bidi-font-size:10.0pt;font-family:"Times New Roman",serif'>HANSUNG Intellectual Property<o:p></o:p></span></b></p>
  </td>
  <td width=183 valign=top style='width:137.3pt;padding:0cm 0cm 0cm 0cm'>
  <p class=MsoNormal align=right style='margin-bottom:0cm;margin-bottom:.0001pt;
  text-align:right;line-height:12.0pt'><b style='mso-bidi-font-weight:normal'><span
  lang=EN-US style='mso-bidi-font-size:10.0pt;font-family:"Times New Roman",serif'><%=TodayEng%><o:p></o:p></span></b></p>
  </td>
 </tr>
</table>

<!--s-->
<% If DataCnt_1_2_app > 0 Then %>

<p class=MsoNormal align=left style='margin-bottom:0cm;margin-bottom:.0001pt;
text-align:left'><span lang=EN-US style='mso-bidi-font-size:10.0pt;line-height:
107%;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p>

<p class=MsoNormal align=left style='margin-bottom:0cm;margin-bottom:.0001pt;
text-align:left'><b style='mso-bidi-font-weight:normal'><span lang=EN-US
style='mso-bidi-font-size:10.0pt;line-height:107%;font-family:"Times New Roman",serif'>Patent Application Nos.:<o:p></o:p></span></b></p>

<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
13.0pt;mso-line-height-rule:exactly;mso-layout-grid-align:none;word-break:keep-all'><b
style='mso-bidi-font-weight:normal'><span lang=EN-US style='mso-bidi-font-size:
10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></b></p>

<table class=MsoTableGrid border=0 cellspacing=0 cellpadding=0 width=0
 style='width:517.4pt;border-collapse:collapse;border:none;mso-yfti-tbllook:
 1184;mso-padding-alt:0cm 0cm 0cm 0cm;mso-border-insideh:none;mso-border-insidev:
 none'>
<%
Dim Comma, RecordCnt
i = 0
Set Rs = Server.CreateObject("ADODB.RECORDSET")
Rs.Open Sql_1_2_app, oConn, 3
RecordCnt = Rs.RecordCount
Do Until Rs.EOF
	RegNo = Rs.Fields(0)

	If RecordCnt = (i+1) Then
		Comma = ""
	Else
		Comma = ","	
	End If	
	
	If i = 0 Then
%>
		<tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'>
<%
	End If

	If i > 0 And i Mod 6 = 0 Then '´ÙÀ½¶óÀÎ
%>
		</tr>
		 <tr style='mso-yfti-irow:1'>
		  <td width=77 valign=top style='width:57.45pt;padding:0cm 0cm 0cm 0cm'>
		  <p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
		  14.0pt;mso-line-height-rule:exactly'><span
		  lang=EN-US style='mso-bidi-font-size:10.0pt;font-family:"Times New Roman",serif;
		  letter-spacing:.2pt'><%=RegNo & Comma%><o:p></o:p></span></p>
		  </td>
<%
	ElseIf i = 0 Then
%>
		<td width=77 valign=top style='width:57.45pt;padding:0cm 0cm 0cm 0cm'>
		  <p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
		  14.0pt;mso-line-height-rule:exactly'><span
		  lang=EN-US style='mso-bidi-font-size:10.0pt;font-family:"Times New Roman",serif;
		  letter-spacing:.2pt'><%=RegNo & Comma%><o:p></o:p></span></p>
		  </td>
<%
	Else		
%>
		<td width=77 valign=top style='width:57.45pt;padding:0cm 0cm 0cm 0cm'>
		  <p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
		  14.0pt;mso-line-height-rule:exactly'><span
		  lang=EN-US style='mso-bidi-font-size:10.0pt;font-family:"Times New Roman",serif;
		  letter-spacing:.2pt'><%=RegNo & Comma%><o:p></o:p></span></p>
		  </td>
<%
	End If

	i = i + 1	
	Rs.MoveNext
Loop
Rs.Close
Set Rs = Nothing

For i = RecordCnt+1 To 6
%>
	<td width=77 valign=top style='width:57.45pt;padding:0cm 0cm 0cm 0cm'>
	  <p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
	  14.0pt;mso-line-height-rule:exactly'><span
	  lang=EN-US style='mso-bidi-font-size:10.0pt;font-family:"Times New Roman",serif;
	  letter-spacing:.2pt'><o:p></o:p></span></p>
	  </td>
<%
Next
%>
</tr>
</table>

<p class=MsoNormal align=left style='margin-bottom:0cm;margin-bottom:.0001pt;
text-align:left'><span lang=EN-US style='mso-bidi-font-size:10.0pt;line-height:
107%;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p>

<p class=MsoNormal align=left style='margin-bottom:0cm;margin-bottom:.0001pt;
text-align:left'><span lang=EN-US
style='mso-bidi-font-size:10.0pt;line-height:107%;font-family:"Times New Roman",serif'>( note: the issue fee is paid, waiting for the registration number. )<o:p></o:p></span></p>

<% End If %>
<!--e-->

<p class=MsoNormal align=left style='margin-bottom:0cm;margin-bottom:.0001pt;
text-align:left'><span lang=EN-US style='mso-bidi-font-size:10.0pt;line-height:
107%;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p>

<p class=MsoNormal align=left style='margin-bottom:0cm;margin-bottom:.0001pt;
text-align:left'><b style='mso-bidi-font-weight:normal'><span lang=EN-US
style='mso-bidi-font-size:10.0pt;line-height:107%;font-family:"Times New Roman",serif'>Patent Nos.:<o:p></o:p></span></b></p>

<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
13.0pt;mso-line-height-rule:exactly;mso-layout-grid-align:none;word-break:keep-all'><b
style='mso-bidi-font-weight:normal'><span lang=EN-US style='mso-bidi-font-size:
10.0pt;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></b></p>

<table class=MsoTableGrid border=0 cellspacing=0 cellpadding=0 width=0
 style='width:517.4pt;border-collapse:collapse;border:none;mso-yfti-tbllook:
 1184;mso-padding-alt:0cm 0cm 0cm 0cm;mso-border-insideh:none;mso-border-insidev:
 none'>
<%
i = 0
Set Rs = Server.CreateObject("ADODB.RECORDSET")
Rs.Open Sql_1_2, oConn, 3
RecordCnt = Rs.RecordCount
Do Until Rs.EOF
	RegNo = Rs.Fields(0)

	If RecordCnt = (i+1) Then
		Comma = ""
	Else
		Comma = ","	
	End If	
	
	If i = 0 Then
%>
		<tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'>
<%
	End If

	If i > 0 And i Mod 9 = 0 Then '´ÙÀ½¶óÀÎ
%>
		</tr>
		 <tr style='mso-yfti-irow:1'>
		  <td width=77 valign=top style='width:57.45pt;padding:0cm 0cm 0cm 0cm'>
		  <p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
		  14.0pt;mso-line-height-rule:exactly'><span
		  lang=EN-US style='mso-bidi-font-size:10.0pt;font-family:"Times New Roman",serif;
		  letter-spacing:.2pt'><%=RegNo & Comma%><o:p></o:p></span></p>
		  </td>
<%
	ElseIf i = 0 Then
%>
		<td width=77 valign=top style='width:57.45pt;padding:0cm 0cm 0cm 0cm'>
		  <p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
		  14.0pt;mso-line-height-rule:exactly'><span
		  lang=EN-US style='mso-bidi-font-size:10.0pt;font-family:"Times New Roman",serif;
		  letter-spacing:.2pt'><%=RegNo & Comma%><o:p></o:p></span></p>
		  </td>
<%
	Else		
%>
		<td width=77 valign=top style='width:57.45pt;padding:0cm 0cm 0cm 0cm'>
		  <p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
		  14.0pt;mso-line-height-rule:exactly'><span
		  lang=EN-US style='mso-bidi-font-size:10.0pt;font-family:"Times New Roman",serif;
		  letter-spacing:.2pt'><%=RegNo & Comma%><o:p></o:p></span></p>
		  </td>
<%
	End If

	i = i + 1	
	Rs.MoveNext
Loop
Rs.Close
Set Rs = Nothing
%>
</tr>
</table>

<span lang=EN-US style='font-size:10.0pt;line-height:107%;font-family:"Times New Roman",serif;
mso-fareast-font-family:"¸¼Àº °íµñ";letter-spacing:-.2pt;mso-font-kerning:1.0pt;
mso-ansi-language:EN-US;mso-fareast-language:KO;mso-bidi-language:AR-SA'><br clear=all style='page-break-before:always'>
</span>

<% End If %>


<% If DataCnt_2_1 > 0 Then %>

<%
PageCnt = 25 'ÆäÀÌÁö´ç Ãâ·Â¼ö
PageLoop = 0 'ÆäÀÌÁö´ç Ãâ·Â¹øÈ£
i = 0
Set Rs = oConn.Execute(Sql_2_1)
If Rs.EOF Then
	Call sbTitle_2_1("A")

Else

	Do Until Rs.EOF
		OurRef = Rs.Fields(0)
		YourRef = Rs.Fields(1)
		ApplNo = Rs.Fields(2)
		AbanDate = Rs.Fields(3)
		AbanDeDate = Rs.Fields(4)
		AbanMethod = Rs.Fields(5)

		If i = 0 Then
			Call sbTitle_2_1("A")
		End If
		
		If i > 0 And PageLoop Mod PageCnt = 0 Then 'ÆäÀÌÂ¡Ã³¸®
			Call sbTableClose_2_1()
			Call sbNextLine()
			Call sbTitle_2_1("B")
			PageCnt = 27
			PageLoop = 0
		End If
%>
		<tr style='mso-yfti-irow:2;height:19.85pt'>
		  <td width=34 style='width:25.65pt;border:solid windowtext 1.0pt;border-top:
		  none;mso-border-top-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;
		  padding:0cm 5.4pt 0cm 5.4pt;height:19.85pt'>
		  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
		  center;line-height:150%;word-break:keep-all'><span lang=EN-US
		  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=i+1%><o:p></o:p></span></p>
		  </td>
		  <td width=116 style='width:87.75pt;border-top:none;border-left:none;
		  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
		  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
		  mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt;height:19.85pt'>
		  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
		  center;line-height:150%;word-break:keep-all'><span lang=EN-US
		  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=OurRef%><o:p></o:p></span></p>
		  </td>
		  <td width=132 style='width:99.25pt;border-top:none;border-left:none;
		  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
		  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
		  mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt;height:19.85pt'>
		  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
		  center;line-height:150%;word-break:keep-all'><span lang=EN-US
		  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=YourRef%><o:p></o:p></span></p>
		  </td>
		  <td width=88 style='width:66.3pt;border-top:none;border-left:none;border-bottom:
		  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:
		  solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:
		  solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt;height:19.85pt'>
		  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
		  center;line-height:150%;word-break:keep-all'><span lang=EN-US
		  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=ApplNo%><o:p></o:p></span></p>
		  </td>
		  <td width=112 colspan=2 style='width:84.5pt;border-top:none;border-left:none;
		  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
		  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
		  mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt;height:19.85pt'>
		  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
		  center;line-height:150%;word-break:keep-all'><span lang=EN-US
		  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=fnDateReplaceEnd(AbanDate)%><o:p></o:p></span></p>
		  </td>
		  <td width=82 style='width:59.75pt;border-top:none;border-left:none;
		  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
		  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
		  mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt;height:19.85pt'>
		  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
		  center;line-height:150%;word-break:keep-all'><span lang=EN-US
		  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=AbanMethod%><o:p></o:p></span></p>
		  </td>
		  <td width=122 colspan=2 style='width:91.95pt;border-top:none;border-left:
		  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
		  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
		  mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt;height:19.85pt'>
		  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
		  center;line-height:150%;word-break:keep-all'><span lang=EN-US
		  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=fnDateReplaceEnd(AbanDeDate)%><o:p></o:p></span></p>
		  </td>
		 </tr>
	<%	
		i = i + 1
		PageLoop = PageLoop + 1
		Rs.MoveNext
	Loop
End If
Rs.Close
Set Rs = Nothing
%>

<![if !supportMisalignedColumns]>
	 <tr height=0>
	  <td width=34 style='border:none'></td>
	  <td width=110 style='border:none'></td>
	  <td width=136 style='border:none'></td>
	  <td width=90 style='border:none'></td>
	  <td width=79 style='border:none'></td>
	  <td width=33 style='border:none'></td>
	  <td width=82 style='border:none'></td>
	  <td width=116 style='border:none'></td>
	  <td width=6 style='border:none'></td>
	 </tr>
	 <![endif]>
	</table>

	<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
	14.0pt;mso-line-height-rule:exactly'><b style='mso-bidi-font-weight:normal'><span
	lang=EN-US style='mso-bidi-font-size:10.0pt;font-family:"Times New Roman",serif;
	letter-spacing:-.2pt'>* Method of abandonment<o:p></o:p></span></b></p>

	<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;text-indent:
	4.8pt;mso-char-indent-count:.5;line-height:14.0pt;mso-line-height-rule:exactly'><b
	style='mso-bidi-font-weight:normal'><span lang=EN-US style='mso-bidi-font-size:
	10.0pt;font-family:"Times New Roman",serif;letter-spacing:-.2pt'>A: by not
	requesting examination<o:p></o:p></span></b></p>

	<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;text-indent:
	4.8pt;mso-char-indent-count:.5;line-height:14.0pt;mso-line-height-rule:exactly'><b
	style='mso-bidi-font-weight:normal'><span lang=EN-US style='mso-bidi-font-size:
	10.0pt;font-family:"Times New Roman",serif;letter-spacing:-.2pt'>B: by not
	responding to an Office Action<o:p></o:p></span></b></p>

	<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;text-indent:
	4.8pt;mso-char-indent-count:.5;line-height:14.0pt;mso-line-height-rule:exactly'><b
	style='mso-bidi-font-weight:normal'><span lang=EN-US style='mso-bidi-font-size:
	10.0pt;font-family:"Times New Roman",serif;letter-spacing:-.2pt'>C: by
	non-payment of official fees<o:p></o:p></span></b></p>

	<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;text-indent:
	4.8pt;mso-char-indent-count:.5;line-height:14.0pt;mso-line-height-rule:exactly'><b
	style='mso-bidi-font-weight:normal'><span lang=EN-US style='mso-bidi-font-size:
	10.0pt;font-family:"Times New Roman",serif;letter-spacing:-.2pt'>D: by filing a
	notice of abandonment<o:p></o:p></span></b></p>

	<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;text-indent:
	4.8pt;mso-char-indent-count:.5;line-height:14.0pt;mso-line-height-rule:exactly'><b
	style='mso-bidi-font-weight:normal'><span lang=EN-US style='mso-bidi-font-size:
	10.0pt;font-family:"Times New Roman",serif;letter-spacing:-.2pt'>E: by not
	filing an appeal to the IPT (Patent Court/ Supreme Court)<o:p></o:p></span></b></p>

	<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;text-indent:
	4.8pt;mso-char-indent-count:.5;line-height:14.0pt;mso-line-height-rule:exactly'><b
	style='mso-bidi-font-weight:normal'><span lang=EN-US style='mso-bidi-font-size:
	10.0pt;font-family:"Times New Roman",serif;letter-spacing:-.2pt'>F: etc.<o:p></o:p></span></b></p>


<b style='mso-bidi-font-weight:normal'><span lang=EN-US style='font-size:10.0pt;
line-height:107%;font-family:"Times New Roman",serif;mso-fareast-font-family:
"¸¼Àº °íµñ";letter-spacing:-.2pt;mso-font-kerning:1.0pt;mso-ansi-language:EN-US;
mso-fareast-language:KO;mso-bidi-language:AR-SA'><br clear=all
style='page-break-before:always'>
</span></b>

<% End If %>


<% If DataCnt_2_2 > 0 Then %>

<p class=MsoNormal align=center style='margin-bottom:0cm;margin-bottom:.0001pt;
text-align:center'><b style='mso-bidi-font-weight:normal'><span lang=EN-US
style='font-size:16.0pt;line-height:107%;font-family:"Times New Roman",serif'>2-2. Applications Abandoned by December 31, <%=nDate-2%><o:p></o:p></span></b></p>

<p class=MsoNormal style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;
margin-left:40.0pt;margin-bottom:.0001pt;text-indent:70.0pt;mso-char-indent-count:
7.0'><b style='mso-bidi-font-weight:normal'><span lang=EN-US style='mso-bidi-font-size:
10.0pt;line-height:107%;font-family:"Times New Roman",serif'>Client: <%=CustName%><o:p></o:p></span></b></p>

<p class=MsoNormal align=center style='margin-bottom:0cm;margin-bottom:.0001pt;
text-align:center'><b style='mso-bidi-font-weight:normal'><span lang=EN-US
style='mso-bidi-font-size:10.0pt;line-height:107%;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></b></p>

<p class=MsoNormal align=center style='margin-bottom:0cm;margin-bottom:.0001pt;
text-align:center'><b style='mso-bidi-font-weight:normal'><span lang=EN-US
style='mso-bidi-font-size:10.0pt;line-height:107%;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></b></p>

<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width=0
 style='width:513.5pt;border-collapse:collapse;mso-yfti-tbllook:1184;
 mso-padding-alt:0cm 5.4pt 0cm 5.4pt;mso-border-insideh:.5pt solid windowtext'>
 <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;mso-yfti-lastrow:yes'>
  <td width=502 valign=top style='width:376.2pt;padding:0cm 0cm 0cm 0cm'>
  <p class=MsoNormal align=left style='margin-bottom:0cm;margin-bottom:.0001pt;
  text-align:left;line-height:12.0pt'><b style='mso-bidi-font-weight:normal'><span
  lang=EN-US style='mso-bidi-font-size:10.0pt;font-family:"Times New Roman",serif'>HANSUNG Intellectual Property<o:p></o:p></span></b></p>
  </td>
  <td width=183 valign=top style='width:137.3pt;padding:0cm 0cm 0cm 0cm'>
  <p class=MsoNormal align=right style='margin-bottom:0cm;margin-bottom:.0001pt;
  text-align:right;line-height:12.0pt'><b style='mso-bidi-font-weight:normal'><span
  lang=EN-US style='mso-bidi-font-size:10.0pt;font-family:"Times New Roman",serif'><%=TodayEng%><o:p></o:p></span></b></p>
  </td>
 </tr>
</table>

<p class=MsoNormal align=left style='margin-bottom:0cm;margin-bottom:.0001pt;
text-align:left'><span lang=EN-US style='mso-bidi-font-size:10.0pt;line-height:
107%;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></p>

<p class=MsoNormal align=left style='margin-bottom:0cm;margin-bottom:.0001pt;
text-align:left'><b style='mso-bidi-font-weight:normal'><span lang=EN-US
style='mso-bidi-font-size:10.0pt;line-height:107%;font-family:"Times New Roman",serif'>Application Nos.:<o:p></o:p></span></b></p>

<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
13.0pt;mso-line-height-rule:exactly;mso-layout-grid-align:none;word-break:keep-all'><b><span
lang=EN-US style='mso-bidi-font-size:10.0pt;font-family:"Times-Bold",serif;
mso-bidi-font-family:Times-Bold;mso-font-kerning:0pt'><o:p>&nbsp;</o:p></span></b></p>

<table class=MsoTableGrid border=0 cellspacing=0 cellpadding=0 width=0
 style='width:517.15pt;border-collapse:collapse;border:none;mso-yfti-tbllook:
 1184;mso-padding-alt:0cm 0cm 0cm 0cm;mso-border-insideh:none;mso-border-insidev:
 none'>
 <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'>
<%
i = 0
Set Rs = Server.CreateObject("ADODB.RECORDSET")
Rs.Open Sql_2_2, oConn, 3
RecordCnt = Rs.RecordCount
Do Until Rs.EOF
	ApplNo = Rs.Fields(0)

	If RecordCnt = (i+1) Then
		Comma = ""
	Else
		Comma = ","	
	End If

	If i > 0 And i Mod 8 = 0 Then '´ÙÀ½¶óÀÎ
%>
		</tr>
		 <tr style='mso-yfti-irow:1'>		 
<%
	End If
%>
	<td width=86 valign=top style='width:64.6pt;padding:0cm 0cm 0cm 0cm'>
	  <p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
	  13.0pt;mso-line-height-rule:exactly;mso-layout-grid-align:none;word-break:
	  keep-all'><span lang=EN-US style='mso-bidi-font-size:10.0pt;font-family:
	  "Times-Bold",serif;mso-bidi-font-family:Times-Bold;mso-font-kerning:0pt'><%=ApplNo & Comma%><o:p></o:p></span></p>
	  </td>
<%

	i = i + 1
	Rs.MoveNext
Loop
Rs.Close
Set Rs = Nothing
%>
 </tr>
</table>


<b><span lang=EN-US style='font-size:10.0pt;line-height:107%;font-family:"Times-Bold",serif;
mso-fareast-font-family:"¸¼Àº °íµñ";mso-bidi-font-family:Times-Bold;mso-ansi-language:
EN-US;mso-fareast-language:KO;mso-bidi-language:AR-SA'><br clear=all style='page-break-before:always'>
</span></b>

<% End If %>

<% If DataCnt_3 > 0 Then %>

<%
i = 0
PageCnt = 21 'ÆäÀÌÁö´ç Ãâ·Â¼ö
PageLoop = 0 'ÆäÀÌÁö´ç Ãâ·Â¹øÈ£
Set Rs = oConn.Execute(Sql_3)
If Rs.EOF Then
	Call sbTitle_3("A")

Else

	Do Until Rs.EOF
		YourRef = Rs.Fields(0)
		OurRef = Rs.Fields(1)
		ApplNo = Rs.Fields(2)
		FilingDate = Rs.Fields(3)
		PresentStatus = Rs.Fields(4)

		If i = 0 Then
			Call sbTitle_3("A")
		End If
		
		If i > 0 And PageLoop Mod PageCnt = 0 Then 'ÆäÀÌÂ¡Ã³¸®
			Call sbTableClose_3()
			Call sbNextLine()			
			Call sbTitle_3("B")
			PageCnt = 23
			PageLoop = 0
		End If
%>
		<tr style='mso-yfti-irow:2;height:27.7pt'>
		  <td width=34 style='width:25.55pt;border:solid windowtext 1.0pt;border-top:
		  none;mso-border-top-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;
		  padding:0cm 5.4pt 0cm 5.4pt;height:27.7pt'>
		  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
		  center;line-height:150%;word-break:keep-all'><span lang=EN-US
		  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=i+1%><o:p></o:p></span></p>
		  </td>
		  <td width=127 style='width:95.05pt;border-top:none;border-left:none;
		  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
		  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
		  mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt;height:27.7pt'>
		  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
		  center;line-height:150%;word-break:keep-all'><span lang=EN-US
		  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=YourRef%><o:p></o:p></span></p>
		  </td>
		  <td width=107 style='width:80.25pt;border-top:none;border-left:none;
		  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
		  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
		  mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt;height:27.7pt'>
		  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
		  center;line-height:150%;word-break:keep-all'><span lang=EN-US
		  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=OurRef%><o:p></o:p></span></p>
		  </td>
		  <td width=104 style='width:78.0pt;border-top:none;border-left:none;
		  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
		  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
		  mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt;height:27.7pt'>
		  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
		  center;line-height:150%;word-break:keep-all'><span lang=EN-US
		  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=ApplNo%><o:p></o:p></span></p>
		  </td>
		  <td width=101 colspan=2 style='width:75.55pt;border-top:none;border-left:
		  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
		  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
		  mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt;height:27.7pt'>
		  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
		  center;line-height:150%;word-break:keep-all'><span lang=EN-US
		  style='font-size:8.0pt;line-height:150%;font-family:"Times New Roman",serif'><%=fnDateReplaceEnd(FilingDate)%><o:p></o:p></span></p>
		  </td>
		  <td width=208 style='width:155.95pt;border-top:none;border-left:none;
		  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
		  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
		  mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt;height:27.7pt'>
		  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
		  center;line-height:8.0pt;mso-line-height-rule:exactly;word-break:keep-all'><span
		  lang=EN-US style='font-size:8.0pt;font-family:"Times New Roman",serif'><%=PresentStatus%><o:p></o:p></span></p>
		  </td>
		 </tr>
<%
		i = i + 1
		PageLoop = PageLoop + 1
		Rs.MoveNext
	Loop
End If
Rs.Close
Set Rs = Nothing
%>

<% End If %>


</div>

</body>

</html>

<%
oConn.Close
Set oConn = Nothing
'--------------------------------------------------------------------------------------------------------------------------
Private Sub sbArrayInitial(ArrayName) '¹è¿­ÃÊ±âÈ­
	Dim i
	For i = 0 To UBound(ArrayName)
		ArrayName(i) = 0
	Next	
End Sub
'--------------------------------------------------------------------------------------------------------------------------
Private Sub sbNextLine() '´ÙÀ½ÆäÀÌÁö
%>
	<br clear=all style='mso-special-character:line-break;page-break-before:always'>
	<p class=MsoNormal align=left style='margin-bottom:0cm;margin-bottom:.0001pt;
	text-align:left;line-height:normal;mso-pagination:widow-orphan;text-autospace:
	ideograph-numeric ideograph-other;word-break:keep-all'><b style='mso-bidi-font-weight:
	normal'><span lang=EN-US style='mso-bidi-font-size:10.0pt;font-family:"Times New Roman",serif;
	letter-spacing:-.2pt'><o:p>&nbsp;</o:p></span></b></p>
<%
End Sub
'--------------------------------------------------------------------------------------------------------------------------
Private Sub sbTableClose_1_1() 'Å×ÀÌºíÁ¾·á
%>
	<![if !supportMisalignedColumns]>
	 <tr height=0>
	  <td width=35 style='border:none'></td>
	  <td width=158 style='border:none'></td>
	  <td width=112 style='border:none'></td>
	  <td width=116 style='border:none'></td>
	  <td width=81 style='border:none'></td>
	  <td width=3 style='border:none'></td>
	  <td width=47 style='border:none'></td>
	  <td width=133 style='border:none'></td>
	  <td width=2 style='border:none'></td>
	 </tr>
	 <![endif]>
	</table>
<%
End Sub
'--------------------------------------------------------------------------------------------------------------------------
Private Sub sbTableClose_2_1() 'Å×ÀÌºíÁ¾·á
%>
	<![if !supportMisalignedColumns]>
	 <tr height=0>
	  <td width=34 style='border:none'></td>
	  <td width=116 style='border:none'></td>
	  <td width=132 style='border:none'></td>
	  <td width=88 style='border:none'></td>
	  <td width=79 style='border:none'></td>
	  <td width=33 style='border:none'></td>
	  <td width=82 style='border:none'></td>
	  <td width=116 style='border:none'></td>
	  <td width=6 style='border:none'></td>
	 </tr>
	 <![endif]>
	</table>

	<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;line-height:
	14.0pt;mso-line-height-rule:exactly'><b style='mso-bidi-font-weight:normal'><span
	lang=EN-US style='mso-bidi-font-size:10.0pt;font-family:"Times New Roman",serif;
	letter-spacing:-.2pt'>* Method of abandonment<o:p></o:p></span></b></p>

	<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;text-indent:
	4.9pt;mso-char-indent-count:.5;line-height:14.0pt;mso-line-height-rule:exactly'><b
	style='mso-bidi-font-weight:normal'><span lang=EN-US style='mso-bidi-font-size:
	10.0pt;font-family:"Times New Roman",serif;letter-spacing:-.2pt'>A: by not
	requesting examination<o:p></o:p></span></b></p>

	<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;text-indent:
	4.9pt;mso-char-indent-count:.5;line-height:14.0pt;mso-line-height-rule:exactly'><b
	style='mso-bidi-font-weight:normal'><span lang=EN-US style='mso-bidi-font-size:
	10.0pt;font-family:"Times New Roman",serif;letter-spacing:-.2pt'>B: by not
	responding to an Office Action<o:p></o:p></span></b></p>

	<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;text-indent:
	4.9pt;mso-char-indent-count:.5;line-height:14.0pt;mso-line-height-rule:exactly'><b
	style='mso-bidi-font-weight:normal'><span lang=EN-US style='mso-bidi-font-size:
	10.0pt;font-family:"Times New Roman",serif;letter-spacing:-.2pt'>C: by
	non-payment of official fees<o:p></o:p></span></b></p>

	<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;text-indent:
	4.9pt;mso-char-indent-count:.5;line-height:14.0pt;mso-line-height-rule:exactly'><b
	style='mso-bidi-font-weight:normal'><span lang=EN-US style='mso-bidi-font-size:
	10.0pt;font-family:"Times New Roman",serif;letter-spacing:-.2pt'>D: by filing a
	notice of abandonment<o:p></o:p></span></b></p>

	<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;text-indent:
	4.9pt;mso-char-indent-count:.5;line-height:14.0pt;mso-line-height-rule:exactly'><b
	style='mso-bidi-font-weight:normal'><span lang=EN-US style='mso-bidi-font-size:
	10.0pt;font-family:"Times New Roman",serif;letter-spacing:-.2pt'>E: by not
	filing an appeal to the IPT (Patent Court/ Supreme Court)<o:p></o:p></span></b></p>

	<p class=MsoNormal style='margin-bottom:0cm;margin-bottom:.0001pt;text-indent:
	4.9pt;mso-char-indent-count:.5;line-height:14.0pt;mso-line-height-rule:exactly'><b
	style='mso-bidi-font-weight:normal'><span lang=EN-US style='mso-bidi-font-size:
	10.0pt;font-family:"Times New Roman",serif;letter-spacing:-.2pt'>F: etc.<o:p></o:p></span></b></p>
<%
End Sub
'--------------------------------------------------------------------------------------------------------------------------
Private Sub sbTableClose_3() 'Å×ÀÌºíÁ¾·á
%>
	<![if !supportMisalignedColumns]>
	 <tr height=0>
	  <td width=34 style='border:none'></td>
	  <td width=127 style='border:none'></td>
	  <td width=107 style='border:none'></td>
	  <td width=104 style='border:none'></td>
	  <td width=80 style='border:none'></td>
	  <td width=21 style='border:none'></td>
	  <td width=208 style='border:none'></td>
	 </tr>
	 <![endif]>
	</table>
<%
End Sub
'--------------------------------------------------------------------------------------------------------------------------
Private Sub sbTitle_1_1(sType) '1-1 ¸®½ºÆ® Á¦¸ñ
	If sType = "A" Then
%>
		<p class=MsoNormal align=center style='margin-bottom:0cm;margin-bottom:.0001pt;
		text-align:center'><b style='mso-bidi-font-weight:normal'><span lang=EN-US
		style='font-size:16.0pt;line-height:107%;font-family:"Times New Roman",serif'>1-1. Registered Patents Administered By HANSUNG<o:p></o:p></span></b></p>

		<p class=MsoNormal style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;
		margin-left:80.0pt;margin-bottom:.0001pt;text-indent:26.3pt;mso-char-indent-count:
		2.63'><b style='mso-bidi-font-weight:normal'><span lang=EN-US style='mso-bidi-font-size:
		10.0pt;line-height:107%;font-family:"Times New Roman",serif'>Client: <%=CustName%><o:p></o:p></span></b></p>

		<p class=MsoNormal align=center style='margin-bottom:0cm;margin-bottom:.0001pt;
		text-align:center'><b style='mso-bidi-font-weight:normal'><span lang=EN-US
		style='mso-bidi-font-size:10.0pt;line-height:107%;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></b></p>

		<p class=MsoNormal align=center style='margin-bottom:0cm;margin-bottom:.0001pt;
		text-align:center'><b style='mso-bidi-font-weight:normal'><span lang=EN-US
		style='mso-bidi-font-size:10.0pt;line-height:107%;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></b></p>
<%
	End If
%>
	<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width=0
	 style='width:515.15pt;border-collapse:collapse;mso-yfti-tbllook:1184;
	 mso-padding-alt:0cm 5.4pt 0cm 5.4pt;mso-border-insideh:.5pt solid windowtext'>
	 <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;mso-row-margin-right:1.65pt'>
	  <td width=502 colspan=5 valign=top style='width:376.2pt;border:none;
	  border-bottom:solid windowtext 1.0pt;mso-border-bottom-alt:solid windowtext .5pt;
	  padding:0cm 0cm 0cm 0cm'>
	  <p class=MsoNormal align=left style='margin-bottom:0cm;margin-bottom:.0001pt;
	  text-align:left;line-height:12.0pt'><b style='mso-bidi-font-weight:normal'><span
	  lang=EN-US style='mso-bidi-font-size:10.0pt;font-family:"Times New Roman",serif'>HANSUNG Intellectual Property<o:p></o:p></span></b></p>
	  </td>
	  <td width=183 colspan=3 valign=top style='width:137.3pt;border:none;
	  border-bottom:solid windowtext 1.0pt;mso-border-bottom-alt:solid windowtext .5pt;
	  padding:0cm 0cm 0cm 0cm'>
	  <p class=MsoNormal align=right style='margin-bottom:0cm;margin-bottom:.0001pt;
	  text-align:right;line-height:12.0pt'><b style='mso-bidi-font-weight:normal'><span
	  lang=EN-US style='mso-bidi-font-size:10.0pt;font-family:"Times New Roman",serif'><%=TodayEng%><o:p></o:p></span></b></p>
	  </td>
	  <td style='mso-cell-special:placeholder;border:none;border-bottom:solid windowtext 1.0pt'
	  width=2><p class='MsoNormal'>&nbsp;</td>
	 </tr> 
	 <tr style='mso-yfti-irow:1'>
	  <td width=35 rowspan=2 style='width:25.9pt;border:solid windowtext 1.0pt;
	  border-top:none;mso-border-top-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;
	  padding:0cm 5.4pt 0cm 5.4pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:150%;word-break:keep-all'><b style='mso-bidi-font-weight:
	  normal'><span lang=EN-US style='font-size:8.0pt;line-height:150%;font-family:
	  "Times New Roman",serif'>No.<o:p></o:p></span></b></p>
	  </td>
	  <td width=158 rowspan=2 style='width:118.35pt;border-top:none;border-left:
	  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
	  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
	  mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:150%;word-break:keep-all'><b style='mso-bidi-font-weight:
	  normal'><span lang=EN-US style='font-size:8.0pt;line-height:150%;font-family:
	  "Times New Roman",serif'>Your Ref.<o:p></o:p></span></b></p>
	  </td>
	  <td width=112 rowspan=2 style='width:83.65pt;border-top:none;border-left:
	  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
	  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
	  mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:150%;word-break:keep-all'><b style='mso-bidi-font-weight:
	  normal'><span lang=EN-US style='font-size:8.0pt;line-height:150%;font-family:
	  "Times New Roman",serif'>Our Ref.<o:p></o:p></span></b></p>
	  </td>
	  <td width=116 rowspan=2 style='width:87.2pt;border-top:none;border-left:none;
	  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
	  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
	  mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:150%;word-break:keep-all'><b style='mso-bidi-font-weight:
	  normal'><span lang=EN-US style='font-size:8.0pt;line-height:150%;font-family:
	  "Times New Roman",serif'>Issue Date<o:p></o:p></span></b></p>
	  </td>
	  <td width=85 colspan=2 rowspan=2 style='width:63.6pt;border-top:none;
	  border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
	  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
	  mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:150%;word-break:keep-all'><b style='mso-bidi-font-weight:
	  normal'><span lang=EN-US style='font-size:8.0pt;line-height:150%;font-family:
	  "Times New Roman",serif'>Patent No.<o:p></o:p></span></b></p>
	  </td>
	  <td width=182 colspan=3 style='width:136.45pt;border-top:none;border-left:
	  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
	  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
	  mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:150%;word-break:keep-all'><b style='mso-bidi-font-weight:
	  normal'><span lang=EN-US style='font-size:8.0pt;line-height:150%;font-family:
	  "Times New Roman",serif'>Next Annuity<o:p></o:p></span></b></p>
	  </td>
	 </tr> 
	 <tr style='mso-yfti-irow:2'>
	  <td width=47 style='width:35.4pt;border-top:none;border-left:none;border-bottom:
	  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:
	  solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:
	  solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:150%;word-break:keep-all'><b style='mso-bidi-font-weight:
	  normal'><span lang=EN-US style='font-size:8.0pt;line-height:150%;font-family:
	  "Times New Roman",serif'>Year<o:p></o:p></span></b></p>
	  </td>
	  <td width=135 colspan=2 style='width:101.05pt;border-top:none;border-left:
	  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
	  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
	  mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:150%;word-break:keep-all'><b style='mso-bidi-font-weight:
	  normal'><span lang=EN-US style='font-size:8.0pt;line-height:150%;font-family:
	  "Times New Roman",serif'>Due Date<o:p></o:p></span></b></p>
	  </td>
	 </tr>
<%
End Sub
'--------------------------------------------------------------------------------------------------------------------------
Private Sub sbTitle_2_1(sType) '2-1 ¸®½ºÆ® Á¦¸ñ
	If sType = "A" Then
%>
		<p class=MsoNormal align=center style='margin-bottom:0cm;margin-bottom:.0001pt;
		text-align:center'><b style='mso-bidi-font-weight:normal'><span lang=EN-US
		style='font-size:14.0pt;line-height:106%;font-family:"Times New Roman",serif'>2-1.
		Applications Abandoned or Instructed to be abandoned in <%=nDate-1%> (<%=nDate%>)<o:p></o:p></span></b></p>

		<p class=MsoNormal style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;
		margin-left:35.0pt;margin-bottom:.0001pt;text-indent:25.0pt;mso-char-indent-count:
		2.5'><b style='mso-bidi-font-weight:normal'><span lang=EN-US style='mso-bidi-font-size:
		10.0pt;line-height:106%;font-family:"Times New Roman",serif'>Client: <%=CustName%><o:p></o:p></span></b></p>

		<p class=MsoNormal align=center style='margin-bottom:0cm;margin-bottom:.0001pt;
		text-align:center'><b style='mso-bidi-font-weight:normal'><span lang=EN-US
		style='mso-bidi-font-size:10.0pt;line-height:106%;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></b></p>

		<p class=MsoNormal align=center style='margin-bottom:0cm;margin-bottom:.0001pt;
		text-align:center'><b style='mso-bidi-font-weight:normal'><span lang=EN-US
		style='mso-bidi-font-size:10.0pt;line-height:106%;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></b></p>
<%
	End If
%>
	<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width=0
	 style='width:515.15pt;border-collapse:collapse;mso-yfti-tbllook:1184;
	 mso-padding-alt:0cm 5.4pt 0cm 5.4pt;mso-border-insideh:.5pt solid windowtext'>
	 <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;mso-row-margin-right:4.85pt'>
	  <td width=449 colspan=5 valign=top style='width:338.55pt;border:none;
	  border-bottom:solid windowtext 1.0pt;mso-border-bottom-alt:solid windowtext .5pt;
	  padding:0cm 0cm 0cm 0cm'>
	  <p class=MsoNormal align=left style='margin-bottom:0cm;margin-bottom:.0001pt;
	  text-align:left;line-height:12.0pt'><b style='mso-bidi-font-weight:normal'><span
	  lang=EN-US style='mso-bidi-font-size:10.0pt;font-family:"Times New Roman",serif'>HANSUNG Intellectual Property<o:p></o:p></span></b></p>
	  </td>
	  <td width=231 colspan=3 valign=top style='width:171.75pt;border:none;
	  border-bottom:solid windowtext 1.0pt;mso-border-bottom-alt:solid windowtext .5pt;
	  padding:0cm 0cm 0cm 0cm'>
	  <p class=MsoNormal align=right style='margin-bottom:0cm;margin-bottom:.0001pt;
	  text-align:right;line-height:12.0pt'><b style='mso-bidi-font-weight:normal'><span
	  lang=EN-US style='mso-bidi-font-size:10.0pt;font-family:"Times New Roman",serif'><%=TodayEng%><o:p></o:p></span></b></p>
	  </td>
	  <td style='mso-cell-special:placeholder;border:none;border-bottom:solid windowtext 1.0pt'
	  width=6><p class='MsoNormal'>&nbsp;</td>
	 </tr>
	 <tr style='mso-yfti-irow:1;height:25.65pt'>
	  <td width=34 style='width:25.65pt;border:solid windowtext 1.0pt;border-top:
	  none;mso-border-top-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;
	  padding:0cm 5.4pt 0cm 5.4pt;height:25.65pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:10.0pt;mso-line-height-rule:exactly;word-break:keep-all'><b
	  style='mso-bidi-font-weight:normal'><span lang=EN-US style='font-size:8.0pt;
	  font-family:"Times New Roman",serif'>No.<o:p></o:p></span></b></p>
	  </td>
	  <td width=116 style='width:87.75pt;border-top:none;border-left:none;
	  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
	  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
	  mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt;height:25.65pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:10.0pt;mso-line-height-rule:exactly;word-break:keep-all'><b
	  style='mso-bidi-font-weight:normal'><span lang=EN-US style='font-size:8.0pt;
	  font-family:"Times New Roman",serif'>Our Ref.<o:p></o:p></span></b></p>
	  </td>
	  <td width=132 style='width:99.25pt;border-top:none;border-left:none;
	  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
	  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
	  mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt;height:25.65pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:10.0pt;mso-line-height-rule:exactly;word-break:keep-all'><b
	  style='mso-bidi-font-weight:normal'><span lang=EN-US style='font-size:8.0pt;
	  font-family:"Times New Roman",serif'>Your Ref.<o:p></o:p></span></b></p>
	  </td>
	  <td width=88 style='width:66.3pt;border-top:none;border-left:none;border-bottom:
	  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:
	  solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:
	  solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt;height:25.65pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:10.0pt;mso-line-height-rule:exactly;word-break:keep-all'><b
	  style='mso-bidi-font-weight:normal'><span lang=EN-US style='font-size:8.0pt;
	  font-family:"Times New Roman",serif'>Korean Patent<br>
	  Application No.<o:p></o:p></span></b></p>
	  </td>
	  <td width=112 colspan=2 style='width:84.5pt;border-top:none;border-left:none;
	  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
	  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
	  mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt;height:25.65pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:10.0pt;mso-line-height-rule:exactly;word-break:keep-all'><b
	  style='mso-bidi-font-weight:normal'><span lang=EN-US style='font-size:8.0pt;
	  font-family:"Times New Roman",serif'>Date of Instruction<br>
	  to Abandon<o:p></o:p></span></b></p>
	  </td>
	  <td width=82 style='width:59.75pt;border-top:none;border-left:none;
	  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
	  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
	  mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt;height:25.65pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:10.0pt;mso-line-height-rule:exactly;word-break:keep-all'><b
	  style='mso-bidi-font-weight:normal'><span lang=EN-US style='font-size:8.0pt;
	  font-family:"Times New Roman",serif'>Method of<br>
	  Abandonment<o:p></o:p></span></b></p>
	  </td>
	  <td width=122 colspan=2 style='width:91.95pt;border-top:none;border-left:
	  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
	  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
	  mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt;height:25.65pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:10.0pt;mso-line-height-rule:exactly;word-break:keep-all'><b
	  style='mso-bidi-font-weight:normal'><span lang=EN-US style='font-size:8.0pt;
	  font-family:"Times New Roman",serif'>Date of<br>
	  Abandonment<o:p></o:p></span></b></p>
	  </td>
	 </tr>
<%
End Sub
'--------------------------------------------------------------------------------------------------------------------------
Private Sub sbTitle_3(sType) '3 ¸®½ºÆ® Á¦¸ñ
	If sType = "A" Then
%>
		<p class=MsoNormal align=center style='margin-bottom:0cm;margin-bottom:.0001pt;
		text-align:center'><b style='mso-bidi-font-weight:normal'><span lang=EN-US
		style='font-size:14.0pt;line-height:106%;font-family:"Times New Roman",serif'>3.
		</span></b><b><span lang=EN-US style='font-size:14.5pt;line-height:106%;
		font-family:"Times-Bold",serif;mso-bidi-font-family:Times-Bold;mso-font-kerning:
		0pt'>Pending Applications</span></b><b style='mso-bidi-font-weight:normal'><span
		lang=EN-US style='font-size:14.0pt;line-height:106%;font-family:"Times New Roman",serif'><o:p></o:p></span></b></p>

		<p class=MsoNormal style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;
		margin-left:175.0pt;margin-bottom:.0001pt;text-indent:20.0pt;mso-char-indent-count:
		2.0'><b style='mso-bidi-font-weight:normal'><span lang=EN-US style='mso-bidi-font-size:
		10.0pt;line-height:106%;font-family:"Times New Roman",serif'>Client: <%=CustName%><o:p></o:p></span></b></p>

		<p class=MsoNormal align=center style='margin-bottom:0cm;margin-bottom:.0001pt;
		text-align:center'><b style='mso-bidi-font-weight:normal'><span lang=EN-US
		style='mso-bidi-font-size:10.0pt;line-height:106%;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></b></p>

		<p class=MsoNormal align=center style='margin-bottom:0cm;margin-bottom:.0001pt;
		text-align:center'><b style='mso-bidi-font-weight:normal'><span lang=EN-US
		style='mso-bidi-font-size:10.0pt;line-height:106%;font-family:"Times New Roman",serif'><o:p>&nbsp;</o:p></span></b></p>
<%
	End If
%>
	<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width=0
	 style='width:510.35pt;border-collapse:collapse;mso-yfti-tbllook:1184;
	 mso-padding-alt:0cm 5.4pt 0cm 5.4pt;mso-border-insideh:.5pt solid windowtext'>
	 <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'>
	  <td width=451 colspan=5 valign=top style='width:338.55pt;border:none;
	  border-bottom:solid windowtext 1.0pt;mso-border-bottom-alt:solid windowtext .5pt;
	  padding:0cm 0cm 0cm 0cm'>
	  <p class=MsoNormal align=left style='margin-bottom:0cm;margin-bottom:.0001pt;
	  text-align:left;line-height:12.0pt'><b style='mso-bidi-font-weight:normal'><span
	  lang=EN-US style='mso-bidi-font-size:10.0pt;font-family:"Times New Roman",serif'>HANSUNG Intellectual Property<o:p></o:p></span></b></p>
	  </td>
	  <td width=229 colspan=2 valign=top style='width:171.8pt;border:none;
	  border-bottom:solid windowtext 1.0pt;mso-border-bottom-alt:solid windowtext .5pt;
	  padding:0cm 0cm 0cm 0cm'>
	  <p class=MsoNormal align=right style='margin-bottom:0cm;margin-bottom:.0001pt;
	  text-align:right;line-height:12.0pt'><b style='mso-bidi-font-weight:normal'><span
	  lang=EN-US style='mso-bidi-font-size:10.0pt;font-family:"Times New Roman",serif'><%=TodayEng%><o:p></o:p></span></b></p>
	  </td>
	 </tr>
	 <tr style='mso-yfti-irow:1;height:25.65pt'>
	  <td width=34 style='width:25.55pt;border:solid windowtext 1.0pt;border-top:
	  none;mso-border-top-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;
	  padding:0cm 5.4pt 0cm 5.4pt;height:25.65pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:10.0pt;mso-line-height-rule:exactly;word-break:keep-all'><b
	  style='mso-bidi-font-weight:normal'><span lang=EN-US style='font-size:8.0pt;
	  font-family:"Times New Roman",serif'>No.<o:p></o:p></span></b></p>
	  </td>
	  <td width=127 style='width:95.05pt;border-top:none;border-left:none;
	  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
	  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
	  mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt;height:25.65pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:10.0pt;mso-line-height-rule:exactly;word-break:keep-all'><b
	  style='mso-bidi-font-weight:normal'><span lang=EN-US style='font-size:8.0pt;
	  font-family:"Times New Roman",serif'>Your Ref.<o:p></o:p></span></b></p>
	  </td>
	  <td width=107 style='width:80.25pt;border-top:none;border-left:none;
	  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
	  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
	  mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt;height:25.65pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:10.0pt;mso-line-height-rule:exactly;word-break:keep-all'><b
	  style='mso-bidi-font-weight:normal'><span lang=EN-US style='font-size:8.0pt;
	  font-family:"Times New Roman",serif'>Our Ref.<o:p></o:p></span></b></p>
	  </td>
	  <td width=104 style='width:78.0pt;border-top:none;border-left:none;
	  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
	  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
	  mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt;height:25.65pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:10.0pt;mso-line-height-rule:exactly;word-break:keep-all'><b
	  style='mso-bidi-font-weight:normal'><span lang=EN-US style='font-size:8.0pt;
	  font-family:"Times New Roman",serif'>Korean Patent<br>
	  Application No.<o:p></o:p></span></b></p>
	  </td>
	  <td width=101 colspan=2 style='width:75.55pt;border-top:none;border-left:
	  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
	  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
	  mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt;height:25.65pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:10.0pt;mso-line-height-rule:exactly;word-break:keep-all'><b
	  style='mso-bidi-font-weight:normal'><span lang=EN-US style='font-size:8.0pt;
	  font-family:"Times New Roman",serif'>Application Filing<br>
	  Date<o:p></o:p></span></b></p>
	  </td>
	  <td width=208 style='width:155.95pt;border-top:none;border-left:none;
	  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
	  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
	  mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt;height:25.65pt'>
	  <p class=MsoNormal align=center style='mso-margin-bottom-alt:auto;text-align:
	  center;line-height:10.0pt;mso-line-height-rule:exactly;word-break:keep-all'><b
	  style='mso-bidi-font-weight:normal'><span lang=EN-US style='font-size:8.0pt;
	  font-family:"Times New Roman",serif'>Present<br>
	  Status<o:p></o:p></span></b></p>
	  </td>
	 </tr>
<%
End Sub
'--------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------
%>
