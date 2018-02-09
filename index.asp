<!-- #include file="include/dbcon.asp" -->
<!-- #include file="include/Session_chk.asp" -->
<!-- #include file="include/function.asp" -->
<%
Const BUTTON_Style = "FONT-SIZE: 9pt; border-width:1px;border-color:#666600;border-style:solid;background-color:#F3EFE2; padding-top:2px;cursor:hand;width:90px;"
Const t_color5 = "#ffffff" 
Const t_color2 = "#EEF3FB"
Const f_style = "FONT-SIZE: 8pt; FONT-FAMILY: '돋움', '굴림'; border: 1px solid #AEAFBB"
Const main_right_C_color = "#C4D2E9"
Dim Sql, Rs, i

%>
<html>
<head>
<script type="text/javascript">

function ButtonDisplay(id) //함수_실적보고표 다운로드 버튼
{
	document.getElementById(id).disabled = false;
	document.getElementById(id).value = " 다운로드 ";
}

function ExchangeDisplay(n) //함수_환율정보 팝업
{
	if (document.getElementById(n).style.display == "none")
	{	document.getElementById(n).style.display = "block";	}
	else
	{	document.getElementById(n).style.display = "none";	}
}

function Sform2Submit()  //함수_실적보고표 종류/기간선택
{
	if (Sform2.sPaper.value == "")
	{
		alert("종류를 선택하여 주세요.");
		return;
	}
	if (Sform2.sKind.value == "")
	{
		alert("기간구분을 선택하여 주세요.");
		return;
	}
	Sform2.submit();	
}

function sPaperChanged(v) //함수_실적보고표 서류선택
{
	cdate = new Date();

	if (v == "A")
	{		
		document.getElementById("dv2MemberID").style.display = "none";
		document.getElementById("Sform2").action = "result_part_exam.asp";
	}
	else if (v == "B")
	{
		document.getElementById("dv2MemberID").style.display = "inline";
		document.getElementById("Sform2").action = "result_part_exam.asp";
	}
	else if (v == "C")
	{
		document.getElementById("dv2MemberID").style.display = "inline";
		document.getElementById("Sform2").action = "result_person_exam.asp";
	}
	else if (v == "D")
	{
		document.getElementById("dv2MemberID").style.display = "none";
		document.getElementById("Sform2").action = "result_score_exam.asp";
		Sform2.StartYear.value = cdate.getFullYear(); // 검색기간 초기화(올해3월~내년2월)
		Sform2.StartMonth.value = 3;
		Sform2.EndYear.value = cdate.getFullYear()+1;
		Sform2.EndMonth.value = 2;
	}
}

function MonthChk(t)  //함수_실적보고표 검색월 입력오류메시지
{
	if (t == "F")
	{
		if (document.getElementById("sPaper").value == "D"&& Sform2.StartMonth.value < 3)
		{
			alert("검색기간 시작월에 2월 이전은 선택할 수 없습니다.");
			Sform2.StartMonth.value = 3;		
		}
	}
}

function Sform4Submit()  //함수_입금수수료 다운로드버튼
{	Sform4.submit();}

function Sform1Submit() //함수_미수금집계표 다운로드버튼
{	document.getElementById().action = "result_unpaidlist.asp"; }

function Sform3Submit() //함수_특허집계표 입력오류메시지
{	
	if (document.getElementById("CustNum").value == "")
	{
		alert("의뢰인을 선택하여 주세요.");
		return;
	}
	document.getElementById("Sform3").action = "result_patent.asp";
	Sform3.submit();	
}

function Sform3Search() //함수_특허집계표 의뢰인코드 입력오류메시지
{
	document.getElementById("Sform3").action = "index.asp";
	if (document.getElementById("CustRef").value == "")
	{
		alert("검색어를 입력하여 주세요.");
		return false;
	}
	return true;
}



</script>

<style type="text/css">
.auto-style1 {
	text-align: right;
	font-Size: 9pt;
}
.auto-style2 {
	text-align: center;
	font-size: large;
}
</style>
</head>

<!------ 본문 START ------>

<body text="#000000" bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<p>&nbsp;</p>

<table border="0">
	<tr>
		<td>
			<table border="0" cellpadding="0" style="padding: 4px; height: 46px;" width="1000" cellspacing="0">
			<tr>
				<td height="20" class="auto-style2"><strong>특허법인 한성</strong></td>
			</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td class="auto-style1"><a href="logout.asp">로그아웃</a> </td>
	</tr>

<!-- 실적보고표 START -->
	<tr>
		<td>
			<form id="Sform2" name="Sform2" method="post" style="margin:0;">
			<table width="980"  border="0" cellpadding="5" cellspacing="1" bgcolor="<%=main_right_C_color%>" align="right">
			<tr align="center">
				<td bgcolor="<%=t_color5%>" align="left"  height="25" width="13%"><span style="font-size:10pt;color:#0080FF;font-weight:bold; margin:0;">실적보고표</span></td>
				<td bgcolor="<%=t_color5%>" align="left"  height="25" width="67%">
					<select id="sPaper" name="sPaper" onchange="javascript:ButtonDisplay('btnSubmit2');sPaperChanged(this.value);">
						<option value="">--- 종류 ---</option>
						<option value="A">부서별실적 보고표</option>
						<option value="B">개인별실적 보고표</option>
						<option value="C">청구실적 보고표</option>
						<option value="D">누적실적 보고표</option>
					</select>&nbsp;
					<select id="dv2MemberID" name="sMemberID" onchange="javascript:ButtonDisplay('btnSubmit2');" style="display:none;">
					<%
					Sql = "SELECT ID, mName FROM Member WHERE ID <> 'sopi' and ID <> 'default' ORDER BY mName "
					Set Rs = oConn.Execute(Sql)
					Do Until Rs.EOF
					%>
						<option value="<%=Rs.Fields(0)%>"><%=Rs.Fields(1)%></option>
					<%
						Rs.MoveNext
					Loop
					Rs.Close
					Set Rs = Nothing
					%>
					</select>&nbsp;
					<select name="sKind" onchange="javascript:ButtonDisplay('btnSubmit2');">
						<option value="">- 기간구분 -</option>
						<option value="A" selected>청구일</option>
					</select>
					<select name="StartYear" onchange="javascript:ButtonDisplay('btnSubmit2');">
					<%
					For i = Year(Date)-3 To Year(Date)+1
					%>
						<option value="<%=i%>" <%If Year(Date) = i Then Response.Write "selected"%>><%=i%>년</option>
					<%
					Next
					%>			
					</select>
					<select name="StartMonth" onchange="javascript:ButtonDisplay('btnSubmit2');MonthChk('F');">
					<%
					For i = 1 To 12
					%>
						<option value="<%=i%>"><%=i%>월</option>
					<%
					Next
					%>			
					</select> ~
					<select name="EndYear" onchange="javascript:ButtonDisplay('btnSubmit2');">
					<%
					For i = Year(Date)-3 To Year(Date)+1
					%>
						<option value="<%=i%>" <%If Year(Date) = i Then Response.Write "selected"%>><%=i%>년</option>
					<%
					Next
					%>			
					</select>
					<select name="EndMonth" onchange="javascript:ButtonDisplay('btnSubmit2');MonthChk('L');">
					<%
					For i = 1 To 12
					%>
						<option value="<%=i%>"><%=i%>월</option>
					<%
					Next
					%>			
					</select>
				</td>
				<td rowspan="3" bgcolor="<%=t_color2%>"  align="left" height="25" style="width: 20%">
					<input id="btnSubmit2" type="button" value=" 다운로드 " onclick="Sform2Submit();" style="<%=BUTTON_Style%>">
					<input id="btnSubmit2_2" type="button" value=" 환율정보 " onclick="ExchangeDisplay('dvExchange2');" style="<%=BUTTON_Style%>"> 
				</td>
			</tr>
			</table>
			</form>

<!-- 환율정보 START -->
			<div id="dvExchange2" style="display:none;position:absolute; left: 455px">
				<table width="300" border="0" cellpadding="2" cellspacing="1" bgcolor="#CFC4E9" align="center" style="font-size:10pt" >	
					<tr>
						<td bgcolor="#F2EEFB" align="center" height="25" width="80">통 화</td>
						<td bgcolor="#F2EEFB" align="center" height="25" width="80">적용일</td>
						<td bgcolor="#F2EEFB" align="center" height="25" width="80">환 율</td>
					</tr></span>
					<%
					Dim LineColor
					Sql = "SELECT CurrencyType, StartDate, Exchange FROM ExchangeRatePeriod ORDER BY CurrencyType, StartDate"
					Set Rs = oConn.Execute(Sql)
					Do Until Rs.EOF
						If Rs("CurrencyType") = "$" Then
							LineColor = "#E8E8E8"
						Else
							LineColor = "#FFFFFF"
						End If
					%>
						<tr>
							<td bgcolor="<%=LineColor%>" align="center" height="25">
								<%=Rs("CurrencyType")%>
							</td>
							<td bgcolor="<%=LineColor%>" align="center" height="25">
								<%=Rs("StartDate")%>
							</td>
							<td bgcolor="<%=LineColor%>" align="right" height="25">
								<%=fnMoneyType(Rs("Exchange"))%>
							</td>
						</tr>
					<%
						Rs.MoveNext
					Loop
					%>				
				</table>
			</div>
		</td>
	</tr>
			
<!-- 입금수수료 START -->		
	<tr>
		<td>
			<form id="Sform4" name="Sform4" action="payfee_list.asp" method="post" style="margin:0;">
			<table width="980"  border="0" cellpadding="5" cellspacing="1" bgcolor="<%=main_right_C_color%>" align="right">
			<tr align="center">
				<td bgcolor="<%=t_color5%>" align="left"  height="25" width="13%"><span style="font-size:10pt;color:#0080FF;font-weight:bold; margin:0;">입금수수료</span></td>
				<td bgcolor="<%=t_color5%>" align="left"  height="25" width="67%">
					<select name="sKind" onchange="javascript:ButtonDisplay('btnSubmit4');">
						<option value="">--- 기간구분 ---</option>
						<option value="A" selected>청구일</option>
					</select>&nbsp;
					<select name="StartYear" onchange="javascript:ButtonDisplay('btnSubmit4');">
					<%
					For i = Year(Date)-3 To Year(Date)
					%>
						<option value="<%=i%>" <%If Year(Date) = i Then Response.Write "selected"%>><%=i%>년</option>
					<%
					Next
					%>			
					</select>
					<select name="StartMonth" onchange="javascript:ButtonDisplay('btnSubmit4');">
					<%
					For i = 1 To 12
					%>
						<option value="<%=i%>"><%=i%>월</option>
					<%
					Next
					%>			
					</select> ~
					<select name="EndYear" onchange="javascript:ButtonDisplay('btnSubmit4');">
					<%
					For i = Year(Date)-3 To Year(Date)+1
					%>
						<option value="<%=i%>" <%If Year(Date) = i Then Response.Write "selected"%>><%=i%>년</option>
					<%
					Next
					%>			
					</select>
					<select name="EndMonth" onchange="javascript:ButtonDisplay('btnSubmit4');">
					<%
					For i = 1 To 12
					%>
						<option value="<%=i%>"><%=i%>월</option>
					<%
					Next
					%>			
					</select>
					&nbsp;<br>
					<select name="sKind2" onchange="javascript:ButtonDisplay('btnSubmit4');">
						<option value="">--- 기간구분 ---</option>
						<option value="A" selected>입금일</option>
					</select>&nbsp;
					<select name="StartYear2" onchange="javascript:ButtonDisplay('btnSubmit4');">
					<%
					For i = Year(Date)-4 To Year(Date)+1
					%>
						<option value="<%=i%>" <%If Year(Date) = i Then Response.Write "selected"%>><%=i%>년</option>
					<%
					Next
					%>			
					</select>
					<select name="StartMonth2" onchange="javascript:ButtonDisplay('btnSubmit4');">
					<%
					For i = 1 To 12
					%>
						<option value="<%=i%>"><%=i%>월</option>
					<%
					Next
					%>			
					</select> ~
					<select name="EndYear2" onchange="javascript:ButtonDisplay('btnSubmit4');">
					<%
					For i = Year(Date)-4 To Year(Date)
					%>
						<option value="<%=i%>" <%If Year(Date) = i Then Response.Write "selected"%>><%=i%>년</option>
					<%
					Next
					%>			
					</select>
					<select name="EndMonth2" onchange="javascript:ButtonDisplay('btnSubmit4');">
					<%
					For i = 1 To 12
					%>
						<option value="<%=i%>"><%=i%>월</option>
					<%
					Next
					%>			
					</select>
					
				</td>
				<td rowspan="3" bgcolor="<%=t_color2%>"  align="left" height="25" width="22%">
					<input id="btnSubmit4" type="button" value=" 다운로드 " onclick="Sform4Submit();" style="<%=BUTTON_Style%>">&nbsp;&nbsp;
				</td>
			</tr>
			</table>
			</form>

<!-- 미수금집계 START -->
	<tr>
		<td>
			<table width="980"  border="0" cellpadding="5" cellspacing="1" bgcolor="<%=main_right_C_color%>" align="right" style="margin:0; height: 5px;">
			<tr align="center">
				<td bgcolor="<%=t_color5%>" align="left"  height="25" width="13%"><span style="font-size:10pt;color:#0080FF;font-weight:bold; margin:0;">미수금 집계표</span></td>
				<td bgcolor="<%=t_color5%>" align="left"  height="25" width="67%">
					<input id="btnSubmit1" type="button" value=" 미수금 관리 " onclick="location.href='manage_invoice.asp';" style="<%=BUTTON_Style%>">
				</td>
				<td rowspan="3" bgcolor="<%=t_color2%>"  align="left" height="25" style="width: 22%"></td>
			</tr>
			</table>
		</td>
	</tr>

<!-- 특허집계표 START -->
<%
Dim CustRef
CustRef = Request("CustRef")
%>
	<tr>
		<td>
			<form id="Sform3" name="Sform3" action="index.asp" method="post" onsubmit="return Sform3Search();" style="margin:0;">
			<table width="980"  border="0" cellpadding="5" cellspacing="1" bgcolor="<%=main_right_C_color%>" align="right">
			<tr align="center">
				<td bgcolor="<%=t_color5%>" align="left"  height="25" width="13%"><span style="font-size:10pt;color:#0080FF;font-weight:bold; margin:0;">특허집계표</span></td>
				<td bgcolor="<%=t_color5%>" align="left"  height="25" width="67%"><span style="font-size:10pt;font-weight:margin:0;">의뢰인코드 : 
					<input type="text" id="CustRef" name="CustRef" value="<%=CustRef%>" size="6" onclick="javascript:ButtonDisplay('btnSubmit3');">
					<input type="submit" value=" 검 색 " >&nbsp;&nbsp;
					<select id="CustNum" name="CustNum">
					<option value="">--- 의뢰인 선택 ---</option>
					<%
					If CustRef <> "" Then
						Sql = "SELECT Num, Field41, mName FROM Customer WHERE Field41 LIKE'%" & CustRef & "%' "
						Set Rs = oConn.Execute(Sql)
						Do Until Rs.EOF
					%>
							<option value="<%=Rs.Fields(0)%>">[<%=Rs.Fields(1) & "] " & Rs.Fields(2)%></option>
					<%
							Rs.MoveNext
						Loop
						Rs.Close
						Set Rs = Nothing
					End If
					%>						
					</select>
				</td>
				<td rowspan="3" bgcolor="<%=t_color2%>"  align="left" height="25" style="width:22%">
					<input id="btnSubmit3" type="button" value=" 다운로드 " onclick="Sform3Submit();" style="<%=BUTTON_Style%>">&nbsp;&nbsp;
				</td>
			</tr>
			</table>
			</form>	
		</td>
	</tr>


<!-- 월간 보고서식 START -->
	<tr>
		<td>
			<table width="980"  border="0" cellpadding="5" cellspacing="1" bgcolor="<%=main_right_C_color%>" align="right" style="margin:0; height: 5px;">
			<tr align="center">
				<td bgcolor="<%=t_color5%>" align="left"  height="25" width="13%"><span style="font-size:10pt;color:#0080FF;font-weight:bold; margin:0;">월간 보고서식</span></td>
				<td bgcolor="<%=t_color5%>" align="left"  height="25" width="67%">
					<input type="button" value="출원 집계" onclick="location.href='appl_stat.asp';" style="<%=BUTTON_Style%>">
					<input type="button" value="국내 집계" onclick="location.href='appl_cnt_custom.asp';" style="<%=BUTTON_Style%>">
					<input type="button" value="해외 집계" onclick="location.href='appl_cnt_OGcustom.asp';" style="<%=BUTTON_Style%>">
				</td>
				<td rowspan="3" bgcolor="<%=t_color2%>"  align="left" height="25" style="width:22%">
					<input type="button" value="타임 시트" onclick="location.href='aspdoc/result_timesheet.asp';" style="<%=BUTTON_Style%>">				
				</td>
			</tr>
			</table>
		</td>
	</tr>

<!-- 기타집계표 START -->
	<tr>
		<td>
			<table width="980"  border="0" cellpadding="5" cellspacing="1" bgcolor="<%=main_right_C_color%>" align="right" style="margin:0;">
			<tr align="center">
				<td bgcolor="<%=t_color5%>" align="left"  height="25" width="13%"><span style="font-size:10pt;color:#0080FF;font-weight:bold; margin:0;">기타집계</span></td>
				<td bgcolor="<%=t_color5%>" align="left"  height="25" width="67%">
					<input type="button" value="사건관리" onclick="location.href='manage_comment.asp';" style="<%=BUTTON_Style%>">
					<input type="button" value="신건/OA/해외" onclick="location.href='list_up.asp';" style="<%=BUTTON_Style%>">
					<input type="button" value="특허청접수" onclick="location.href='manage_RevList.asp';" style="<%=BUTTON_Style%>">
					<input type="button" value="사건현황" onclick="location.href='find_case.asp';" style="<%=BUTTON_Style%>">

				</td>
				<td rowspan="3" bgcolor="<%=t_color2%>"  align="left" height="25" style="width:22%"></td>
			</tr>
			</table>
		</td>
	</tr>

<!-- 본문 END -->
</table>
</body>
</html>
<%
oConn.Close
Set oConn = Nothing
%>
