<!-- #include file="include/dbcon.asp" -->
<!-- #include file="include/Session_chk.asp" -->
<!-- #include file="include/function.asp" -->

<%
Const BUTTON_Style = "border-width:1px;border-color:#666600;border-style:solid;background-color:#F3EFE2; padding-top:2px;cursor:hand; width:120px;"
Const t_color5 = "#ffffff" 
Const t_color2 = "#EEF3FB"
Const f_style = "FONT-SIZE: 9pt; FONT-FAMILY: '돋움', '굴림'; border: 1px solid #AEAFBB"
Const main_right_C_color = "#C4D2E9"
Dim Sql, Rs, i, sNKind, sNKindA
Dim SearchReceiveStartDate, SearchReceiveEndDate, SearchFilingStartDate, SearchFilingEndDate, SearchExamDueStartDate, SearchExamDueEndDate, SearchExamStartDate, SearchExamEndDate
Dim SearchGrantStartDate, SearchGrantEndDate
%>

<html>
<head>

<style>
	body { background-color:#C0C0C0; }
</style>

<script type="text/javascript">

function Sform5Submit()  //심사미청구 다운로드버튼
{	Sform5.submit();}

function SformSubmit()				//사건리스트 검색버튼
{
	if (Sform.CustomerCode.value == "")
	{
		alert("고객코드를 입력해 주세요.");
		return;
	}
	document.getElementById("Sform").action="aspdoc/result_status.asp";
	Sform.submit();	
}

function SformSubmitApplicant()		//출원인코드 검색버튼
{
	if (Sform.CustomerCode.value == "")
	{
		alert("고객코드를 입력해 주세요.");
		return;
	}
	document.getElementById("Sform").action="aspdoc/result_ApplicantCode.asp";
	Sform.submit();	
}

</script>
</head>

<%
Dim h_style, g_style, gs_style, hs_style
hs_style = "border-width: 1px; border-left-style: solid;	border-right-style: solid;	border-top-style: solid; border-bottom-style: solid;"   '헤더앞셀 테두리
h_style =  "border-width: 1px; border-right-style: solid;	border-top-style: solid;	border-bottom-style: solid;"							'헤더뒷셀 테두리
gs_style = "border-width: 1px; border-left-style: solid;	border-right-style: solid;	border-bottom-style: solid;"							'내용앞셀 테두리
g_style =  "border-width: 1px; border-right-style: solid;	border-bottom-style: solid;"														'내용뒷셀 테두리
%>

<!-- 심청리스트 START -->		
<table border="0" cellpadding="0" style="padding: 4px" width="100%" cellspacing="0" align="left">
	<td>&nbsp;</td>
	<tr>
	<td><span style="font-size:13pt;font-weight:bold; margin:0;">&nbsp; 심청리스트</span></td>
	</tr>
</table>
		
<form id="Sform5" name="Sform5" action="aspdoc/result_Requestexamlist.asp" method="post" style="margin:0;">
<table width="900"  border="1" cellpadding="5" cellspacing="0" align="left">
	<tr align="center">
	<td align="left" width="800">		
	<span style="font-size:10pt;valign=center">구분:	
	<select name="sNKind">
		<option value="A">내국</option>
		<option value="B" selected>인커밍</option>
		<option value="C">모두</option>
	</select>&nbsp;
	코드: <input type="text" valign = "center" name="CustomerCode" size="9" />&nbsp;&nbsp;
	검색기간:
	<select name="StartYear">
	<%
	For i = Year(Date)-3 To Year(Date)+5
	%>
	<option value="<%=i%>" <%If Year(Date) = i-1 Then Response.Write "selected"%>><%=i%>년</option>
	<%
	Next
	%>			
	</select>
	<select name="StartMonth">
	<%
	For i = 1 To 12
	%>
		<option value="<%=i%>"><%=i%>월</option>
		<%
	Next
		%>			
	</select> ~
	<select name="EndYear">
		<%
			For i = Year(Date)-3 To Year(Date)+6
		%>
			<option value="<%=i%>" <%If Year(Date) = i-1 Then Response.Write "selected"%>><%=i%>년</option>
		<%
			Next
		%>			
			</select>
			<select name="EndMonth">
		<%
			For i = 1 To 12
		%>
			<option value="<%=i%>"<%If i = 12 Then Response.Write "selected"%>><%=i%>월</option>
		<%
			Next
		%>			
	</select>
	</span>
	</td>
	<td rowspan="3" valign="center" align="center" width="100">
		<input id="btnSubmit5" type="button" value=" 영문양식 " onclick="Sform5Submit();" style="<%=BUTTON_Style%>">&nbsp;
	</td>
	</tr>
</table>
</form>
<!-- 심청마감일 END -->

<!-- 메인 사건리스트 START -->		
<table border="0" cellpadding="5" width="100%" cellspacing="0" align="left">
	<td>&nbsp;</td>
	<tr>
	<td height="20"><span style="font-size:13pt;font-weight:bold; margin:0;">&nbsp; 국내출원 검색</span></td>
	</tr>
</table>
		
<form id="Sform" name="Sform" method="post" style="margin:0;">
<table width="900"  border="1" cellpadding="5" cellspacing="0" align="left">
	<tr align="center">
	<td align="left"  height="20" width="200" valign="top">		
		<span style="font-size:10pt;">
		구&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;분: 
		<select name="sNKind">
			<option value="A">내국</option>
			<option value="B">인커밍</option>
			<option value="C" selected>모두</option>
		</select>&nbsp;<br />
		권&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;리: 
		<select name="sGubun"> 
			<option value="A">특허</option>
			<option value="B">실용</option>
			<option value="C">디자인</option>
			<option value="D">상표</option>
			<option value="E" selected>모두</option>
		</select>&nbsp;<br />
		고객코드: <input name="CustomerCode" type="text" size="9"><br />
		포&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;기: 
		<select name="sKind"> 
			<option value="A">포기제외</option>
			<option value="B" selected>모두</option>
		</select>&nbsp;<br />
		등&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;록: 
		<select name="sRegistration"> 
			<option value="A">등록제외</option>
			<option value="B">등록사건</option>
			<option value="C" selected>모두</option>
		</select>&nbsp;<br />
		메&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;모: <input type="text" name="CompareMemo" size="15"><BR />
	</span>
	</td>
	<td align="left"  height="20" width="300" valign="top">
	<span style="font-size:10pt;">
		접 수 일: <input name="SearchReceiveStartDate" type="Date" value="<%=SearchReceiveStartDate%>" size="10"> ~ <input name="SearchReceiveEndDate" type="Date" value="<%=SearchReceiveEndDate%>"  size="10"> <br />
		출 원 일: <input name="SearchFilingStartDate"  type="Date" value="<%=SearchFilingStartDate%>"  size="10"> ~ <input name="SearchFilingEndDate"  type="Date" value="<%=SearchFilingEndDate%>"   size="10"> <br />
		심청마감: <input name="SearchExamDueStartDate" type="Date" value="<%=SearchExamDueStartDate%>" size="10"> ~ <input name="SearchExamDueEndDate" type="Date" value="<%=SearchExamDueEndDate%>"  size="10"> <br />
		심사청구: <input name="SearchExamStartDate"    type="Date" value="<%=SearchExamStartDate%>"    size="10"> ~ <input name="SearchExamEndDate"    type="Date" value="<%=SearchExamEndDate%>"     size="10"> <br />
		등 록 일: <input name="SearchGrantStartDate"    type="Date" value="<%=SearchGrantStartDate%>"  size="10"> ~ <input name="SearchGrantEndDate"   type="Date" value="<%=SearchGrantEndDate%>"    size="10"> <br />
	</span>


	</td>
		<td rowspan="3" valign="top" align="left" height="20" width="200">
			<input id="btnSubmit" type="button" value=" 사 건 현 황(영) " onclick="SformSubmit();" style="<%=BUTTON_Style%>"><br />
			<input id="btnSubmit" type="button" value=" 출원인코드(영) " onclick="SformSubmitApplicant();" style="<%=BUTTON_Style%>">
		</td>
	</tr>
</table>
</form>
<!-- 메인 사건현황리스트 END -->

</html>