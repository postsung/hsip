<!-- #include file="include/dbcon.asp" -->
<!-- #include file="include/Session_chk.asp" -->
<!-- #include file="include/function.asp" -->

<%
Const BUTTON_Style = "border-width:1px;border-color:#666600;border-style:solid;background-color:#F3EFE2; padding-top:2px;cursor:hand; width:120px;"
Const t_color5 = "#ffffff" 
Const t_color2 = "#EEF3FB"
Const f_style = "FONT-SIZE: 9pt; FONT-FAMILY: '����', '����'; border: 1px solid #AEAFBB"
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

function Sform5Submit()  //�ɻ��û�� �ٿ�ε��ư
{	Sform5.submit();}

function SformSubmit()				//��Ǹ���Ʈ �˻���ư
{
	if (Sform.CustomerCode.value == "")
	{
		alert("���ڵ带 �Է��� �ּ���.");
		return;
	}
	document.getElementById("Sform").action="aspdoc/result_status.asp";
	Sform.submit();	
}

function SformSubmitApplicant()		//������ڵ� �˻���ư
{
	if (Sform.CustomerCode.value == "")
	{
		alert("���ڵ带 �Է��� �ּ���.");
		return;
	}
	document.getElementById("Sform").action="aspdoc/result_ApplicantCode.asp";
	Sform.submit();	
}

</script>
</head>

<%
Dim h_style, g_style, gs_style, hs_style
hs_style = "border-width: 1px; border-left-style: solid;	border-right-style: solid;	border-top-style: solid; border-bottom-style: solid;"   '����ռ� �׵θ�
h_style =  "border-width: 1px; border-right-style: solid;	border-top-style: solid;	border-bottom-style: solid;"							'����޼� �׵θ�
gs_style = "border-width: 1px; border-left-style: solid;	border-right-style: solid;	border-bottom-style: solid;"							'����ռ� �׵θ�
g_style =  "border-width: 1px; border-right-style: solid;	border-bottom-style: solid;"														'����޼� �׵θ�
%>

<!-- ��û����Ʈ START -->		
<table border="0" cellpadding="0" style="padding: 4px" width="100%" cellspacing="0" align="left">
	<td>&nbsp;</td>
	<tr>
	<td><span style="font-size:13pt;font-weight:bold; margin:0;">&nbsp; ��û����Ʈ</span></td>
	</tr>
</table>
		
<form id="Sform5" name="Sform5" action="aspdoc/result_Requestexamlist.asp" method="post" style="margin:0;">
<table width="900"  border="1" cellpadding="5" cellspacing="0" align="left">
	<tr align="center">
	<td align="left" width="800">		
	<span style="font-size:10pt;valign=center">����:	
	<select name="sNKind">
		<option value="A">����</option>
		<option value="B" selected>��Ŀ��</option>
		<option value="C">���</option>
	</select>&nbsp;
	�ڵ�: <input type="text" valign = "center" name="CustomerCode" size="9" />&nbsp;&nbsp;
	�˻��Ⱓ:
	<select name="StartYear">
	<%
	For i = Year(Date)-3 To Year(Date)+5
	%>
	<option value="<%=i%>" <%If Year(Date) = i-1 Then Response.Write "selected"%>><%=i%>��</option>
	<%
	Next
	%>			
	</select>
	<select name="StartMonth">
	<%
	For i = 1 To 12
	%>
		<option value="<%=i%>"><%=i%>��</option>
		<%
	Next
		%>			
	</select> ~
	<select name="EndYear">
		<%
			For i = Year(Date)-3 To Year(Date)+6
		%>
			<option value="<%=i%>" <%If Year(Date) = i-1 Then Response.Write "selected"%>><%=i%>��</option>
		<%
			Next
		%>			
			</select>
			<select name="EndMonth">
		<%
			For i = 1 To 12
		%>
			<option value="<%=i%>"<%If i = 12 Then Response.Write "selected"%>><%=i%>��</option>
		<%
			Next
		%>			
	</select>
	</span>
	</td>
	<td rowspan="3" valign="center" align="center" width="100">
		<input id="btnSubmit5" type="button" value=" ������� " onclick="Sform5Submit();" style="<%=BUTTON_Style%>">&nbsp;
	</td>
	</tr>
</table>
</form>
<!-- ��û������ END -->

<!-- ���� ��Ǹ���Ʈ START -->		
<table border="0" cellpadding="5" width="100%" cellspacing="0" align="left">
	<td>&nbsp;</td>
	<tr>
	<td height="20"><span style="font-size:13pt;font-weight:bold; margin:0;">&nbsp; ������� �˻�</span></td>
	</tr>
</table>
		
<form id="Sform" name="Sform" method="post" style="margin:0;">
<table width="900"  border="1" cellpadding="5" cellspacing="0" align="left">
	<tr align="center">
	<td align="left"  height="20" width="200" valign="top">		
		<span style="font-size:10pt;">
		��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��: 
		<select name="sNKind">
			<option value="A">����</option>
			<option value="B">��Ŀ��</option>
			<option value="C" selected>���</option>
		</select>&nbsp;<br />
		��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��: 
		<select name="sGubun"> 
			<option value="A">Ư��</option>
			<option value="B">�ǿ�</option>
			<option value="C">������</option>
			<option value="D">��ǥ</option>
			<option value="E" selected>���</option>
		</select>&nbsp;<br />
		���ڵ�: <input name="CustomerCode" type="text" size="9"><br />
		��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��: 
		<select name="sKind"> 
			<option value="A">��������</option>
			<option value="B" selected>���</option>
		</select>&nbsp;<br />
		��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��: 
		<select name="sRegistration"> 
			<option value="A">�������</option>
			<option value="B">��ϻ��</option>
			<option value="C" selected>���</option>
		</select>&nbsp;<br />
		��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��: <input type="text" name="CompareMemo" size="15"><BR />
	</span>
	</td>
	<td align="left"  height="20" width="300" valign="top">
	<span style="font-size:10pt;">
		�� �� ��: <input name="SearchReceiveStartDate" type="Date" value="<%=SearchReceiveStartDate%>" size="10"> ~ <input name="SearchReceiveEndDate" type="Date" value="<%=SearchReceiveEndDate%>"  size="10"> <br />
		�� �� ��: <input name="SearchFilingStartDate"  type="Date" value="<%=SearchFilingStartDate%>"  size="10"> ~ <input name="SearchFilingEndDate"  type="Date" value="<%=SearchFilingEndDate%>"   size="10"> <br />
		��û����: <input name="SearchExamDueStartDate" type="Date" value="<%=SearchExamDueStartDate%>" size="10"> ~ <input name="SearchExamDueEndDate" type="Date" value="<%=SearchExamDueEndDate%>"  size="10"> <br />
		�ɻ�û��: <input name="SearchExamStartDate"    type="Date" value="<%=SearchExamStartDate%>"    size="10"> ~ <input name="SearchExamEndDate"    type="Date" value="<%=SearchExamEndDate%>"     size="10"> <br />
		�� �� ��: <input name="SearchGrantStartDate"    type="Date" value="<%=SearchGrantStartDate%>"  size="10"> ~ <input name="SearchGrantEndDate"   type="Date" value="<%=SearchGrantEndDate%>"    size="10"> <br />
	</span>


	</td>
		<td rowspan="3" valign="top" align="left" height="20" width="200">
			<input id="btnSubmit" type="button" value=" �� �� �� Ȳ(��) " onclick="SformSubmit();" style="<%=BUTTON_Style%>"><br />
			<input id="btnSubmit" type="button" value=" ������ڵ�(��) " onclick="SformSubmitApplicant();" style="<%=BUTTON_Style%>">
		</td>
	</tr>
</table>
</form>
<!-- ���� �����Ȳ����Ʈ END -->

</html>