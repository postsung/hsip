<!-- #include file="include/dbcon.asp" -->
<!-- #include file="include/Session_chk.asp" -->
<!-- #include file="include/function.asp" -->
<%
Const BUTTON_Style = "FONT-SIZE: 9pt; border-width:1px;border-color:#666600;border-style:solid;background-color:#F3EFE2; padding-top:2px;cursor:hand;width:90px;"
Const t_color5 = "#ffffff" 
Const t_color2 = "#EEF3FB"
Const f_style = "FONT-SIZE: 8pt; FONT-FAMILY: '����', '����'; border: 1px solid #AEAFBB"
Const main_right_C_color = "#C4D2E9"
Dim Sql, Rs, i

%>
<html>
<head>
<script type="text/javascript">

function ButtonDisplay(id) //�Լ�_��������ǥ �ٿ�ε� ��ư
{
	document.getElementById(id).disabled = false;
	document.getElementById(id).value = " �ٿ�ε� ";
}

function ExchangeDisplay(n) //�Լ�_ȯ������ �˾�
{
	if (document.getElementById(n).style.display == "none")
	{	document.getElementById(n).style.display = "block";	}
	else
	{	document.getElementById(n).style.display = "none";	}
}

function Sform2Submit()  //�Լ�_��������ǥ ����/�Ⱓ����
{
	if (Sform2.sPaper.value == "")
	{
		alert("������ �����Ͽ� �ּ���.");
		return;
	}
	if (Sform2.sKind.value == "")
	{
		alert("�Ⱓ������ �����Ͽ� �ּ���.");
		return;
	}
	Sform2.submit();	
}

function sPaperChanged(v) //�Լ�_��������ǥ ��������
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
		Sform2.StartYear.value = cdate.getFullYear(); // �˻��Ⱓ �ʱ�ȭ(����3��~����2��)
		Sform2.StartMonth.value = 3;
		Sform2.EndYear.value = cdate.getFullYear()+1;
		Sform2.EndMonth.value = 2;
	}
}

function MonthChk(t)  //�Լ�_��������ǥ �˻��� �Է¿����޽���
{
	if (t == "F")
	{
		if (document.getElementById("sPaper").value == "D"&& Sform2.StartMonth.value < 3)
		{
			alert("�˻��Ⱓ ���ۿ��� 2�� ������ ������ �� �����ϴ�.");
			Sform2.StartMonth.value = 3;		
		}
	}
}

function Sform4Submit()  //�Լ�_�Աݼ����� �ٿ�ε��ư
{	Sform4.submit();}

function Sform1Submit() //�Լ�_�̼�������ǥ �ٿ�ε��ư
{	document.getElementById().action = "result_unpaidlist.asp"; }

function Sform3Submit() //�Լ�_Ư������ǥ �Է¿����޽���
{	
	if (document.getElementById("CustNum").value == "")
	{
		alert("�Ƿ����� �����Ͽ� �ּ���.");
		return;
	}
	document.getElementById("Sform3").action = "result_patent.asp";
	Sform3.submit();	
}

function Sform3Search() //�Լ�_Ư������ǥ �Ƿ����ڵ� �Է¿����޽���
{
	document.getElementById("Sform3").action = "index.asp";
	if (document.getElementById("CustRef").value == "")
	{
		alert("�˻�� �Է��Ͽ� �ּ���.");
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

<!------ ���� START ------>

<body text="#000000" bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<p>&nbsp;</p>

<table border="0">
	<tr>
		<td>
			<table border="0" cellpadding="0" style="padding: 4px; height: 46px;" width="1000" cellspacing="0">
			<tr>
				<td height="20" class="auto-style2"><strong>Ư����� �Ѽ�</strong></td>
			</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td class="auto-style1"><a href="logout.asp">�α׾ƿ�</a> </td>
	</tr>

<!-- ��������ǥ START -->
	<tr>
		<td>
			<form id="Sform2" name="Sform2" method="post" style="margin:0;">
			<table width="980"  border="0" cellpadding="5" cellspacing="1" bgcolor="<%=main_right_C_color%>" align="right">
			<tr align="center">
				<td bgcolor="<%=t_color5%>" align="left"  height="25" width="13%"><span style="font-size:10pt;color:#0080FF;font-weight:bold; margin:0;">��������ǥ</span></td>
				<td bgcolor="<%=t_color5%>" align="left"  height="25" width="67%">
					<select id="sPaper" name="sPaper" onchange="javascript:ButtonDisplay('btnSubmit2');sPaperChanged(this.value);">
						<option value="">--- ���� ---</option>
						<option value="A">�μ������� ����ǥ</option>
						<option value="B">���κ����� ����ǥ</option>
						<option value="C">û������ ����ǥ</option>
						<option value="D">�������� ����ǥ</option>
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
						<option value="">- �Ⱓ���� -</option>
						<option value="A" selected>û����</option>
					</select>
					<select name="StartYear" onchange="javascript:ButtonDisplay('btnSubmit2');">
					<%
					For i = Year(Date)-3 To Year(Date)+1
					%>
						<option value="<%=i%>" <%If Year(Date) = i Then Response.Write "selected"%>><%=i%>��</option>
					<%
					Next
					%>			
					</select>
					<select name="StartMonth" onchange="javascript:ButtonDisplay('btnSubmit2');MonthChk('F');">
					<%
					For i = 1 To 12
					%>
						<option value="<%=i%>"><%=i%>��</option>
					<%
					Next
					%>			
					</select> ~
					<select name="EndYear" onchange="javascript:ButtonDisplay('btnSubmit2');">
					<%
					For i = Year(Date)-3 To Year(Date)+1
					%>
						<option value="<%=i%>" <%If Year(Date) = i Then Response.Write "selected"%>><%=i%>��</option>
					<%
					Next
					%>			
					</select>
					<select name="EndMonth" onchange="javascript:ButtonDisplay('btnSubmit2');MonthChk('L');">
					<%
					For i = 1 To 12
					%>
						<option value="<%=i%>"><%=i%>��</option>
					<%
					Next
					%>			
					</select>
				</td>
				<td rowspan="3" bgcolor="<%=t_color2%>"  align="left" height="25" style="width: 20%">
					<input id="btnSubmit2" type="button" value=" �ٿ�ε� " onclick="Sform2Submit();" style="<%=BUTTON_Style%>">
					<input id="btnSubmit2_2" type="button" value=" ȯ������ " onclick="ExchangeDisplay('dvExchange2');" style="<%=BUTTON_Style%>"> 
				</td>
			</tr>
			</table>
			</form>

<!-- ȯ������ START -->
			<div id="dvExchange2" style="display:none;position:absolute; left: 455px">
				<table width="300" border="0" cellpadding="2" cellspacing="1" bgcolor="#CFC4E9" align="center" style="font-size:10pt" >	
					<tr>
						<td bgcolor="#F2EEFB" align="center" height="25" width="80">�� ȭ</td>
						<td bgcolor="#F2EEFB" align="center" height="25" width="80">������</td>
						<td bgcolor="#F2EEFB" align="center" height="25" width="80">ȯ ��</td>
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
			
<!-- �Աݼ����� START -->		
	<tr>
		<td>
			<form id="Sform4" name="Sform4" action="payfee_list.asp" method="post" style="margin:0;">
			<table width="980"  border="0" cellpadding="5" cellspacing="1" bgcolor="<%=main_right_C_color%>" align="right">
			<tr align="center">
				<td bgcolor="<%=t_color5%>" align="left"  height="25" width="13%"><span style="font-size:10pt;color:#0080FF;font-weight:bold; margin:0;">�Աݼ�����</span></td>
				<td bgcolor="<%=t_color5%>" align="left"  height="25" width="67%">
					<select name="sKind" onchange="javascript:ButtonDisplay('btnSubmit4');">
						<option value="">--- �Ⱓ���� ---</option>
						<option value="A" selected>û����</option>
					</select>&nbsp;
					<select name="StartYear" onchange="javascript:ButtonDisplay('btnSubmit4');">
					<%
					For i = Year(Date)-3 To Year(Date)
					%>
						<option value="<%=i%>" <%If Year(Date) = i Then Response.Write "selected"%>><%=i%>��</option>
					<%
					Next
					%>			
					</select>
					<select name="StartMonth" onchange="javascript:ButtonDisplay('btnSubmit4');">
					<%
					For i = 1 To 12
					%>
						<option value="<%=i%>"><%=i%>��</option>
					<%
					Next
					%>			
					</select> ~
					<select name="EndYear" onchange="javascript:ButtonDisplay('btnSubmit4');">
					<%
					For i = Year(Date)-3 To Year(Date)+1
					%>
						<option value="<%=i%>" <%If Year(Date) = i Then Response.Write "selected"%>><%=i%>��</option>
					<%
					Next
					%>			
					</select>
					<select name="EndMonth" onchange="javascript:ButtonDisplay('btnSubmit4');">
					<%
					For i = 1 To 12
					%>
						<option value="<%=i%>"><%=i%>��</option>
					<%
					Next
					%>			
					</select>
					&nbsp;<br>
					<select name="sKind2" onchange="javascript:ButtonDisplay('btnSubmit4');">
						<option value="">--- �Ⱓ���� ---</option>
						<option value="A" selected>�Ա���</option>
					</select>&nbsp;
					<select name="StartYear2" onchange="javascript:ButtonDisplay('btnSubmit4');">
					<%
					For i = Year(Date)-4 To Year(Date)+1
					%>
						<option value="<%=i%>" <%If Year(Date) = i Then Response.Write "selected"%>><%=i%>��</option>
					<%
					Next
					%>			
					</select>
					<select name="StartMonth2" onchange="javascript:ButtonDisplay('btnSubmit4');">
					<%
					For i = 1 To 12
					%>
						<option value="<%=i%>"><%=i%>��</option>
					<%
					Next
					%>			
					</select> ~
					<select name="EndYear2" onchange="javascript:ButtonDisplay('btnSubmit4');">
					<%
					For i = Year(Date)-4 To Year(Date)
					%>
						<option value="<%=i%>" <%If Year(Date) = i Then Response.Write "selected"%>><%=i%>��</option>
					<%
					Next
					%>			
					</select>
					<select name="EndMonth2" onchange="javascript:ButtonDisplay('btnSubmit4');">
					<%
					For i = 1 To 12
					%>
						<option value="<%=i%>"><%=i%>��</option>
					<%
					Next
					%>			
					</select>
					
				</td>
				<td rowspan="3" bgcolor="<%=t_color2%>"  align="left" height="25" width="22%">
					<input id="btnSubmit4" type="button" value=" �ٿ�ε� " onclick="Sform4Submit();" style="<%=BUTTON_Style%>">&nbsp;&nbsp;
				</td>
			</tr>
			</table>
			</form>

<!-- �̼������� START -->
	<tr>
		<td>
			<table width="980"  border="0" cellpadding="5" cellspacing="1" bgcolor="<%=main_right_C_color%>" align="right" style="margin:0; height: 5px;">
			<tr align="center">
				<td bgcolor="<%=t_color5%>" align="left"  height="25" width="13%"><span style="font-size:10pt;color:#0080FF;font-weight:bold; margin:0;">�̼��� ����ǥ</span></td>
				<td bgcolor="<%=t_color5%>" align="left"  height="25" width="67%">
					<input id="btnSubmit1" type="button" value=" �̼��� ���� " onclick="location.href='manage_invoice.asp';" style="<%=BUTTON_Style%>">
				</td>
				<td rowspan="3" bgcolor="<%=t_color2%>"  align="left" height="25" style="width: 22%"></td>
			</tr>
			</table>
		</td>
	</tr>

<!-- Ư������ǥ START -->
<%
Dim CustRef
CustRef = Request("CustRef")
%>
	<tr>
		<td>
			<form id="Sform3" name="Sform3" action="index.asp" method="post" onsubmit="return Sform3Search();" style="margin:0;">
			<table width="980"  border="0" cellpadding="5" cellspacing="1" bgcolor="<%=main_right_C_color%>" align="right">
			<tr align="center">
				<td bgcolor="<%=t_color5%>" align="left"  height="25" width="13%"><span style="font-size:10pt;color:#0080FF;font-weight:bold; margin:0;">Ư������ǥ</span></td>
				<td bgcolor="<%=t_color5%>" align="left"  height="25" width="67%"><span style="font-size:10pt;font-weight:margin:0;">�Ƿ����ڵ� : 
					<input type="text" id="CustRef" name="CustRef" value="<%=CustRef%>" size="6" onclick="javascript:ButtonDisplay('btnSubmit3');">
					<input type="submit" value=" �� �� " >&nbsp;&nbsp;
					<select id="CustNum" name="CustNum">
					<option value="">--- �Ƿ��� ���� ---</option>
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
					<input id="btnSubmit3" type="button" value=" �ٿ�ε� " onclick="Sform3Submit();" style="<%=BUTTON_Style%>">&nbsp;&nbsp;
				</td>
			</tr>
			</table>
			</form>	
		</td>
	</tr>


<!-- ���� ������ START -->
	<tr>
		<td>
			<table width="980"  border="0" cellpadding="5" cellspacing="1" bgcolor="<%=main_right_C_color%>" align="right" style="margin:0; height: 5px;">
			<tr align="center">
				<td bgcolor="<%=t_color5%>" align="left"  height="25" width="13%"><span style="font-size:10pt;color:#0080FF;font-weight:bold; margin:0;">���� ������</span></td>
				<td bgcolor="<%=t_color5%>" align="left"  height="25" width="67%">
					<input type="button" value="��� ����" onclick="location.href='appl_stat.asp';" style="<%=BUTTON_Style%>">
					<input type="button" value="���� ����" onclick="location.href='appl_cnt_custom.asp';" style="<%=BUTTON_Style%>">
					<input type="button" value="�ؿ� ����" onclick="location.href='appl_cnt_OGcustom.asp';" style="<%=BUTTON_Style%>">
				</td>
				<td rowspan="3" bgcolor="<%=t_color2%>"  align="left" height="25" style="width:22%">
					<input type="button" value="Ÿ�� ��Ʈ" onclick="location.href='aspdoc/result_timesheet.asp';" style="<%=BUTTON_Style%>">				
				</td>
			</tr>
			</table>
		</td>
	</tr>

<!-- ��Ÿ����ǥ START -->
	<tr>
		<td>
			<table width="980"  border="0" cellpadding="5" cellspacing="1" bgcolor="<%=main_right_C_color%>" align="right" style="margin:0;">
			<tr align="center">
				<td bgcolor="<%=t_color5%>" align="left"  height="25" width="13%"><span style="font-size:10pt;color:#0080FF;font-weight:bold; margin:0;">��Ÿ����</span></td>
				<td bgcolor="<%=t_color5%>" align="left"  height="25" width="67%">
					<input type="button" value="��ǰ���" onclick="location.href='manage_comment.asp';" style="<%=BUTTON_Style%>">
					<input type="button" value="�Ű�/OA/�ؿ�" onclick="location.href='list_up.asp';" style="<%=BUTTON_Style%>">
					<input type="button" value="Ư��û����" onclick="location.href='manage_RevList.asp';" style="<%=BUTTON_Style%>">
					<input type="button" value="�����Ȳ" onclick="location.href='find_case.asp';" style="<%=BUTTON_Style%>">

				</td>
				<td rowspan="3" bgcolor="<%=t_color2%>"  align="left" height="25" style="width:22%"></td>
			</tr>
			</table>
		</td>
	</tr>

<!-- ���� END -->
</table>
</body>
</html>
<%
oConn.Close
Set oConn = Nothing
%>
