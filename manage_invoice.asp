<!-- #include file="include/dbcon.asp" -->
<!-- #include file="include/Session_chk.asp" -->
<!-- #include file="include/function.asp" -->
<!-- #include file="include/function_invoice.asp" -->
<%
Server.ScriptTimeOut = 6000

Dim sKind, sKindTitle, sql, sql1, rs, rs1, rs2, InvAgent, InvAgentNum, sAgent, AgentNumTemp
Dim SearchStartDate, SearchEndDate
Dim h_style, g_style, gs_style, hs_style

hs_style = "border-width: 1px; border-left-style: solid;	border-right-style: solid;	border-top-style: solid; border-bottom-style: solid;"   '����ռ� �׵θ�
h_style =  "border-width: 1px; border-right-style: solid;	border-top-style: solid;	border-bottom-style: solid;"							'����޼� �׵θ�
gs_style = "border-width: 1px; border-left-style: solid;	border-right-style: solid;	border-bottom-style: solid;"							'����ռ� �׵θ�
g_style =  "border-width: 1px; border-right-style: solid;	border-bottom-style: solid;"														'����޼� �׵θ�

sKind = Request("sKind")
sAgent = Request("sAgent")
InvAgentNum = Request("InvAgentNum")
InvAgent = Request("InvAgent")
SearchStartDate = Request("SearchStartDate")
SearchEndDate = Request("SearchEndDate")

If SearchStartDate = "" Then '�ʱⰪ
	SearchStartDate = Date()
	SearchEndDate = Date()
End If

Select Case sKind
	Case "A"
		sKindTitle = "�̼��ݸ���Ʈ(��ȭ)"
	Case "B"
		sKindTitle = "�̼��ݸ���Ʈ(��ȭ)"
End Select
%>

<html>
<head>
<meta content="ko" http-equiv="Content-Language">
<link rel="stylesheet" type="text/css" href="../../../include/style.css">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
	background-color: #7694C8;
}
td {  font-size: 9pt;}
.style1 {color: #01215A}
.style2 {font-weight: bold}
.style3 {color: #000000}
.style4 {color: #000000}
-->
</style>

<script type="text/javascript">
function SformSubmit()  //�Լ�_�̼��� Ȯ�ι�ư
{	Sform.submit();}

function Sform1Submit() //�Լ�_AgentNum�Է� �����޽���
{	Sform1.submit();}


</script>

</head>

<body text="#000000" bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" class="style3">
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td valign="top" bgcolor="#ffffff">																				
			<table width="100%" border="0" cellspacing="1" cellpadding="0">
			  <!--  ������ ���� ��θ� �����ִ� tr-->
				<tr>
					<td bgcolor="#7694C8">
					<table width="100%" height="28" border="0" cellpadding="0" cellspacing="1">					   
						<tr>
							<td align="center" bgcolor="#EEF3FB">
							<table width="100%" height="28" border="0" cellpadding="0" cellspacing="1">
								<tr>
									<td align="left" bgcolor="#B6C7E5">&nbsp;
										<strong>Main &gt; û��������/ �̼��ݸ���Ʈ</strong>
									</td>
									<td align="right"bgcolor="#B6C7E5"><a href="logout.asp">(�α׾ƿ�)</a>
									</td>
								</tr>
							</table>
							</td>
						</tr>
					</table>
					</td>
				</tr>
				<tr><td>
			</table>
			<table width="100%" border="0" cellspacing="0" cellpadding="1">					  

<!---------------- ���� START ------------------------------------------------->

						<tr>
							<td>
								<form id="Sform" name="Sform" action="manage_invoice.asp" method="post">
									<input type="hidden" name="Ps" value="Ps">
									<td bgcolor="#D8DEEA" align="left" width="250" height="25">&nbsp&nbsp;�̼�ȭ�� ����:
										<select id="sKind" name="sKind">
											<option value="A" <%If sKind = "A" Then Response.Write "selected"%>>��ȭ</option>
											<option value="B" <%If sKind = "B" Then Response.Write "selected"%>>��ȭ</option>
										</select>&nbsp;
									<input id="btnSubmit" type="button" value="�˻��ϱ�" onclick="SformSubmit();">&nbsp;
									</td>
								</form>
							</td>
							<td>
								<form id="Sform1" name="Sform1" action="manage_invoice.asp" method="post">
									<input type="hidden" name="Ps1" value="Ps1">
									<td bgcolor="#D8DEEA" align="left"  height="10">&nbsp&nbsp;
										����: 
											<select id="sAgent" name="sAgent">
												<option value="A"<%If sAgent = "A" Then Response.Write "selected"%>>��ü</option>
												<option value="B"<%If sAgent = "B" Then Response.Write "selected"%>>��</option>
											</select>&nbsp;
										
										�۱� Agent: <input type="text" id="InvAgent" name="InvAgent" value="<%=InvAgent%>" size="15">
										<input type="submit" value=" �˻� ">
										<select id="InvAgentNum" name="InvAgentNum" >
											<option value="">---- Agent ���� ----</option>
											<%
											If InvAgent <> "" Then
												Sql = "SELECT Num, Field41, mName FROM Customer WHERE mName LIKE'%" & InvAgent & "%' "
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
										�˻��Ⱓ: 
											<input name="SearchStartDate"	type="Date" value="<%=SearchStartDate%>" size="10"> ~ 
											<input name="SearchEndDate"		type="Date" value="<%=SearchEndDate%>"	 size="10"> &nbsp;
											<input id="btnSubmit1" type="button" value="�˻��ϱ�" onclick="Sform1Submit();">&nbsp;
										</td>
								</form>								
							</td>
						</tr>
				</table>


		<%
		If sKind = "A" Then
			Call sbUnpaidListW() '��ȭ�̼��ݸ���Ʈ
		ElseIf sKind ="B" Then 
			Call sbUnpaidListD() '��ȭ�̼��ݸ���Ʈ
		End If

		If sAgent ="A" Then
			Call sbOGTotSendMoney()
		ElseIf (sAgent="B" And InvAgentNum <> "") Then
			Call sbOGSendMoney()
		End If
		%>

					<!-- ���� END -->			
</table>

</body>
</html>
<%
oConn.Close
Set oConn = Nothing
'------------------------------------------------------------------------------------------------------------------------
%>