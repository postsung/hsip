<!-- #include file="include/dbcon.asp" -->
<!-- #include file="include/Session_chk.asp" -->
<!-- #include file="include/function_cmt.asp" -->
<%
Server.ScriptTimeOut = 6000

Dim h_style, g_style, gs_style, hs_style
hs_style = "border-width: 1px; border-left-style: solid;	border-right-style: solid;	border-top-style: solid; border-bottom-style: solid;"   '����ռ� �׵θ�
h_style =  "border-width: 1px; border-right-style: solid;	border-top-style: solid;	border-bottom-style: solid;"							'����޼� �׵θ�
gs_style = "border-width: 1px; border-left-style: solid;	border-right-style: solid;	border-bottom-style: solid;"							'����ռ� �׵θ�
g_style =  "border-width: 1px; border-right-style: solid;	border-bottom-style: solid;"														'����޼� �׵θ�

Dim sKind, Ps, Sql, Rs, sKindTitle, sMethod, sPart, sDeley, Cmt, CmtTitle
Dim SearchStartDate, SearchEndDate

Ps = Request("Ps")
sKind = Request("sKind")
sPart = Request("sPart")
sMethod = Request("sMethod")
sDeley = Request("sDeley")
CmtTitle = Request("CmtTitle")

SearchStartDate = Request("SearchStartDate")
SearchEndDate = Request("SearchEndDate")

If Ps = "" Then '�ʱⰪ
	SearchStartDate = Date()
	SearchEndDate = Date()
End If

Select Case sKind
	Case "A"
		sKindTitle = "����������"
	Case "B"
		sKindTitle = "Ư��û����"
	Case "C"
		sKindTitle = "���Ű���"
	Case "D"
		sKindTitle = "Due Date����"
	Case "E"
		sKindTitle = "�ؿ�����ȳ�"
	Case "F"
		sKindTitle = "�ɻ�û�� �����Ͼȳ�"
End Select

If (sKind = "A" Or sKind = "B" Or sKind = "C" Or sKind = "F" Or sKind = "G" ) then
	Select Case sMethod
		Case "A"
			sKindTitle = sKindTitle & "(����/�߼ۼ���)"
		Case "B"
			sKindTitle = sKindTitle & "(��������)"
		Case "C"
			sKindTitle = sKindTitle & "(�߼ۼ���)"
	End Select
End If


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

<script language="javascript">
function SformSubmit()   '�����ư
{	
	document.getElementById("btnSearch").disabled = true;
	document.getElementById("btnSearch").value = " �˻��� ... ";
	return true;
}

</script>
</head>

<body text="#000000" bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" class="style3">
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td valign="top" bgcolor="#E9EBF3">
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
										<strong><a href="index.asp">Main</a> &gt; Ư��û/������</strong>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<a href="logout.asp">(�α׾ƿ�)</a>
									</td>
								</tr>
							</table>
							</td>
						</tr>
					</table>
					</td>
				</tr>
				<tr><td>
				<table width="100%" border="0" cellspacing="0" cellpadding="0">					  
				<tr>
					<td>

					<!-- ���� START -->

					<table width="100%" border="0" cellspacing="0" cellpadding="0">
						<tr>
							<td height="5">
							<table width="100%" border="0" cellspacing="5" cellpadding="1" id="table34">
								<tr>
									<td>
										
										<table border="0" cellpadding="0" style="padding: 4px" width="1000" id="table17" cellspacing="0" align="left">
										<tr>
											<td>
												
												<form id="Sform" name="Sform" action="manage_comment.asp" method="post" onsubmit="return SformSubmit();">
												<input type="hidden" name="Ps" value="Ps">
												<table width="1000"  border="0" cellpadding="1" cellspacing="1" id="table1" align="center">
													<tr align="center">	
														<td bgcolor="#ffffff" align="left"  height="20">&nbsp;
															<select id="sKind" name="sKind">
																<option value="">---- �� �� ----</option>
																<option value="A" <%If sKind = "A" Then Response.Write "selected"%>>������ ���� </option>
																<option value="B" <%If sKind = "B" Then Response.Write "selected"%>>Ư��û����</option>
																<option value="C" <%If sKind = "C" Then Response.Write "selected"%>>���� ����</option>
																<option value="D" <%If sKind = "D" Then Response.Write "selected"%>>Due Date����</option>
																<option value="E" <%If sKind = "E" Then Response.Write "selected"%>>�ؿ��������</option>
																<option value="F" <%If sKind = "F" Then Response.Write "selected"%>>�ɻ��û�� ����Ʈ</option>
															</select>&nbsp;

															�μ�:
															<select id="sPart" name="sPart">
																<option value="A" <%If sPart = "A" Then Response.Write "selected"%>>���</option>
																<option value="B" <%If sPart = "B" Then Response.Write "selected"%>>ȭ��</option>
																<option value="C" <%If sPart = "C" Then Response.Write "selected"%>>���</option>
																<option value="D" <%If sPart = "D" Then Response.Write "selected"%>>����</option>
																<option value="E" <%If sPart = "E" Then Response.Write "selected"%>>��ǥ</option>
															</select>&nbsp;

															<select id="sMethod" name="sMethod">
																<option value="A" <%If sMethod = "A" Then Response.Write "selected"%>>��繮��</option>
																<option value="B" <%If sMethod = "B" Then Response.Write "selected"%>>��������</option>
																<option value="C" <%If sMethod = "C" Then Response.Write "selected"%>>�߼۹���</option>
															</select>&nbsp;

															�˻��Ⱓ: <input name="SearchStartDate" type="Date" value="<%=SearchStartDate%>" size="10"> ~ <input name="SearchEndDate" type="Date" value="<%=SearchEndDate%>"  size="10"> &nbsp;

															ó��:
															<select id="sDeley" name="sDeley">
																<option value="A" <%If sDeley = "A" Then Response.Write "selected"%>>���</option>
																<option value="B" <%If sDeley = "B" Then Response.Write "selected"%>>��ó��</option>
																<option value="C" <%If sDeley = "C" Then Response.Write "selected"%>>����</option>
															</select>&nbsp;

															�˻�����:&nbsp; <input type="text" name="CmtTitle" size="18" />
														</td>
														<td  bgcolor="#ffffff" align="center"   width="100" rowspan="2" style='padding: 5px;'>
															<input type="submit" id="btnSearch" value="�˻��ϱ�">
														</td>
													</tr>
												</table>
												</form>
											</td>
										</tr>
										</table>
								</td>
							</tr>
					</table>

					<table>
							<%
							If sKind = "A" Then
								Call sbComment()
								Cmt = "[���� �� ��ó�� ���� �ȳ�] " + CStr(Date())
							ElseIf sKind ="B" Then 
								Call sbCommentChk()
								Cmt = "* �ּ�"
							ElseIf sKind ="C" Then 
								Call sbLetter()
								Cmt = "* �ּ�"
							ElseIf sKind ="D" Then 
								Call sbDueDate()
								Cmt = "* ���ϱ��� ó���ؾ� �ϴ� ��� ��, ������� SOPI�� ��ó���� ǥ��� �����Դϴ�. ������ �����Ͻñ� �ٶ��ϴ�." + "<br><br> [���ϱ��ϰ���] " + CStr(Date()) + " ��ó�� ��Ǿȳ�"
							ElseIf sKind ="E" Then 
								Call sbOutGoingList()
								Cmt = "* �ּ�"
							ElseIf sKind ="F" Then 
								Call sbRExamList()
								Cmt = "* �ּ�"
							End If
							%>

							<td align="left"><span style="font-weight: bold;"><%=Cmt%></span></td>
							</table>
							</td>
						</tr>						
					</table>
					<!-- ���� END -->			
</table>
</body>
</html>
<%
oConn.Close
Set oConn = Nothing
'------------------------------------------------------------------------------------------------------------------------
%>