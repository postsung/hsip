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

Dim sKind, Ps, Sql, Rs, sKindTitle, sMethod, Cmt, sPart, sSection, sHandling, CmtTitle
Dim SearchStartDate, SearchEndDate

Ps = Request("Ps")
sKind = Request("sKind")
sMethod = Request("sMethod")
sSection = Request("sSection")
sHandling = Request("sHandling")
sPart = Request("sPart")
CmtTitle = Request("CmtTitle")

SearchStartDate = Request("SearchStartDate")
SearchEndDate = Request("SearchEndDate")

If Ps = "" Then '�ʱⰪ
	SearchStartDate = DateSerial(Year(Date),Month(Date),1)
	SearchEndDate = Date()
End If

Select Case sKind
	Case "A"
		sKindTitle = "�Ű����"
	Case "B"
		sKindTitle = "OA ��������Ʈ"
	Case "C"
		sKindTitle = "�ؿܼ��� ��������Ʈ"
End Select

%>
<html>
<head>
<meta content="ko" http-equiv="Content-Language">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
	background-color: #7694C8;
}
td {  font-size: 10pt;}
.style1 {color: #01215A}
.style2 {font-weight: bold}
.style3 {color: #000000}
.style4 {color: #000000}
th {  font-size: 10pt;}
.style1 {color: #01215A}
.style2 {font-weight: bold}
.style3 {color: #000000}
.style4 {color: #000000}
-->
</style>
<link rel="stylesheet" type="text/css" href="include/base.css" />

<script src="include/jquery.min.js"></script>
<script type="text/javascript" src="include/jquery.fixheadertable.min.js"></script>
<style type="text/css">
.ui-widget-content { border: 1px solid #aaaaaa; background: #D8DEEA repeat-x; color: #000000; }
.ui-widget-header { border: 1px solid #aaaaaa; background: #D8DEEA repeat-x; color: #ffffff; }
</style>
<script language="javascript">
$(document).ready(function() {
	$('.fixedhead').fixheadertable({
		<% If sKind = "A" Then %>
			colratio    : [50, 90, 190, 150, 210, 90, 90, 90, 90, 90, 90, 200, 30],
			height      : 650,
			width       : $(window).width()-15, 
			resizeCol   : true,
			sortType    : ['string', 'string', 'string', 'string', 'string', 'string', 'string', 'string', 'string', 'string', 'string', 'string', 'string'],

		<% ElseIf sKind = "B" Then %>
			colratio    : [50, 90, 350, 150, 250, 90, 90, 90, 30],
			height      : 650, 
			width       : $(window).width()-15, 
			resizeCol   : true,
			sortType    : ['string', 'string', 'string', 'string', 'string', 'string', 'string', 'string', 'string'],

		<% ElseIf sKind = "C" Then %>
			colratio    : [80, 150, 200, 250, 90, 90, 90, 90, 90, 60, 30],
			height      : 650, 
			width       : $(window).width()-15, 
			resizeCol   : true,
			sortType    : ['string', 'string', 'string', 'string', 'string', 'string', 'string', 'string', 'string', 'string', 'string'],

		<% End If %>		
			
		minColWidth : 50 
	});
});

function SformSubmit()   //�����ư
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
			  <!--  ������ Ÿ��Ʋ -->
				<tr>
					<td bgcolor="#7694C8">
					<table width="100%" height="28" border="0" cellpadding="0" cellspacing="1">					   
						<tr>
							<td align="center" bgcolor="#EEF3FB">
							<table width="100%" height="28" border="0" cellpadding="0" cellspacing="1">
								<tr>
									<td align="left" bgcolor="#B6C7E5">&nbsp;
										<strong>&gt; ����Ʈ ����</strong>&nbsp;&nbsp;<a href="logout.asp">(�α׾ƿ�)</a>
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
					<td bgcolor="#E9EBF3">

					<!-- ���� START -->

					<table width="100%" border="0" cellspacing="0" cellpadding="0">
						<tr>
							<td height="5">
							<table width="100%" border="0" cellspacing="5" cellpadding="1" id="table34">
								<tr>
									<td bgcolor="#ffffff">
										
										<table border="0" cellpadding="0" style="padding: 4px" width="1300" id="table17" cellspacing="0" align="left">
										<tr>
											<td>
												
												<form id="Sform" name="Sform" action="list_up.asp" method="post" onsubmit="return SformSubmit();">
												<input type="hidden" name="Ps" value="Ps">
												<table width="1300"  border="0" cellpadding="1" cellspacing="1" bgcolor="#C4D2E9" id="table1" align="center">
													<tr align="center">	
														<td bgcolor="#ffffff" align="left"  height="25">
															&nbsp&nbsp;
															����:<select id="sKind" name="sKind">
																<option value="A" <%If sKind = "A" Then Response.Write "selected"%>>�Ű� ����Ʈ </option>
																<option value="B" <%If sKind = "B" Then Response.Write "selected"%>>OA ����Ʈ </option>
																<option value="C" <%If sKind = "C" Then Response.Write "selected"%>>�ؿܼ���</option>
															</select>&nbsp&nbsp;

															�μ�:<select id="sPart" name="sPart">
																<option value="A" <%If sPart = "A" Then Response.Write "selected"%>>���</option>
																<option value="B" <%If sPart = "B" Then Response.Write "selected"%>>ȭ��</option>
																<option value="C" <%If sPart = "C" Then Response.Write "selected"%>>���</option>
																<option value="D" <%If sPart = "D" Then Response.Write "selected"%>>����</option>
																<option value="E" <%If sPart = "E" Then Response.Write "selected"%>>��ǥ</option>
															</select>&nbsp&nbsp;

															����:<select id="sSection" name="sSection">
																<option value="A" <%If sSection = "A" Then Response.Write "selected"%>>���</option>
																<option value="B" <%If sSection = "B" Then Response.Write "selected"%>>����</option>
																<option value="C" <%If sSection = "C" Then Response.Write "selected"%>>��Ŀ��</option>
															</select>&nbsp&nbsp;

															�Ǹ�:<select id="sMethod" name="sMethod">
																<option value="A" <%If sMethod = "A" Then Response.Write "selected"%>>Ư��</option>
																<option value="B" <%If sMethod = "B" Then Response.Write "selected"%>>�ǿ�ž�</option>
																<option value="C" <%If sMethod = "C" Then Response.Write "selected"%>>������</option>
																<option value="D" <%If sMethod = "D" Then Response.Write "selected"%>>��ǥ</option>
																<option value="F" <%If sMethod = "F" Then Response.Write "selected"%>>���</option>
															</select>&nbsp&nbsp;

															ó��:<select id="sHandling" name="sHandling">
																<option value="A" <%If sHandling = "A" Then Response.Write "selected"%>>���</option>
																<option value="B" <%If sHandling = "B" Then Response.Write "selected"%>>��ó��</option>
															</select>&nbsp&nbsp;

															�����ϰ˻�: <input name="SearchStartDate" type="Date" value="<%=SearchStartDate%>" size="10"> ~ <input name="SearchEndDate" type="Date" value="<%=SearchEndDate%>"  size="10"> 
															&nbsp&nbsp;
															�˻�����:<input type="text" name="CmtTitle" size="20" />
														</td>
														<td  bgcolor="#ffffff" align="center"   width="150" rowspan="2" style='padding: 5px;'>
															<input type="submit" id="btnSearch" value=" �˻��ϱ� ">
														</td>
													</tr>
												</table>
												</form>
											</td>
										</tr>
										</table>
								</td>
							</tr>

							<%
							If sKind = "A" Then
								Call sbNewCase()
								Cmt = "* �߰����ӻ�� ����, TIBO/ JNJN/ JANS/ .. �� ���� Ÿ�ҹ��� ����� ����������� ���� ������ ������ ����"
							ElseIf sKind ="B" Then 
								Call sbOACase()
								Cmt = "* ���Ĺ��: �μ�-�����-����Due ��"
							ElseIf sKind ="C" Then 
								Call sbRLetter()
								Cmt = "* �ּ�- CD������:������ ������, CD��:��������, AD������:�ؿܴ븮�� ���ø�����, AD��:�ؿܴ븮�� ������, OD��:��Ǳ���"
							End If
							%>
							<td align="left"><span style="font-weight: bold;"><%=Cmt%></span></td>
							</table>
							</td>
						</tr>						
					</table>

					<!-- ���� END -->
							
							</td>
						</tr>
						<tr>
							<td>&nbsp;</td>
						</tr>
					</table>
							</td>
						</tr>
					</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>

</body>
</html>
<%
oConn.Close
Set oConn = Nothing
'------------------------------------------------------------------------------------------------------------------------
%>