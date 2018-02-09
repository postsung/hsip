<!-- #include file="include/dbcon.asp" -->
<!-- #include file="include/Session_chk.asp" -->
<!-- #include file="include/function_cmt.asp" -->
<%
Server.ScriptTimeOut = 6000

Dim h_style, g_style, gs_style, hs_style
hs_style = "border-width: 1px; border-left-style: solid;	border-right-style: solid;	border-top-style: solid; border-bottom-style: solid;"   '헤더앞셀 테두리
h_style =  "border-width: 1px; border-right-style: solid;	border-top-style: solid;	border-bottom-style: solid;"							'헤더뒷셀 테두리
gs_style = "border-width: 1px; border-left-style: solid;	border-right-style: solid;	border-bottom-style: solid;"							'내용앞셀 테두리
g_style =  "border-width: 1px; border-right-style: solid;	border-bottom-style: solid;"														'내용뒷셀 테두리

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

If Ps = "" Then '초기값
	SearchStartDate = Date()
	SearchEndDate = Date()
End If

Select Case sKind
	Case "A"
		sKindTitle = "보고문서관리"
	Case "B"
		sKindTitle = "특허청서류"
	Case "C"
		sKindTitle = "서신관리"
	Case "D"
		sKindTitle = "Due Date관리"
	Case "E"
		sKindTitle = "해외출원안내"
	Case "F"
		sKindTitle = "심사청구 마감일안내"
End Select

If (sKind = "A" Or sKind = "B" Or sKind = "C" Or sKind = "F" Or sKind = "G" ) then
	Select Case sMethod
		Case "A"
			sKindTitle = sKindTitle & "(접수/발송서류)"
		Case "B"
			sKindTitle = sKindTitle & "(접수서류)"
		Case "C"
			sKindTitle = sKindTitle & "(발송서류)"
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
function SformSubmit()   '제출버튼
{	
	document.getElementById("btnSearch").disabled = true;
	document.getElementById("btnSearch").value = " 검색중 ... ";
	return true;
}

</script>
</head>

<body text="#000000" bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" class="style3">
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td valign="top" bgcolor="#E9EBF3">
			<table width="100%" border="0" cellspacing="1" cellpadding="0">
			  <!--  맨위에 현재 경로를 보여주는 tr-->
				<tr>
					<td bgcolor="#7694C8">
					<table width="100%" height="28" border="0" cellpadding="0" cellspacing="1">					   
						<tr>
							<td align="center" bgcolor="#EEF3FB">
							<table width="100%" height="28" border="0" cellpadding="0" cellspacing="1">
								<tr>
									<td align="left" bgcolor="#B6C7E5">&nbsp;
										<strong><a href="index.asp">Main</a> &gt; 특허청/보고문서</strong>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<a href="logout.asp">(로그아웃)</a>
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

					<!-- 본문 START -->

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
																<option value="">---- 구 분 ----</option>
																<option value="A" <%If sKind = "A" Then Response.Write "selected"%>>보고문서 관리 </option>
																<option value="B" <%If sKind = "B" Then Response.Write "selected"%>>특허청서류</option>
																<option value="C" <%If sKind = "C" Then Response.Write "selected"%>>서신 관리</option>
																<option value="D" <%If sKind = "D" Then Response.Write "selected"%>>Due Date관리</option>
																<option value="E" <%If sKind = "E" Then Response.Write "selected"%>>해외출원관리</option>
																<option value="F" <%If sKind = "F" Then Response.Write "selected"%>>심사미청구 리스트</option>
															</select>&nbsp;

															부서:
															<select id="sPart" name="sPart">
																<option value="A" <%If sPart = "A" Then Response.Write "selected"%>>모두</option>
																<option value="B" <%If sPart = "B" Then Response.Write "selected"%>>화학</option>
																<option value="C" <%If sPart = "C" Then Response.Write "selected"%>>기계</option>
																<option value="D" <%If sPart = "D" Then Response.Write "selected"%>>전자</option>
																<option value="E" <%If sPart = "E" Then Response.Write "selected"%>>상표</option>
															</select>&nbsp;

															<select id="sMethod" name="sMethod">
																<option value="A" <%If sMethod = "A" Then Response.Write "selected"%>>모든문서</option>
																<option value="B" <%If sMethod = "B" Then Response.Write "selected"%>>접수문서</option>
																<option value="C" <%If sMethod = "C" Then Response.Write "selected"%>>발송문서</option>
															</select>&nbsp;

															검색기간: <input name="SearchStartDate" type="Date" value="<%=SearchStartDate%>" size="10"> ~ <input name="SearchEndDate" type="Date" value="<%=SearchEndDate%>"  size="10"> &nbsp;

															처리:
															<select id="sDeley" name="sDeley">
																<option value="A" <%If sDeley = "A" Then Response.Write "selected"%>>모두</option>
																<option value="B" <%If sDeley = "B" Then Response.Write "selected"%>>미처리</option>
																<option value="C" <%If sDeley = "C" Then Response.Write "selected"%>>지연</option>
															</select>&nbsp;

															검색문서:&nbsp; <input type="text" name="CmtTitle" size="18" />
														</td>
														<td  bgcolor="#ffffff" align="center"   width="100" rowspan="2" style='padding: 5px;'>
															<input type="submit" id="btnSearch" value="검색하기">
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
								Cmt = "[지연 및 미처리 서신 안내] " + CStr(Date())
							ElseIf sKind ="B" Then 
								Call sbCommentChk()
								Cmt = "* 주석"
							ElseIf sKind ="C" Then 
								Call sbLetter()
								Cmt = "* 주석"
							ElseIf sKind ="D" Then 
								Call sbDueDate()
								Cmt = "* 당일까지 처리해야 하는 사건 중, 현재까지 SOPI에 미처리로 표기된 업무입니다. 업무에 참고하시기 바랍니다." + "<br><br> [당일기일관리] " + CStr(Date()) + " 미처리 사건안내"
							ElseIf sKind ="E" Then 
								Call sbOutGoingList()
								Cmt = "* 주석"
							ElseIf sKind ="F" Then 
								Call sbRExamList()
								Cmt = "* 주석"
							End If
							%>

							<td align="left"><span style="font-weight: bold;"><%=Cmt%></span></td>
							</table>
							</td>
						</tr>						
					</table>
					<!-- 본문 END -->			
</table>
</body>
</html>
<%
oConn.Close
Set oConn = Nothing
'------------------------------------------------------------------------------------------------------------------------
%>