<!-- #include file="include/dbcon.asp" -->
<!-- #include file="include/Session_chk.asp" -->
<!-- #include file="include/function_appl.asp" -->
<%
Server.ScriptTimeOut = 12000

Dim h_style, g_style, gs_style, hs_style
hs_style = "border-width: 1px; border-left-style: solid;	border-right-style: solid;	border-top-style: solid; border-bottom-style: solid;"   '헤더앞셀 테두리
h_style =  "border-width: 1px; border-right-style: solid;	border-top-style: solid;	border-bottom-style: solid;"							'헤더뒷셀 테두리
gs_style = "border-width: 1px; border-left-style: solid;	border-right-style: solid;	border-bottom-style: solid;"							'내용앞셀 테두리
g_style =  "border-width: 1px; border-right-style: solid;	border-bottom-style: solid;"														'내용뒷셀 테두리

Dim sKind, OGNCode, StartYear, StartMonth, EndYear, EndMonth, Ps, OGAgent
Dim Sql, Rs, Rs2, i, k, m
Dim StartDate, EndDate, LastDay, sKindTitle

Ps = Request("Ps")
sKind = Request("sKind")
StartYear = Int(Request("StartYear"))
StartMonth = Int(Request("StartMonth"))
EndYear = Int(Request("EndYear"))
EndMonth = Int(Request("EndMonth"))
OGNCode = Request("OGNCode")
OGAgent = Request("OGAgent")

If Ps = "" Then '초기값
	StartYear = Year(Date)
	StartMonth = 1
	EndYear = Year(Date)
	EndMonth = Month(Date)
End If

Select Case sKind
	Case "A"
		sKindTitle = "해외대리인/년도별 출원집계"
	Case "B"
		sKindTitle = "해외대리인/년도별 접수집계"
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
.style4 {color: #000000}
-->
</style>

<script language="javascript">
function SformSubmit()
{	
	document.getElementById("btnSearch").disabled = true;
	document.getElementById("btnSearch").value = " 검색중 ... ";
	return true;
}
</script>
</head>

<body text="#000000" bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" class="style4">
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
										<strong><a href="index.asp">Main</a> &gt; 해외사건 집계</strong>
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
					<td bgcolor="#E9EBF3">

					<!-- 본문 START -->

					<table width="100%" border="0" cellspacing="0" cellpadding="0">
						<tr>
							<td height="5">
							<table width="100%" border="0" cellspacing="5" cellpadding="1" id="table34">
								<tr>
									<td bgcolor="#ffffff">
										
										<table border="0" cellpadding="0" style="padding: 4px" width="1000" id="table17" cellspacing="0" align="left">
										<tr>
											<td>
												
												<form id="Sform" name="Sform" action="appl_cnt_OGcustom.asp" method="post" onsubmit="return SformSubmit();">
												<input type="hidden" name="Ps" value="Ps">
												<table width="1000"  border="0" cellpadding="1" cellspacing="1" bgcolor="#C4D2E9" id="table1" align="center">
													<tr align="center">	
														<td bgcolor="#ffffff" align="left"  height="25">
															<select id="sKind" name="sKind">
																<option value="">--- 구 분 ---</option>
																<option value="A" <%If sKind = "A" Then Response.Write "selected"%>>해외대리인/년도별 출원</option>
																<option value="B" <%If sKind = "B" Then Response.Write "selected"%>>해외대리인/년도별 접수</option>
															</select>&nbsp&nbsp&nbsp;
																국가코드:&nbsp; <input type="text" name="OGNCode" size="6" />
																해외대리인:&nbsp; <input type="text" name="OGAgent" size="20" />
														</td>
														<td  bgcolor="#ffffff" align="center"   width="20%" rowspan="2" style='padding: 5px;'>
															<input type="submit" id="btnSearch" value=" 검색하기 ">
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

								Call sbOGApplCnt()

							ElseIf sKind ="B" Then 

								Call sbOGRcvCnt()

							End if
							%>					
							</table>
							</td>
						</tr>						
					</table>

					<!-- 본문 END -->
							
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