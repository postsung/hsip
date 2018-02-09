<!-- #include file="include/dbcon.asp" -->
<!-- #include file="include/Session_chk.asp" -->
<!-- #include file="include/function_appl.asp" -->
<%
Server.ScriptTimeOut = 6000

Dim h_style, g_style, gs_style, hs_style
hs_style = "border-width: 1px; border-left-style: solid;	border-right-style: solid;	border-top-style: solid; border-bottom-style: solid;"   '헤더앞셀 테두리
h_style =  "border-width: 1px; border-right-style: solid;	border-top-style: solid;	border-bottom-style: solid;"							'헤더뒷셀 테두리
gs_style = "border-width: 1px; border-left-style: solid;	border-right-style: solid;	border-bottom-style: solid;"							'내용앞셀 테두리
g_style =  "border-width: 1px; border-right-style: solid;	border-bottom-style: solid;"														'내용뒷셀 테두리

Dim sKind, sNation, StartYear, StartMonth, EndYear, EndMonth, Ps, CCode
Dim Sql, Rs, Rs2, i, k, m
Dim StartDate, EndDate, LastDay, sKindTitle

Ps = Request("Ps")
sKind = Request("sKind")
sNation = Request("sNation")
StartYear = Int(Request("StartYear"))
StartMonth = Int(Request("StartMonth"))
EndYear = Int(Request("EndYear"))
EndMonth = Int(Request("EndMonth"))
CCode = Request("CCode")

If Ps = "" Then '초기값
	StartYear = Year(Date)
	StartMonth = 1
	EndYear = Year(Date)
	EndMonth = Month(Date)
End If

Select Case sKind
	Case "A"
		sKindTitle = "의뢰인/년도별 출원집계"
	Case "B"
		sKindTitle = "의뢰인/년도별 접수집계"
End Select

Select Case sNation
	Case "A"
		sKindTitle = sKindTitle & " (국내)"
	Case "B"
		sKindTitle = sKindTitle & " (외국)"
	Case "C"
		sKindTitle = sKindTitle & " (국내/외국)"
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
	if (document.getElementById("sKind").value == "")
	{
		alert("구분을 선택하여 주세요.");
		return false;
	}	
	if (document.getElementById("sNation").value == "")
	{
		alert("국가 구분을 선택하여 주세요.");
		return false;
	}	
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
										<strong><a href="index.asp">Main</a> &gt; 국내사건 집계</strong>
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
												
												<form id="Sform" name="Sform" action="appl_cnt_custom.asp" method="post" onsubmit="return SformSubmit();">
												<input type="hidden" name="Ps" value="Ps">
												<table width="1000"  border="0" cellpadding="1" cellspacing="1" bgcolor="#C4D2E9" id="table1" align="center">
													<tr align="center">	
														<td bgcolor="#ffffff" align="left"  height="25">
															<select id="sKind" name="sKind">
																<option value="">--- 구 분 ---</option>
																<option value="A" <%If sKind = "A" Then Response.Write "selected"%>>의뢰인/년도별 출원</option>
																<option value="B" <%If sKind = "B" Then Response.Write "selected"%>>의뢰인/년도별 접수</option>
															</select>&nbsp&nbsp&nbsp;
																국가구분:&nbsp;
															<select id="sNation" name="sNation">
																<option value="">-- 국가구분 --</option>
																<option value="A" <%If sNation = "A" Then Response.Write "selected"%>>내국</option>
																<option value="B" <%If sNation = "B" Then Response.Write "selected"%>>외국</option>
																<option value="C" <%If sNation = "C" Then Response.Write "selected"%>>모두</option>
															</select>&nbsp&nbsp&nbsp;
																고객코드:&nbsp; <input type="text" name="CCode" size="6" />
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

								Call sbApplCnt()

							ElseIf sKind ="B" Then 

								Call sbRcvCnt()

							End if
							%>

							<tr>
								<td>
									<table border="0" cellpadding="3" cellspacing="0" align="left" style="width: 1000px">
										<tr>
											<td align="left"><span style="font-weight: bold;">* 중간수임사건 제외(DOW는 포함)</span></td>
										</tr>
									</table>									
								</td>
							</tr>
								
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
Private Function fnClientOrder() '정렬쿼리
	Dim ReturnValue
	ReturnValue = "ORDER BY CASE "

	ReturnValue = ReturnValue & "WHEN ClientRef = 'JOKD' OR ClientRef = 'KCCC' OR ClientRef = 'LGCR' OR ClientRef = 'POSC' OR ClientRef = 'SAMY' OR ClientRef = 'SYGC' THEN 1 "
	ReturnValue = ReturnValue & "WHEN Nation = 'KR' THEN 2 "
	ReturnValue = ReturnValue & "WHEN ClientRef = 'BAYG' OR ClientRef = 'BAYM' OR ClientRef = 'BAYS' OR ClientRef = 'BAYC' OR ClientRef = 'BAYH' OR ClientRef = 'BAYA' OR ClientRef = 'BAYX' OR ClientRef = 'BAYB' THEN 10 "
	ReturnValue = ReturnValue & "WHEN ClientRef = 'JNJN' THEN 11 "
	ReturnValue = ReturnValue & "WHEN ClientRef = 'JANS' THEN 12 "
	ReturnValue = ReturnValue & "WHEN ClientRef = 'TIBO' THEN 13 "
	ReturnValue = ReturnValue & "WHEN ClientRef = 'RHEM' THEN 14 "
	ReturnValue = ReturnValue & "WHEN ClientRef = 'RHCO' THEN 15 "
	ReturnValue = ReturnValue & "WHEN ClientRef = 'RHCH' THEN 16 "
	ReturnValue = ReturnValue & "WHEN ClientRef = 'RHCA' THEN 17 "
	ReturnValue = ReturnValue & "WHEN ClientRef = 'BERG' THEN 18 "
	ReturnValue = ReturnValue & "WHEN ClientRef = 'ALCN' OR ClientRef = 'NOVA' THEN 19 "
	ReturnValue = ReturnValue & "WHEN Nation = 'JP' THEN 21 "

	ReturnValue = ReturnValue & "WHEN Nation = 'AU' OR Nation = 'CN' OR Nation = 'HK' OR Nation = 'ID' OR Nation = 'IN' OR Nation = 'MY' OR Nation = 'NP' OR Nation = 'PH' OR Nation = 'SA' OR Nation = 'SG' OR Nation = 'TH' OR Nation = 'TW' OR Nation = 'VN' OR Nation = 'NZ' THEN 22 "

	ReturnValue = ReturnValue & "WHEN Nation = 'AT' OR Nation = 'BE' OR Nation = 'CH' OR Nation = 'CZ' OR Nation = 'DZ' OR Nation = 'DK' OR Nation = 'DE' OR Nation = 'ES' OR Nation = 'FI' OR Nation = 'FR' OR Nation = 'GB' OR Nation = 'HU' OR Nation = 'IE' OR Nation = 'IS' OR Nation = 'IT' OR Nation = 'IL' OR Nation = 'LT' OR Nation = 'LU' OR Nation = 'NL' OR Nation = 'NO' OR Nation = 'PL' OR Nation = 'PT' OR Nation = 'RO' OR Nation = 'RS' OR Nation = 'RU' OR Nation = 'SE' OR Nation = 'SI' OR Nation = 'SJ' OR Nation = 'SK' OR Nation = 'SU' OR Nation = 'TR' OR Nation = 'UA' THEN 23 "
	
	ReturnValue = ReturnValue & "WHEN Nation = 'US' THEN 24 "
	
	ReturnValue = ReturnValue & "WHEN Nation = 'CA' OR Nation = 'BR' OR Nation = 'CL' OR Nation = 'CO' OR Nation = 'CU' OR Nation = 'DM' OR Nation = 'GD' OR Nation = 'MX' OR Nation = 'PE' OR Nation = 'PY' OR Nation = 'UY' OR Nation = 'VE' THEN 25 "
	
	ReturnValue = ReturnValue & "WHEN Nation = 'AO' OR Nation = 'CD' OR Nation = 'CF' OR Nation = 'CG' OR Nation = 'CM' OR Nation = 'ET' OR Nation = 'GA' OR Nation = 'GH' OR Nation = 'NG' OR Nation = 'SN' OR Nation = 'SO' OR Nation = 'SZ' OR Nation = 'TG' OR Nation = 'TZ' OR Nation = 'UG' OR Nation = 'ZA' OR Nation = 'ZM' OR Nation = 'ZR' OR Nation = 'ZW' THEN 26 "
	
	ReturnValue = ReturnValue & "ELSE 100 "
	ReturnValue = ReturnValue & "END "

	fnClientOrder = ReturnValue
End Function
'------------------------------------------------------------------------------------------------------------------------
%>