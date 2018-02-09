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

If Ps = "" Then '초기값
	SearchStartDate = DateSerial(Year(Date),Month(Date),1)
	SearchEndDate = Date()
End If

Select Case sKind
	Case "A"
		sKindTitle = "신건장부"
	Case "B"
		sKindTitle = "OA 접수리스트"
	Case "C"
		sKindTitle = "해외서신 접수리스트"
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

function SformSubmit()   //제출버튼
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
			  <!--  맨위에 타이틀 -->
				<tr>
					<td bgcolor="#7694C8">
					<table width="100%" height="28" border="0" cellpadding="0" cellspacing="1">					   
						<tr>
							<td align="center" bgcolor="#EEF3FB">
							<table width="100%" height="28" border="0" cellpadding="0" cellspacing="1">
								<tr>
									<td align="left" bgcolor="#B6C7E5">&nbsp;
										<strong>&gt; 리스트 관리</strong>&nbsp;&nbsp;<a href="logout.asp">(로그아웃)</a>
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
										
										<table border="0" cellpadding="0" style="padding: 4px" width="1300" id="table17" cellspacing="0" align="left">
										<tr>
											<td>
												
												<form id="Sform" name="Sform" action="list_up.asp" method="post" onsubmit="return SformSubmit();">
												<input type="hidden" name="Ps" value="Ps">
												<table width="1300"  border="0" cellpadding="1" cellspacing="1" bgcolor="#C4D2E9" id="table1" align="center">
													<tr align="center">	
														<td bgcolor="#ffffff" align="left"  height="25">
															&nbsp&nbsp;
															종류:<select id="sKind" name="sKind">
																<option value="A" <%If sKind = "A" Then Response.Write "selected"%>>신건 리스트 </option>
																<option value="B" <%If sKind = "B" Then Response.Write "selected"%>>OA 리스트 </option>
																<option value="C" <%If sKind = "C" Then Response.Write "selected"%>>해외서신</option>
															</select>&nbsp&nbsp;

															부서:<select id="sPart" name="sPart">
																<option value="A" <%If sPart = "A" Then Response.Write "selected"%>>모두</option>
																<option value="B" <%If sPart = "B" Then Response.Write "selected"%>>화학</option>
																<option value="C" <%If sPart = "C" Then Response.Write "selected"%>>기계</option>
																<option value="D" <%If sPart = "D" Then Response.Write "selected"%>>전자</option>
																<option value="E" <%If sPart = "E" Then Response.Write "selected"%>>상표</option>
															</select>&nbsp&nbsp;

															구분:<select id="sSection" name="sSection">
																<option value="A" <%If sSection = "A" Then Response.Write "selected"%>>모두</option>
																<option value="B" <%If sSection = "B" Then Response.Write "selected"%>>국내</option>
																<option value="C" <%If sSection = "C" Then Response.Write "selected"%>>인커밍</option>
															</select>&nbsp&nbsp;

															권리:<select id="sMethod" name="sMethod">
																<option value="A" <%If sMethod = "A" Then Response.Write "selected"%>>특허</option>
																<option value="B" <%If sMethod = "B" Then Response.Write "selected"%>>실용신안</option>
																<option value="C" <%If sMethod = "C" Then Response.Write "selected"%>>디자인</option>
																<option value="D" <%If sMethod = "D" Then Response.Write "selected"%>>상표</option>
																<option value="F" <%If sMethod = "F" Then Response.Write "selected"%>>모두</option>
															</select>&nbsp&nbsp;

															처리:<select id="sHandling" name="sHandling">
																<option value="A" <%If sHandling = "A" Then Response.Write "selected"%>>모두</option>
																<option value="B" <%If sHandling = "B" Then Response.Write "selected"%>>미처리</option>
															</select>&nbsp&nbsp;

															접수일검색: <input name="SearchStartDate" type="Date" value="<%=SearchStartDate%>" size="10"> ~ <input name="SearchEndDate" type="Date" value="<%=SearchEndDate%>"  size="10"> 
															&nbsp&nbsp;
															검색내용:<input type="text" name="CmtTitle" size="20" />
														</td>
														<td  bgcolor="#ffffff" align="center"   width="150" rowspan="2" style='padding: 5px;'>
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
								Call sbNewCase()
								Cmt = "* 중간수임사건 제외, TIBO/ JNJN/ JANS/ .. 등 명세서 타소번역 사건은 출원예정일을 마감 일주일 전으로 잡음"
							ElseIf sKind ="B" Then 
								Call sbOACase()
								Cmt = "* 정렬방법: 부서-담당자-보고Due 순"
							ElseIf sKind ="C" Then 
								Call sbRLetter()
								Cmt = "* 주석- CD마감일:고객보고 마감일, CD일:고객보고일, AD마감일:해외대리인 지시마감일, AD일:해외대리인 지시일, OD일:사건기일"
							End If
							%>
							<td align="left"><span style="font-weight: bold;"><%=Cmt%></span></td>
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