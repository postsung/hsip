<!-- #include file="include/dbcon.asp" -->
<!-- #include file="include/Session_chk.asp" -->
<!-- #include file="include/function.asp" -->
<!-- #include file="include/function_stat.asp" -->
<%
Server.ScriptTimeOut = 6000

Dim h_style, g_style, gs_style, hs_style
hs_style = "border-width: 0.5px; border-left-style: solid;	border-right-style: solid;	border-top-style: solid; border-bottom-style: solid;"   '����ռ� �׵θ�
h_style =  "border-width: 0.5px; border-right-style: solid;	border-top-style: solid;	border-bottom-style: solid;"							'����޼� �׵θ�
gs_style = "border-width: 0.5px; border-left-style: solid;	border-right-style: solid;	border-bottom-style: solid;"							'����ռ� �׵θ�
g_style =  "border-width: 0.5px; border-right-style: solid;	border-bottom-style: solid;"														'����޼� �׵θ�

Dim sKind, StartYear, StartMonth, EndYear, EndMonth, Ps
Dim Sql, Rs, Rs2, i, k, m
Dim StartDate, EndDate, LastDay, sKindTitle

Ps = Request("Ps")
sKind = Request("sKind")
StartYear = Int(Request("StartYear"))
StartMonth = Int(Request("StartMonth"))
EndYear = Int(Request("EndYear"))
EndMonth = Int(Request("EndMonth"))

If Ps = "" Then '�ʱⰪ
	StartYear = Year(Date)
	StartMonth = 1
	EndYear = Year(Date)
	EndMonth = Month(Date)
End If

Select Case sKind
	Case "E"
		sKindTitle = "������ " & StartMonth & "��~" & EndMonth & "�� �����Ǽ� ����"
	Case "A"
		sKindTitle = "4. ���� ����Ư�� �������"
	Case "B"
		sKindTitle = "3. ���� ȸ��Ư�� �������"
	Case "C"
		sKindTitle = "2. �ܱ� �븮�κ� Ư���������"
	Case "D"
		sKindTitle = "1. �ܱ� �ֿ�� Ư���������"
End Select
%>
<html>
<head>

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

<script  src="include/jquery.min.js"></script>

<script language="javascript">
function SformSubmit()
{	
	if (document.getElementById("sKind").value == "")
	{
		alert("������ �����Ͽ� �ּ���.");
		return false;
	}	
	document.getElementById("btnSearch").disabled = true;
	document.getElementById("btnSearch").value = " �˻��� ... ";
	return true;
}
function YearChanged(v)
{
	if (document.getElementById("StartYear").value != v)
	{
		document.getElementById("EndYear").value = document.getElementById("StartYear").value;
		alert("������ �⵵�� �˻��Ͽ� �ּ���.");		
		return;
	}
}
$.fn.rowspan = function(idx, isType) {       
    return this.each(function(){      
        var that;     
        $('tr', this).each(function(row) {      
			
            $('td:eq('+idx+')', this).filter(':visible').each(function(col) {                
                if ($(this).html() == $(that).html() && 
						( !isType || isType && $(this).prev().html() == $(that).prev().html() )
                    ) {            
                    rowspan = $(that).attr("rowspan") || 1;
                    rowspan = Number(rowspan)+1;
 
                    $(that).attr("rowspan",rowspan);
                     
                    $(this).hide();
                     
                } else {            
                    that = this;         
                }
                that = (that == null) ? this : that;      
            });     
        });    
    });  
}; 

</script>
</head>

<body text="#000000" bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td valign="top" bgcolor="#FFFFFF">
			<table width="100%" border="0" cellspacing="1" cellpadding="0">
			  <!--  ������ ���� ��θ� �����ִ� tr-->
				<tr>
					<td bgcolor="#7694C8">
					<table width="100%" height="28" border="0" cellpadding="0" cellspacing="1">					   
						<tr>
							<td align="center" bgcolor="#FFFFFF">
							<table width="100%" height="28" border="0" cellpadding="0" cellspacing="1">
								<tr>
									<td align="left" bgcolor="#B6C7E5">&nbsp;
										<strong><a href="index.asp">Main</a> &gt; ����������ǥ</strong>
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
					<td bgcolor="#FFFFFF">

					<!-- ���� START -->

					<table width="100%" border="0" cellspacing="0" cellpadding="0">
						<tr>
							<td height="5">
							<table width="100%" border="0" cellspacing="5" cellpadding="1" id="table34">
								<tr>
									<td bgcolor="#ffffff">
										
										<table border="0" cellpadding="0" style="padding: 4px" width="1000" id="table17" cellspacing="0" align="left">
										<tr>
											<td>
												
												<form id="Sform" name="Sform" action="appl_stat.asp" method="post" onsubmit="return SformSubmit();">
												<input type="hidden" name="Ps" value="Ps">
												<table width="1000"  border="0" cellpadding="1" cellspacing="1" bgcolor="#C4D2E9" id="table1" align="center">
													<tr align="center">	
														<td bgcolor="#ffffff" align="left"  height="25">
															<select id="sKind" name="sKind">
																<option value="">--- �� �� ---</option>
																<option value="E" <%If sKind = "E" Then Response.Write "selected"%>>����������ǥ</option>
																<option value="D" <%If sKind = "D" Then Response.Write "selected"%>>1.�ܱ��ֿ��</option>
																<option value="C" <%If sKind = "C" Then Response.Write "selected"%>>2.�ܱ���</option>
																<option value="B" <%If sKind = "B" Then Response.Write "selected"%>>3.��������</option>
																<option value="A" <%If sKind = "A" Then Response.Write "selected"%>>4.��������</option>
																<option value="F" <%If sKind = "F" Then Response.Write "selected"%>>5.�̺з���</option>
																<option value="G" <%If sKind = "G" Then Response.Write "selected"%>>6.����κ� �ڵ�</option>
															</select>&nbsp;

															<select>
																<option>�����</option>
															</select>&nbsp;

															<select id="StartYear" name="StartYear">
																<%
																For i = Year(Date)-3 To Year(Date)+1
																%>
																	<option value="<%=i%>" <%If StartYear = i Then Response.Write "selected"%>><%=i%></option>
																<%
																Next
																%>																
															</select>��&nbsp;
															<select id="StartMonth" name="StartMonth">
																<%
																For i = 1 To 12
																%>
																	<option value="<%=i%>" <%If StartMonth = i Then Response.Write "selected"%>><%=i%></option>
																<%
																Next
																%>																
															</select>��&nbsp;
															~
															<select id="EndYear" name="EndYear" onchange="javascript:YearChanged(this.value);">
																<%
																For i = Year(Date)-3 To Year(Date)+1
																%>
																	<option value="<%=i%>" <%If EndYear = i Then Response.Write "selected"%>><%=i%></option>
																<%
																Next
																%>																
															</select>��&nbsp;
															<select id="EndMonth" name="EndMonth">
																<%
																For i = 1 To 12
																%>
																	<option value="<%=i%>" <%If EndMonth = i Then Response.Write "selected"%>><%=i%></option>
																<%
																Next
																%>																
															</select>��

														</td>
														<td  bgcolor="#ffffff" align="center"   width="20%" rowspan="2" style='padding: 5px;'>
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
							If sKind = "E" Then	
								Call sbApplPartStat()
							ElseIf sKind = "G" Then
								Call sbApplicantCode()
							Else
								Call sbApplStat()
							End If
							%>

							<tr>
								<td>
									<table border="0" cellpadding="3" width="1000" cellspacing="0" align="left">
										<tr>
											<td align="left"><span style="font-weight: bold;">* �߰����ӻ�� ����(DOW�� ����)</span></td>
										</tr>
									</table>									
								</td>
							</tr>
								
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
Private Function fnClientOrder() '��������
	Dim ReturnValue
	ReturnValue = "ORDER BY CASE "

	ReturnValue = ReturnValue & "WHEN ClientRef = 'JOKD' OR ClientRef = 'LGCR' OR ClientRef = 'SAMY' OR ClientRef = 'SYGC' OR ClientRef = 'DWPH' OR ClientRef = 'APLD' OR ClientRef = 'KONI' OR ClientRef = 'KOOK' OR ClientRef = 'SJCN' OR ClientRef = 'AMPA' OR ClientRef = 'KONS' OR ClientRef = 'SEGO' OR ClientRef = 'SCNU' OR ClientRef = 'SKBP' OR ClientRef = 'SAMB' OR ClientRef = 'SAMK' THEN 1 "
	ReturnValue = ReturnValue & "WHEN Nation = 'KR' THEN 2 "
	ReturnValue = ReturnValue & "WHEN ClientRef = 'BAYG' OR ClientRef = 'BAYM' OR ClientRef = 'BAYS' THEN 10 "
	ReturnValue = ReturnValue & "WHEN ClientRef = 'JNJN' THEN 11 "
	ReturnValue = ReturnValue & "WHEN ClientRef = 'JANS' THEN 12 "
	ReturnValue = ReturnValue & "WHEN ClientRef = 'TIBO' OR ClientRef = 'TIBP' THEN 13 "
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
	ReturnValue = ReturnValue & "END, Nation "

	fnClientOrder = ReturnValue
End Function
'------------------------------------------------------------------------------------------------------------------------
%>