<%@ Language = "VBScript" CodePage=949%>
<!-- #include file="include/dbcon.asp" -->
<!-- #include file="include/Session_chk.asp" -->
<!-- #include file="include/function.asp" -->

<%
Server.ScriptTimeout = 7600

Dim Sql, Rs, i, k
Dim arCustomer
Dim sPaper, sMemberID, CustomerCode
Dim StartYear, StartMonth, EndYear, EndMonth, EndDay, sNKind
Dim StartDate, EndDate, TodayEng, DuedateEng, FdateEng

Dim Fs

sNKind = Request("sNKind")
CustomerCode = Request("CustomerCode")

TodayEng = fnDateReplaceEnd(Date)
%>

'�������� ����
'-----------------------------------------------------------------------------------------------------------
<html xmlns:v="urn:schemas-microsoft-com:vml"
xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:w="urn:schemas-microsoft-com:office:word"
xmlns:m="http://schemas.microsoft.com/office/2004/12/omml"
xmlns="http://www.w3.org/TR/REC-html40">

<head>

<style>
<!--
 /* Font Definitions */
 @font-face
	{font-family:"Cambria Math";
	panose-1:2 4 5 3 5 4 6 3 2 4;
	mso-font-charset:0;
	mso-generic-font-family:roman;
	mso-font-pitch:variable;
	mso-font-signature:3 0 0 0 1 0;}
@font-face
	{font-family:"Malgun Gothic";
	panose-1:2 11 5 3 2 0 0 2 0 4;
	mso-font-charset:129;
	mso-generic-font-family:modern;
	mso-font-pitch:variable;
	mso-font-signature:-1879048145 701988091 18 0 524289 0;}

 /* Style Definitions */
 p.MsoNormal, li.MsoNormal, div.MsoNormal
	{mso-style-unhide:no;
	mso-style-qformat:yes;
	mso-style-parent:"";
	margin:0cm;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:"Times New Roman",serif;
	mso-fareast-font-family:"Malgun Gothic";
	mso-fareast-theme-font:minor-fareast;}
p.msonormal0, li.msonormal0, div.msonormal0
	{mso-style-name:msonormal;
	mso-style-unhide:no;
	mso-margin-top-alt:auto;
	margin-right:0cm;
	mso-margin-bottom-alt:auto;
	margin-left:0cm;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:"Times New Roman",serif;
	mso-fareast-font-family:"Malgun Gothic";
	mso-fareast-theme-font:minor-fareast;}
.MsoChpDefault
	{mso-style-type:export-only;
	mso-default-props:yes;
	mso-bidi-font-size:10.0pt;
	mso-ascii-font-family:"Times New Roman";
	mso-hansi-font-family:"Times New Roman";
	mso-font-kerning:0pt;}
@page WordSection1
	{size:595.3pt 841.9pt;
	margin:72.0pt 45.0pt 45.0pt 45.0pt;
	mso-header-margin:42.55pt;
	mso-footer-margin:40.0pt;
	mso-paper-source:0;}
div.WordSection1
	{page:WordSection1;}
-->
</style>

</head>

<body lang=KO style='tab-interval:40.0pt'>

<div class=WordSection1>
<p class=MsoNormal align=center style='text-align:center'><b><span lang=EN-US style='font-size:16.0pt'>Status of Patents / Applications</span></b></p>
<p class=MsoNormal><span lang=EN-US style='mso-fareast-font-family:"Times New Roman"'><o:p>&nbsp;</o:p></span></p>
<p class=MsoNormal align=center style='text-align:center'><span lang=EN-US>&nbsp;<o:p></o:p></span></p>

<table class=MsoNormalTable border=1 cellspacing=0 cellpadding=0 width=671 style='width:503.25pt;border-collapse:collapse;border:none;mso-border-alt: solid windowtext .5pt;mso-yfti-tbllook:1184;mso-padding-alt:0cm 0cm 0cm 0cm'>
 <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;height:16.9pt'>
  <td width=400 colspan=4 style='width:300.05pt;border:none;border-bottom:solid windowtext 1.0pt; mso-border-bottom-alt:solid windowtext .5pt;padding:0cm 0cm 0cm 0cm; height:16.9pt'>
  <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto'><span lang=EN-US style='font-size:10.0pt'></span></p>
  </td>
  <td width=271 colspan=3 style='width:203.2pt;border:none;border-bottom:solid windowtext 1.0pt;  mso-border-bottom-alt:solid windowtext .5pt;padding:0cm 0cm 0cm 0cm;  height:16.9pt'>
  <p class=MsoNormal align=right style='mso-margin-top-alt:auto;mso-margin-bottom-alt: auto;text-align:right'><span lang=EN-US style='font-size:10.0pt'><%=TodayEng%></span></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:1;height:22.05pt'>
  <td width=22 style='width:16.7pt;border:solid windowtext 1.0pt;border-top: none;mso-border-top-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt; padding:0cm 0cm 0cm 0cm;height:22.05pt'>
  <p class=MsoNormal align=center style='text-align:center'><b style='mso-bidi-font-weight:normal'><span lang=EN-US style='font-size:8.0pt'>No.<o:p></o:p></span></b></p>  
  </td>
  <td width=93 style='width:69.75pt;padding:0cm 0cm 0cm 0cm;height:22.05pt'>
  <p class=MsoNormal align=center style='text-align:center'><b style='mso-bidi-font-weight:normal'><span lang=EN-US style='font-size:8.0pt'>Application No.<o:p></o:p></span></b></p>
  </td>
  <td width=125 style='width:114.0pt;padding:0cm 0cm 0cm 0cm;height:22.05pt'>
  <p class=MsoNormal align=center style='text-align:center'><b style='mso-bidi-font-weight:normal'><span lang=EN-US style='font-size:8.0pt'>Your Ref.<o:p></o:p></span></b></p>
  </td>
  <td width=103 style='width:56.9pt;padding:0cm 0cm 0cm 0cm;height:22.05pt'>
  <p class=MsoNormal align=center style='text-align:center'><b style='mso-bidi-font-weight:normal'><span lang=EN-US style='font-size:8.0pt'>Our Ref.<o:p></o:p></span></b></p>
  </td>
  <td width=82 style='width:61.6pt;padding:0cm 0cm 0cm 0cm;height:22.05pt'>
  <p class=MsoNormal align=center style='text-align:center'><b style='mso-bidi-font-weight:normal'><span lang=EN-US style='font-size:8.0pt'>Patent No.<o:p></o:p></span></b></p>
  </td>
  <td width=82 style='width:61.6pt;padding:0cm 0cm 0cm 0cm;height:22.05pt'>
  <p class=MsoNormal align=center style='text-align:center'><b style='mso-bidi-font-weight:normal'><span lang=EN-US style='font-size:8.0pt'>Applicant(s)<o:p></o:p></span></b></p>
  </td>
  <td width=246 style='width:184.3pt;padding:0cm 0cm 0cm 0cm;height:22.05pt'>
  <p class=MsoNormal align=center style='text-align:center'><b style='mso-bidi-font-weight:normal'><span lang=EN-US style='font-size:8.0pt'>Present status<o:p></o:p></span></b></p>
  </td>

<%
'�������
'-----------------------------------------------------------------------
Dim ListCnt, Customer, AbandonDate, FilingDueDate
Dim ApplNo, YourRef, OurRef, RegNum, Applicant, DDate, AState, ADState


Sql = "SELECT Field34, Field42, Field6, Field5, Field74, Field37, AutoStateDate, AutoState, AutoDetailState, Field85, Field28 "
Sql = Sql & "FROM LeftMenu0001 "
Sql = Sql & "WHERE  Field85 is Null "

Select Case sNKind
	Case "A" '�������
		Sql = Sql & "AND PatCode in ('A01','A02', 'A03', 'A04') "
	Case "B" '�ܱ����
		Sql = Sql & "AND PatCode in ('A05','A06', 'A07', 'A08') "
End Select

If CustomerCode <>"" Then Sql = Sql & "AND Field5 Like '%"&CustomerCode&"%' "  '���ڵ� �˻�

Sql = Sql & "ORDER BY Field42 "

'	response.write sql & "<br>"

Set Rs = oConn.Execute(Sql)
ListCnt = 1
Do Until Rs.EOF '��Ǽ���ŭ ����
	
	Customer =	Rs.Fields(0)
	ApplNo =	Rs.Fields(1)
	YourRef =	Rs.Fields(2)
	OurRef =	Rs.Fields(3)
	RegNum =	Left(Rs.Fields(4),10)
	Applicant =	Rs.Fields(5)
	DDate =		fnDateReplaceEnd(Rs.Fields(6))
	AState =	Rs.Fields(7)
	ADState =	Rs.Fields(8)
	AbandonDate		= fnDateReplaceEnd(Rs.Fields(9))
	FilingDueDate	= fnDateReplaceEnd(Rs.Fields(10))

	If ADState <> "" Then AState = ADState   '�ڵ���Ȳ�� ���� �� ����Ȳ ����

	If Astate = "�������� X" Or InStr(Astate,"��������") >0 Or InStr(Astate,"����������") >0 Then 
		Astate ="Registered <br> (Annuity has not been paid by Hansung)"
	ElseIf InStr(Astate,"����������") >0 Or InStr(Astate,"��������") >0 Then 
		Astate ="Annuity Due on "& DDate
	ElseIf Astate = "��ϸ��� ����" Or Astate = "��ϴ�� ��" Then 
		Astate ="Notice of allowance <br>(Registration due date: "& DDate &")"
	ElseIf InStr(Astate,"����") >0 Then 
		Astate ="Abandoned <br>(Instruction on: "& AbandonDate &")"
	ElseIf Astate = "�����Һ� ���� ��" Or Astate = "������" Then 
		Astate ="Appeal to IPT"
	ElseIf Astate = "����û�� ���� ����" Or Astate = "����������" Or InStr(Astate,"�������� ������� �ۼ� ��") >0 Then 
		Astate ="Final rejection <br>(Appeal due date: "& DDate &")"
	ElseIf Astate = "�ǰ�����������" Or InStr(Astate,"�ǰ�����������") >0 Or InStr(Astate,"OA ������� �ۼ� ��") >0 Then 
		Astate ="Office Action<br>(Due Date: " & DDate &")"
	ElseIf Astate = "" Or Astate = "�ɸ� ��" Or InStr(Astate,"�ɻ� ��") >0 Or Astate = "��ɻ� ��" Or InStr(Astate,"�ɻ���") >0 Then 
		Astate ="Under Examination"
	ElseIf Astate = "�ɻ� �� û�� ��" Or  Astate = "�ɻ��û��"  Or InStr(Astate,"�ɻ�û�� ����") >0 Then 
		Astate ="No request for examination<br>(Due date: " & DDate &")"
	ElseIf Astate = "������� ��" Or Astate = "��� �غ� ��" Then 
		Astate ="Filing Due date: " & FilingDueDate
	End If
	Rs.MoveNext

%>
 <tr>
	<td style='height: 22pt'><p class=MsoNormal align=center style='text-align:center'><span lang=EN-US style='font-size:7pt;line-height:100%;font-family:"Times New Roman",serif'><%=ListCnt%></p></span></td>
	<td style='height: 22pt'><p class=MsoNormal align=center style='text-align:center'><span lang=EN-US style='font-size:7pt;line-height:100%;font-family:"Times New Roman",serif'><%=ApplNo%></p></span></td>
	<td style='height: 22pt'><p class=MsoNormal align=center style='text-align:left'><span lang=EN-US style='font-size:7pt;line-height:100%;font-family:"Times New Roman",serif'>&nbsp;<%=YourRef%></p></span></td>
	<td style='height: 22pt'><p class=MsoNormal align=center style='text-align:center'><span lang=EN-US style='font-size:7pt;line-height:100%;font-family:"Times New Roman",serif'><%=OurRef%></p></span></td>
	<td style='height: 22pt'><p class=MsoNormal align=center style='text-align:center'><span lang=EN-US style='font-size:7pt;line-height:100%;font-family:"Times New Roman",serif'><%=RegNum%></p></span></td>
	<td style='height: 22pt'><p class=MsoNormal align=center style='text-align:left'><span lang=EN-US style='font-size:7pt;line-height:100%;font-family:"Times New Roman",serif'>&nbsp;<%=Applicant%></p></span></td>
	<td style='height: 22pt'><p class=MsoNormal align=center style='text-align:center'><span lang=EN-US style='font-size:7pt;line-height:100%;font-family:"Times New Roman",serif'><%=AState%></p></span></td>
</tr>
<%

ListCnt = ListCnt + 1

Loop

Rs.Close
Set Rs = Nothing	

oConn.Close
Set oConn = Nothing

Response.AddHeader "Content-Disposition","attachment; filename=unpaidlist.xls"
Response.ContentType = "application/vnd.ms-excel" 

'���볡
'-----------------------------------------------------------------------
%>
 
</table>

<p class=MsoNormal><span lang=EN-US style='font-size:9pt;bold; mso-fareast-font-family:"Times New Roman"'><o:p>HANSUNG Intellectual Property</o:p></span></p>

</div>

</body>

</html>