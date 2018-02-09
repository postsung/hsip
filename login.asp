<%
Dim MemberPWD, Ps

MemberPWD = Request("MemberPWD")
Ps = Request("Ps")

If Ps = "Ps" Then
	If MemberPWD = "!hs8799web" Then		'관리페이지 로그인
		Session("MemberID") = "admin"
		Response.Redirect "index.asp"
	ElseIf MemberPWD = "5555" Then			'대표님(신건장부, OA장부, 해외출원 문서관리)
		Session("MemberID") = "admin"
		Response.Redirect "list_up.asp"
	ElseIf MemberPWD = "6997" Then			'서브관리자 페이지
		Session("MemberID") = "admin"
		Response.Redirect "manage_comment.asp"
	ElseIf MemberPWD = "8103" Then			'경리팀
		Session("MemberID") = "admin"
		Response.Redirect "manage_invoice.asp"
	ElseIf MemberPWD = "6888" Then			'김성은(사건리스트/심청마감리스트)
		Session("MemberID") = "admin"
		Response.Redirect "find_case.asp"
	ElseIf MemberPWD = "1234" Then			'특허청중간서류접수
		Session("MemberID") = "admin"
		Response.Redirect "manage_RevList.asp"
	Else
%>
		<script language="javascript">
		alert("비밀번호를 다시 확인하여 주세요.");
		</script>
<%
	End If
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<title> New Document </title>
		<meta name="Generator" content="EditPlus">
		<meta name="Author" content="SOPI">
		<meta name="Keywords" content="LOGIN">
		<meta name="Description" content="SOPI-HANSUNG">
	</head>
	<body>
		<center>
			<br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
			<tr><td align="center"><b>특허법인 한성 (HANSUNG Intellectual Property)</b></td></tr><p />
			<form name="Sform" action="login.asp" method="post">
				<input type="hidden" name="Ps" value="Ps">
				비밀번호: <input type="password" name="MemberPWD" size="20">
				<input type="submit" value=" 로 그 인 " style="height:22px;" >
	</form>
	</center>

 </body>
</html>
