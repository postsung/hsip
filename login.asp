<%
Dim MemberPWD, Ps

MemberPWD = Request("MemberPWD")
Ps = Request("Ps")

If Ps = "Ps" Then
	If MemberPWD = "!hs8799web" Then		'���������� �α���
		Session("MemberID") = "admin"
		Response.Redirect "index.asp"
	ElseIf MemberPWD = "5555" Then			'��ǥ��(�Ű����, OA���, �ؿ���� ��������)
		Session("MemberID") = "admin"
		Response.Redirect "list_up.asp"
	ElseIf MemberPWD = "6997" Then			'��������� ������
		Session("MemberID") = "admin"
		Response.Redirect "manage_comment.asp"
	ElseIf MemberPWD = "8103" Then			'�渮��
		Session("MemberID") = "admin"
		Response.Redirect "manage_invoice.asp"
	ElseIf MemberPWD = "6888" Then			'�輺��(��Ǹ���Ʈ/��û��������Ʈ)
		Session("MemberID") = "admin"
		Response.Redirect "find_case.asp"
	ElseIf MemberPWD = "1234" Then			'Ư��û�߰���������
		Session("MemberID") = "admin"
		Response.Redirect "manage_RevList.asp"
	Else
%>
		<script language="javascript">
		alert("��й�ȣ�� �ٽ� Ȯ���Ͽ� �ּ���.");
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
			<tr><td align="center"><b>Ư����� �Ѽ� (HANSUNG Intellectual Property)</b></td></tr><p />
			<form name="Sform" action="login.asp" method="post">
				<input type="hidden" name="Ps" value="Ps">
				��й�ȣ: <input type="password" name="MemberPWD" size="20">
				<input type="submit" value=" �� �� �� " style="height:22px;" >
	</form>
	</center>

 </body>
</html>
