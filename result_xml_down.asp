<!-- #include file="include/dbcon.asp" -->
<!-- #include file="include/Session_chk.asp" -->
<!-- #include file="include/function.asp" -->
<%
Session.CodePage = 949

oConn.Close
Set oConn = Nothing

Dim n, f
n = Request("n")
f = Request("f")
Select Case f
	Case "A"
		f = "�μ�����������ǥ"
	Case "B"
		f = "���κ���������ǥ"
	Case "C"
		f = "û����������ǥ"
	Case "D"
		f = "������������ǥ"
End Select

Response.ContentType = "application/unknown"
Response.AddHeader "Content-Disposition","attachment; filename=" & f & ".doc"

Dim objStream
Set objStream = Server.CreateObject("ADODB.Stream")
objStream.Open
objStream.Type = 1
objStream.LoadFromFile Server.Mappath("./") & "\Temp\" & n & ".doc"

Response.BinaryWrite objStream.Read 

Set objstream = nothing 
%>
