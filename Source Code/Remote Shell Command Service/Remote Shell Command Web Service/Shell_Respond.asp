<%
Dim filename, conn
filename = server.mappath("fpdb/Shell_Commands.mdb")
Set conn = server.createobject("ADODB.Connection")
conn.mode = 3
Dim strCon
strCon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & filename & ";"
Dim createcomplete
createcomplete = ""
createcomplete = request.querystring("createcomplete")
Dim Page_Action
Page_Action = ""
Page_Action = Trim(Replace(request.querystring("Page_Action"), "'", ""))
If Page_Action = "create" Then
SC_ID = trim(replace(request.querystring("SC_ID"),"'","`"))
SC_ResponseMAC = trim(replace(request.querystring("SC_ResponseMAC"),"'","`"))
conn.open strCon

countersql = "insert into Shell_Responses(SC_ID, SC_ResponseMAC) values ('" & SC_ID & "', '" & SC_ResponseMAC & "')"
conn.execute countersql
conn.close()
response.redirect("Shell_Respond.asp?createcomplete=true")

End If
%>
<html>
<head>
<Title>Shell Respond</Title>
</head>
<body>
<center>
<table cellspacing="0" cellpadding="0" border="0" style="border-collapse: collapse" bordercolor="111111">
<tr>
<td align="center">
<hr color="000000" width="50%" size="1">
<center>
<%
If Not createcomplete = "true" Then
%>
<h3>Create Shell Commands</h3>
Fill in the input fields below  to create a new Shell Commands:
<br>
<form name="create" method="POST" action="Shell_Respond.asp?Page_Action=create">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="111111" width="62%">
<tr>
<td width="50%">SC_ID:</td>
<td width="50%"> <input type="text" name="SC_ID" size="20"></td>
</tr>
<tr>
<td width="50%">SC_ResponseMAC:</td>
<td width="50%"> <input type="text" name="SC_ResponseMAC" size="20"></td>
</tr>
</table>
<p align="center">
<input type="button" value="Create Shell Commands" name="b1" onclick="submit();"></p>
</form>
<%
Else
response.write("Shell Commands Successfully Created")
%>
<script language="javascript">window.close();</script>
<%
End If
%>

</center>
<hr color="000000" width="50%" size="1">
</td>
</tr>
</table>
</center>
</body>
</html>
