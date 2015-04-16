<html>
<body>
<%
Dim filename, conn
filename=server.mappath("fpdb/Shell_Commands.mdb")
Set conn = server.createobject("ADODB.Connection")
conn.mode = 3
Dim strCon
strCon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & filename & ";"
conn.open(strCon)
Set objRec = server.createobject("ADODB.Recordset")
strSQL = "select * from Shell_Commands order by SC_ID"
objRec.ActiveConnection = conn
objRec.open(strSQL)
if not objRec.eof then
objRec.moveFirst
end if
%>
<p><b>SC_ID</b><br><b>SC_Command</b><br>
<% while not objRec.eof %>
<!--start command-->
<% response.write objRec.Fields("SC_ID").value %><br>
<% response.write objRec.Fields("SC_Command").value %><br>
<!--end command-->
<%
objRec.MoveNext
wend
%>
</p>
<%
objRec.close()
conn.close()
%>
</body>
</html>
