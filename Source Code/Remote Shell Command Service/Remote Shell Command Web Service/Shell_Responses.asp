<html>
<body>
<% 
dim SC_ID
SC_ID = request.querystring("SC_ID")
%>
<p><b>SC_ID: <% response.write SC_ID %></b><br>
<%
Dim filename, conn
filename=server.mappath("fpdb/Shell_Commands.mdb")
Set conn = server.createobject("ADODB.Connection")
conn.mode = 3
Dim strCon
strCon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & filename & ";"
conn.open(strCon)
Set objRec = server.createobject("ADODB.Recordset")
strSQL = "select * from Shell_Responses where SC_ID = '" & SC_ID & "' order by SC_ResponseMAC"
objRec.ActiveConnection = conn
objRec.open(strSQL)
if not objRec.eof then
objRec.moveFirst
end if
 while not objRec.eof 
  response.write objRec.Fields("SC_ResponseMAC").value %><br>
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
