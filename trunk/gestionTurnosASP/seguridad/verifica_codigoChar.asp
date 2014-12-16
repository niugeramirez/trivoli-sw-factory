<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Gesti&oacute;n de Tiempos - Heidt & Asociados S.A.</title>
<body>
<% 
dim l_rs
dim l_sql
dim l_codigo
dim l_campo_dest
dim l_campo_cod
dim l_tabla
dim l_resp
dim l_i

l_codigo     = request("codigo")
l_campo_dest = request("campodest")
l_campo_cod  = request("campocod")
l_tabla      = request("tabla")

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT " & l_campo_dest & " FROM " & l_tabla & " WHERE " & l_campo_cod & " = '" & l_codigo & "'"
'RS.Maxrecords = 1
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
  l_resp = l_rs(0)
  for l_i = 1 to l_rs.fields.count - 1
     l_resp = l_resp & ", " & l_rs(l_i)
  next
%>
<script> 			 
window.returnValue = '<%= l_resp %>';
window.close();
</script>
<% 
else 
%>
<script> 			 
window.returnValue = "";
window.close();
</script>
<% 
end if
l_rs.close
%>
</body>
</html>
