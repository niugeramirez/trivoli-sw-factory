<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<% 


'on error goto 0
Dim l_cm
Dim l_rs
Dim l_sql
Dim l_id
	
l_id = request.querystring("cabnro")
Set l_rs = Server.CreateObject("ADODB.RecordSet")
set l_cm = Server.CreateObject("ADODB.Command")
l_rs.close
'l_sql = "SELECT depnro"
'l_sql = l_sql & " FROM tkt_pro_dep"
'l_sql  = l_sql  & " WHERE depnro = " & l_depnro
'rsOpen l_rs, cn, l_sql, 0 
'if not l_rs.eof then
'	Response.write "<script>alert('Existen Productos asociados a este Depósito.\nNo es posible dar de baja.');window.close();</script>"
'else
	l_sql = " DELETE FROM obrassociales WHERE id = " & l_id
'end if
'l_rs.close

l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0

cn.Close
Set cn = Nothing

Response.write "OK"

%>

	




