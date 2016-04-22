<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->

<% 
'Archivo: bancos_con_04.asp
'Descripción: Script Baja bancos
'Autor : Trivoli
'Fecha: 31/05/2015

'on error goto 0
Dim l_cm
Dim l_rs
Dim l_sql
Dim l_id
	
'l_id = request.querystring("cabnro")
l_id = request.Form("cabnro")

Set l_rs = Server.CreateObject("ADODB.RecordSet")
Set l_cm = Server.CreateObject("ADODB.Command")

l_sql = "SELECT id"
l_sql = l_sql & " FROM cheques "
l_sql  = l_sql  & " WHERE id_banco = " & l_id
rsOpen l_rs, cn, l_sql, 0 

if not l_rs.eof then
	Response.write "Existen Cheques asociados al Banco. No se permite eliminar."
	l_rs.close
else
	l_sql = "DELETE FROM bancos  WHERE id = " & l_id

	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0

	cn.Close
	Set cn = Nothing
	
	Response.write "OK"
end if



%>