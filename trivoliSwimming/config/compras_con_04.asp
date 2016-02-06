<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->

<% 
'Archivo: compras_con_04.asp
'Descripción: Script Baja Compras
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
l_sql = l_sql & " FROM cajaMovimientos"
l_sql  = l_sql  & " WHERE idcompraOrigen = " & l_id
rsOpen l_rs, cn, l_sql, 0 

if not l_rs.eof then
	Response.write "Existen pagos asociados a la compra. No se permite eliminar."
	l_rs.close
else
	l_sql = "DELETE FROM compras  WHERE id = " & l_id

	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0

	'Borro detalle compras
	l_sql = "DELETE FROM detalleCompras  WHERE idcompra = " & l_id

	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	
	cn.Close
	Set cn = Nothing
	
	Response.write "OK"
end if



%>