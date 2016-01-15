<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<% 


'on error goto 0
Dim l_cm
Dim l_rs
Dim l_sql
Dim l_id
	
l_id = request.Form("cabnro")
Set l_rs = Server.CreateObject("ADODB.RecordSet")
set l_cm = Server.CreateObject("ADODB.Command")
l_rs.close
l_sql = "SELECT id"
l_sql = l_sql & " FROM clientespacientes"
l_sql  = l_sql  & " WHERE idobrasocial = " & l_id
l_sql = l_sql & " and empnro = " & Session("empnro") 
 
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	Response.write "Existen pacientes asociados a la Obra Social. No se permite eliminar."
	l_rs.close
else
	l_rs.close
	l_sql = "SELECT id"
	l_sql = l_sql & " FROM pagos"
	l_sql  = l_sql  & " WHERE idobrasocial = " & l_id
	l_sql = l_sql & " and empnro = " & Session("empnro") 
	 
	rsOpen l_rs, cn, l_sql, 0 
	if not l_rs.eof then
		Response.write "Existen pagos asociados a la Obra Social. No se permite eliminar."
		l_rs.close
	else
		l_rs.close
		l_sql = "SELECT id"
		l_sql = l_sql & " FROM listaprecioscabecera"
		l_sql  = l_sql  & " WHERE idobrasocial = " & l_id
		l_sql = l_sql & " and empnro = " & Session("empnro") 
		 
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			Response.write "Existen Listas de precios asociados a la Obra Social. No se permite eliminar."
			l_rs.close
		else	
			l_sql = " DELETE FROM obrassociales WHERE id = " & l_id
			
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0

			cn.Close
			Set cn = Nothing

			Response.write "OK"
		end if
	end if
end if


%>

	




