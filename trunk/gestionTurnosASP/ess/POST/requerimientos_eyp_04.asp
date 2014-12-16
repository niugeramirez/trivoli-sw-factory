<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/adovbs.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/asistente.inc"-->
<!--
Archivo: requerimientos_eyp_04.asp
Descripción: Abm de requerimientos
Autor : Lisandro Moro
Fecha: 19/04/2004
modificacion: 14/06/2004 - Lisandro Moro - Se agrego que borre las relaciones con las estructuras
Modificado  : 12/09/2006 Raul Chinestra - se agregó Requerimientos de Personal en Autogestión   
-->
<% 
on error goto 0
Dim l_cm
Dim l_rs
Dim l_sql
Dim l_reqpernro
	
l_reqpernro = request.querystring("reqpernro")

cn.BeginTrans

Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = "SELECT reqpernro "
l_sql = l_sql & " FROM pos_reqbus "
l_sql = l_sql & " WHERE  reqpernro = " & l_reqpernro
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
	l_rs.close
	set l_rs = nothing
	Response.write "<script>alert('Este Requerimiento está asociado a una Búsqueda. No se eliminó el registro.');window.close();</script>"
else	
	l_rs.close
	l_sql = "SELECT reqpernro, tenro "
	l_sql = l_sql & " FROM pos_reqestr"
	l_sql = l_sql & " WHERE  reqpernro = " & l_reqpernro
	rsOpenCursor l_rs, cn, l_sql, 0, adOpenKeySet
	set l_cm = Server.CreateObject("ADODB.Command")
	l_cm.activeconnection = Cn
	do until l_rs.eof
		l_sql = "DELETE FROM pos_reqestr "
		l_sql = l_sql & " WHERE reqpernro = " & l_reqpernro
		l_sql = l_sql & " AND tenro =  " & l_rs("tenro")
		cmExecute l_cm, l_sql, 0
		l_rs.MoveNext
	loop
   	l_rs.close
	
	l_sql = "DELETE FROM pos_reqpersonal WHERE reqpernro = " & l_reqpernro
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	
	call actualizarPasos(45, l_reqpernro, 0)
	call actualizarPasos(46, l_reqpernro, 0)
	call actualizarPasos(47, l_reqpernro, 0)
end if

cn.CommitTrans

'l_cm.close
cn.Close
Set cn = Nothing
Response.write "<script>alert('Operación Realizada.');"
Response.write "		window.opener.ifrm.location.reload();window.close();"
Response.write "</script>"
 	
%>
