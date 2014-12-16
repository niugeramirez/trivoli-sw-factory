<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<%
' variable base de datos 
  Dim l_cm
  Dim l_sql
  Dim l_rs
  
' parametro entrada
  Dim l_evaevenro
  l_evaevenro  = Request.QueryString("evaevenro")

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT * FROM  evacab WHERE  evacab.evaevenro = " & l_evaevenro
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then
	if cint(ccodelco) = -1 then
		set l_cm = Server.CreateObject("ADODB.Command")
		l_sql = "DELETE FROM evatipoobjpor WHERE evaevenro = " & l_evaevenro
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
		
		set l_cm = Server.CreateObject("ADODB.Command")
		l_sql = "DELETE FROM evaobjinicial WHERE evaevenro = " & l_evaevenro
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0	
	end if
	
	set l_cm = Server.CreateObject("ADODB.Command")
	l_sql = "DELETE FROM evaevento WHERE evaevenro = " & l_evaevenro
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location.reload();window.close();</script>"
else
	Response.write "<script>alert('Existen evaluaciones para este Evento. No se elimino el Registro');window.opener.ifrm.location.reload();window.close();</script>"	
end if	
l_rs.Close
set l_rs=nothing
	
%>
