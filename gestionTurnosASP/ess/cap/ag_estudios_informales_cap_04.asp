<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
Archivo: ag_estudios_informales_cap_04.asp
Descripción: Abm de Estudios Informales. Eliminar.
Autor : Lisandro Moro
Fecha: 29/03/2004
-->
<% 
Dim l_cm
Dim l_rs
Dim l_sql
Dim l_estinfnro
	
	l_estinfnro = request.querystring("cabnro")
	
'	Set l_rs = Server.CreateObject("ADODB.RecordSet")
'	l_sql = "SELECT arenro"
'	l_sql = l_sql & " FROM cap_tema"
'	l_sql = l_sql & " WHERE arenro= " & l_arenro
'	rsOpen l_rs, cn, l_sql, 0
'	if not l_rs.eof then
'		l_rs.close
'		set l_rs = nothing
'		Response.write "<script>alert('Este Area está asociado a un Contenido. No se eliminó el registro.');window.close();</script>"
'    else	
'    	l_rs.close
		set l_cm = Server.CreateObject("ADODB.Command")
		l_sql = "DELETE FROM cap_estinformal WHERE estinfnro = " & l_estinfnro
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
'	end if	
	
	'l_cm.close
	cn.Close
	Set cn = Nothing
	Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location.reload();window.close();</script>"
 	
%>



