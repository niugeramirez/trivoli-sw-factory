<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<%
'================================================================================
'Archivo		: proyecto_eva_ag_04.asp
'Descripción	: Baja de Proyectos
'Autor			: CCRossi
'Fecha			: 30-08-2004
'Modificado		: 08-08-2005 L.A - borrar evento, si no hay ningun empleado con evaluac.
'================================================================================
'on error goto 0

Dim l_cm
Dim l_rs, l_rs1
Dim l_sql
Dim l_evaproynro
dim l_borrar
	
l_borrar = "SI"
	
	l_evaproynro = request.querystring("cabnro")
	l_evaproynro = trim(l_evaproynro)
	
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
	
	l_sql = "SELECT evaproynro  FROM evaevento  WHERE evaproynro =" & l_evaproynro
	rsOpen l_rs, cn, l_sql, 0
	
	if not l_rs.eof then 'me fijo si tiene alguna evaluacion asociada 
		l_sql = "SELECT evaevenro FROM  evacab WHERE evaproynro=" & l_evaproynro
		rsOpen l_rs1, cn, l_sql, 0 
		if not l_rs1.eof then
			l_borrar="NO"
		else 
		   	set l_cm = Server.CreateObject("ADODB.Command")
			l_sql = "DELETE FROM evaevento WHERE evaproynro = " & l_evaproynro
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0
		end if
		l_rs1.close
		set l_rs1 = nothing
	end if
	l_rs.close
	set l_rs = nothing
	
	
	if l_borrar="NO" then
		Response.write "<script>alert('Este Proyecto está asociado a un Evento de Evaluación. No se eliminó el registro.');window.close();</script>"
    else	
    	set l_cm = Server.CreateObject("ADODB.Command")
		l_sql = "DELETE FROM evaproyemp WHERE evaproynro = " & l_evaproynro
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
		
		  ' set l_cm = Server.CreateObject("ADODB.Command")
		l_sql = "DELETE FROM evaproyecto WHERE evaproynro = " & l_evaproynro
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
		
		Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location.reload();window.close();</script>"
	end if	
	
	cn.Close
	Set cn = Nothing
%>
