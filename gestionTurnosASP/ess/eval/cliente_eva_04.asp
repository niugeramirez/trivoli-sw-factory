<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<%
'================================================================================
'Archivo		: cliente_eva_00.asp
'Descripción	: Baja de Clientes
'Autor			: CCRossi
'Fecha			: 13-12-2004
'Modificado		:  
'================================================================================
Dim l_cm
Dim l_rs
Dim l_sql
Dim l_evaclinro
	
	l_evaclinro = request.querystring("cabnro")
	
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT evaclinro "
	l_sql = l_sql & " FROM evaengage "
	l_sql = l_sql & " WHERE evaclinro = " & l_evaclinro
	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof then
		l_rs.close
		set l_rs = nothing
		Response.write "<script>alert('Este Cliente está asociado a un Engagement. No se eliminó el registro.');window.close();</script>"
    else	
    	l_rs.close
	
		set l_cm = Server.CreateObject("ADODB.Command")
		l_sql = "DELETE FROM evacliente WHERE evaclinro = " & l_evaclinro
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
	end if	
	
	cn.Close
	Set cn = Nothing
	Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location.reload();window.close();</script>"
 	
%>
