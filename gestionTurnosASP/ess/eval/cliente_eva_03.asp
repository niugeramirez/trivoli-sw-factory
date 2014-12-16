<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<script>
//var xc = screen.availWidth;
//var yc = screen.availHeight;
//window.moveTo(xc,yc);	
//window.resizeTo(150,150);
</script>
<%
'================================================================================
'Archivo		: engagement_eva_03.asp
'Descripción	: Grabar Engagement
'Autor			: CCRossi
'Fecha			: 26-08-2004
'Modificado		: 13-12-2004 CCRossi- Cambiar tercero por tabla evacliente
'================================================================================

Dim l_tipo
Dim l_rs
Dim l_cm
Dim l_sql

Dim l_evaclinro
Dim l_evaclinom
Dim l_evaclicodext

l_tipo 			= request.querystring("tipo")

l_evaclinro		= request.Form("evaclinro")
l_evaclinom 	= request.Form("evaclinom")
l_evaclicodext	= request.Form("evaclicodext")

set l_cm = Server.CreateObject("ADODB.Command")
if l_tipo = "A" then 
	l_sql = "INSERT INTO evacliente "
	l_sql = l_sql & "(evaclinom,evaclicodext)"
	l_sql = l_sql & "VALUES ('" & l_evaclinom &"','"&l_evaclicodext &  "')"
else
	l_sql = "UPDATE evacliente SET"
	l_sql = l_sql & " evaclinom    = '" & l_evaclinom    & "', "
	l_sql = l_sql & " evaclicodext = '" & l_evaclicodext & "'"
	l_sql = l_sql & " WHERE evaclinro = " & l_evaclinro
end if

'response.write l_sql & "<br>"
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location.reload();window.close();</script>"
%>
