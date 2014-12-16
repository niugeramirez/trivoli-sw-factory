<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: lugares_con_03.asp
'Descripción: Abm de Lugares
'Autor : Raul Chinestra	
'Fecha: 04/08/2005
'Modificado : Raul Chinestra 02/03/2006 Se eliminaron los campos lugpro, lugbaj y se agregó el campo lugzon que indica la 
' zona comercial a la que pertenece el lugar y que se va a usar para bajar los cupos, contratos y ordenes de trabajo.
'Modificado : Gustavo 13/09/2006 Se agregó la direccion local del lugar.

Dim l_cm
Dim l_sql

Dim l_lugnro
Dim l_lugdir
Dim l_lugestacion
Dim l_lugdesvio


l_lugnro 	= request.Form("lugnro")
l_lugdir	= request.Form("lugdir")
l_lugestacion = request.Form("lugestacion")
l_lugdesvio	= request.Form("lugdesvio")

	set l_cm = Server.CreateObject("ADODB.Command")

	l_sql = "UPDATE tkt_lugar"
	l_sql = l_sql & " SET lugdir = '"	& l_lugdir & "'"
	l_sql = l_sql & " , estacion = '"	& l_lugestacion & "'"
	l_sql = l_sql & " , desvio = '"	& l_lugdesvio & "'"
	l_sql = l_sql & " WHERE lugnro = " & l_lugnro

	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	Set l_cm = Nothing

	Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
%>

