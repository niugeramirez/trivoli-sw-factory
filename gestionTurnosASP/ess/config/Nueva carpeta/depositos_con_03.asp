<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: depositos_con_03.asp
'Descripción: ABM de Depósitos
'Autor : Alvaro Bayon
'Fecha: 11/02/2005

Dim l_tipo
Dim l_cm
Dim l_sql

Dim l_depnro
Dim l_depdes
Dim l_depmul
Dim l_depcod
Dim l_deptip

l_tipo 		= request.querystring("tipo")
l_depnro 	= request.Form("depnro")
l_depdes	= request.Form("depdes")
l_depcod 	= request.Form("depcod")
l_depmul	= request.Form("depmul")
l_deptip	= request.Form("deptip")

if len(l_depmul)>0 then
	l_depmul = -1
else
	l_depmul = 0
end if

	set l_cm = Server.CreateObject("ADODB.Command")
	if l_tipo = "A" then 
		l_sql = "INSERT INTO tkt_deposito"
		l_sql = l_sql & " (depdes, depcod, deptip, depmul)"
		l_sql = l_sql & " VALUES ('" & l_depdes 
		l_sql = l_sql & "','" & l_depcod & "','" & l_deptip & "'," & l_depmul
		l_sql = l_sql & ")"
	else
		l_sql = "UPDATE tkt_deposito"
		l_sql = l_sql & " SET depdes = '" & l_depdes & "'"
		l_sql = l_sql & ", depcod = '" & l_depcod & "'"
		l_sql = l_sql & ", deptip = '" & l_deptip & "'"
		l_sql = l_sql & ", depmul = " & l_depmul
		l_sql = l_sql & " WHERE depnro = " & l_depnro
	end if
	'response.write l_sql & "<br>"
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	Set l_cm = Nothing

	Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
%>

