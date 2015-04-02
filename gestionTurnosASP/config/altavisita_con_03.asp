<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--#include virtual="/turnos/shared/inc/fecha.inc"-->
<% 

Dim l_tipo
Dim l_cm
Dim l_sql
Dim l_rs

Dim l_id

Dim l_calfec
Dim l_descripcion

Dim l_idrecursoreservable

Dim l_pacienteid


l_tipo 		           = request.querystring("tipo")
l_idrecursoreservable  = request.Form("idrecursoreservable")
l_pacienteid     	   = request.Form("pacienteid")
l_calfec               = request("calfec")


Set l_rs = Server.CreateObject("ADODB.RecordSet")


set l_cm = Server.CreateObject("ADODB.Command")


	l_sql = "INSERT INTO visitas "
	l_sql = l_sql & "(fecha, idrecursoreservable, idpaciente , idturno ,created_by,creation_date,last_updated_by,last_update_date) "
	l_sql = l_sql & "VALUES (" & cambiaformato (l_calfec, "00:00" )  & "," & l_idrecursoreservable  & "," &  l_pacienteid & ",0"&",'"&session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE())"
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0

Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
%>

