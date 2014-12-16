<%  Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 
Dim l_cm
Dim l_sql
Dim l_grabar

Dim l_cystipnro
Dim l_userid
Dim l_ubicacion

l_grabar = Request.QueryString("grabar")
l_cystipnro = request.QueryString("cystipnro")


set l_cm = Server.CreateObject("ADODB.Command")
' PRIMERO BORRAR -  gti_turforpago  --------------------
	l_sql = "DELETE FROM cysfincirc "
	l_sql = l_sql & " WHERE cystipnro  = "  & l_cystipnro

	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0	
	

do until (len(trim(l_grabar)) = 0) or (trim(l_grabar) = ",")
' buscar turno -------------------------------------------------
	l_ubicacion	= InStr(l_grabar,",")
	l_ubicacion = l_ubicacion + 1
	l_grabar = mid(l_grabar,l_ubicacion)
	
	l_ubicacion	= InStr(l_grabar,",")
	l_ubicacion	= l_ubicacion - 1
	l_userid = left(l_grabar,l_ubicacion)
	l_grabar = right(l_grabar,(len(l_grabar)-l_ubicacion))

	
' grabar en gti_config_ctacte lo que venia en la lista --------
	l_sql = "insert into cysfincirc "
	l_sql = l_sql & "(userid, cystipnro) "
	l_sql = l_sql & " VALUES ('" & l_userid & "'," 
	l_sql = l_sql & l_cystipnro  & ")"

	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
loop	
Set cn = Nothing
Set l_cm = Nothing

Response.write "<script>alert('Operacion realizada');window.close();</script>"
%>
