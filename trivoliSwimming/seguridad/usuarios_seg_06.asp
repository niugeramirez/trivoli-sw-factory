<%  Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/inc/sec.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/const.inc"-->
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/fecha.inc"-->
<% 
Dim l_cm
Dim l_sql
Dim l_rs
Dim l_grabar
dim l_lista

dim l_userid

Dim l_ubicacion
Dim l_cystipnro

l_grabar = Request.QueryString("grabar")
l_userid = Request.QueryString("userid")

if len(l_grabar) <> 0 then
	' PRIMERO BORRAR -  cysfincirc  --------------------
	set l_cm = Server.CreateObject("ADODB.Command")
	l_sql = "DELETE FROM cysfincirc "
	l_sql = l_sql & " WHERE cysfincirc.userid = '" & l_userid & "'"
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	l_cm.Execute
end if

do until (len(trim(l_grabar)) = 0) or (trim(l_grabar) = ",")
		' buscar cysfincirc.cystipnro -------------------------------------------------
		l_ubicacion	= InStr(l_grabar,",")
		l_ubicacion = l_ubicacion + 1
		l_grabar = mid(l_grabar,l_ubicacion)
	
		l_ubicacion	= InStr(l_grabar,",")
		l_ubicacion	= l_ubicacion - 1
		l_cystipnro = left(l_grabar,l_ubicacion)
		l_grabar = right(l_grabar,(len(l_grabar)-l_ubicacion))
	
		' grabar en cysfininc  lo que venia en la lista --------
		l_sql = "INSERT INTO cysfincirc "
		l_sql = l_sql & "(cystipnro, userid ) "
		l_sql = l_sql & " VALUES (" & l_cystipnro & ",'" 
		l_sql = l_sql &  l_userid & "') "
	
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		l_cm.Execute
	
loop	
Set cn = Nothing
Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.close();</script>"
%>
