<% ' Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
Dim l_cm
Dim l_sql
Dim l_grabar
Dim l_iduser
Dim l_tenro 

Dim l_ubicacion

l_grabar = Request.QueryString("grabar")
l_iduser = request.QueryString("iduser")

' buscar tenro -------------------------------------------------
if len(l_grabar) <> 0 then
	l_ubicacion	= InStr(l_grabar,",")
	l_ubicacion = l_ubicacion + 1
	l_grabar = mid(l_grabar,l_ubicacion)
	
	l_ubicacion	= InStr(l_grabar,",")
	l_ubicacion	= l_ubicacion - 1
	l_tenro = left(l_grabar,l_ubicacion)
	l_grabar = right(l_grabar,(len(l_grabar)-l_ubicacion))
end if

set l_cm = Server.CreateObject("ADODB.Command")
' PRIMERO BORRAR -  gti_config_ctacte  --------------------
	l_sql = "DELETE FROM usupuedever "
	l_sql = l_sql & " WHERE iduser  = '"  & l_iduser
	l_sql = l_sql & "' AND   tenro    = "  & l_tenro
		
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0	
	
do until (len(trim(l_grabar)) = 0) or (trim(l_grabar) = ",")

' buscar estrnro  -------------------------------------------------	
	l_ubicacion	= InStr(l_grabar,",")
	l_ubicacion = l_ubicacion + 1
	l_grabar = mid(l_grabar,l_ubicacion)
	
	l_ubicacion	= InStr(l_grabar,",")
	l_ubicacion	= l_ubicacion - 1
	l_estrnro = left(l_grabar,l_ubicacion)
	l_grabar = right(l_grabar,(len(l_grabar)-l_ubicacion))

' grabar en gti_per_est lo que venia en la lista --------
	l_sql = "insert into usupuedever	 "
	l_sql = l_sql & "(iduser, tenro, estrnro) "
	l_sql = l_sql & " VALUES ('" & l_iduser & "'," 
	l_sql = l_sql & l_tenro & ","
	l_sql = l_sql & l_estrnro & ")"

	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0		
loop	
Set cn = Nothing
Set l_cm = Nothing

Response.write "<script>alert('Operacion Realizada');window.close();</script>"
%>
