<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<% 
' variables
' parametros de entrada ----------------------------------------
  Dim l_evatnnro
  Dim l_evldrnro
  Dim l_evanotadesc
  Dim l_cantidad
  
' variables de base de datos ------------------------------------
  Dim l_cm
  Dim l_sql
  Dim l_rs
    
' parametros de entrada
  l_evatnnro	= Request.QueryString("evatnnro")
  l_evanotadesc = trim(request.querystring("evanotadesc"))
  l_cantidad = len(l_evanotadesc)
  l_evldrnro    = request.querystring("evldrnro")
  
Response.write("l_evanotadesc=")
Response.write(l_evanotadesc)
Response.write("<br>")
Response.write("l_evatnnro=")
Response.write(l_evatnnro)

Response.write("<br>")
Response.write("l_evldrnro=")
Response.write(l_evldrnro)

'BODY ----------------------------------------------------------


if l_cantidad <= 200 then
	l_sql = "UPDATE evanotas "
	l_sql = l_sql & " SET evanotadesc = '"		   & trim(l_evanotadesc) & "'"
	l_sql = l_sql & " WHERE evanotas.evatnnro = "  & l_evatnnro
	l_sql = l_sql & " AND   evanotas.evldrnro = "  & l_evldrnro
	set l_cm = Server.CreateObject("ADODB.Command")  
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	Response.write("<br>")
	Response.write(l_sql)
else
    response.write("<script>alert('Las Notas deben no deben superar los 200 caracteres')</script>")
end if
%>
