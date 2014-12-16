<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<% 
'---------------------------------------------------------------------------------------
'Archivo	: pedido_emp_vac_gti_06.asp
'Descripción: control de la superposicion de fecha
'Autor		: Scarpa D.
'Fecha		: 08/10/2004
'Modificado	: 
'---------------------------------------------------------------------------------------
on error goto 0

Dim l_tipo
Dim l_sql
Dim l_rs

dim l_vacnro
dim l_ternro
dim l_desde
dim l_hasta
Dim l_vdiapednro

l_tipo			= Request("tipo")

l_desde  		= request("desde")
l_hasta      	= request("hasta")
l_vdiapednro	= request("vdiapednro")

' INFORMIX cambiar
l_desde	= cambiaFecha(l_desde, "YMD", false)
l_hasta	= cambiaFecha(l_hasta, "YMD", false)

Set l_rs = Server.CreateObject("ADODB.RecordSet")

dim leg
leg = Session("empleg")
if leg = "" then
    response.write "NO SE HA SELECCIONADO UN EMPLEADO<BR>"
	Response.End
end if

l_sql = "SELECT ternro FROM empleado WHERE empleado.empleg = " & leg
l_rs.Open l_sql, cn
if l_rs.eof then
    response.write "NO SE HA SELECCIONADO UN EMPLEADO<BR>"
	response.end
else 
  l_ternro = l_rs("ternro")
end if
l_rs.close

l_sql =         " SELECT * FROM vacdiasped "
l_sql = l_sql & " WHERE  NOT ( (" & l_hasta & " < vdiapeddesde ) OR (" & l_desde & " > vdiapedhasta )) "
l_sql = l_sql & " AND ternro = " & l_ternro 
l_sql = l_sql & " AND vdiaspedestado = -1 "

if l_tipo = "M" then
   l_sql = l_sql & " AND   vacdiasped.vdiapednro <> " & l_vdiapednro
end if

rsOpen l_rs, cn, l_sql, 0 

if l_rs.eof then
%>
 <script>
  parent.rangoCorrecto();
 </script>   
<%
else
%>
 <script>
  parent.rangoIncorrecto();
 </script>   
<%end if

l_rs.close

%>

