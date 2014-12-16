<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--
Archivo    : pedido_emp_vac_gti_04.asp
Descripción: Pedido Vacaciones
Autor      : Scarpa D.
Fecha      : 08/10/2004
Modificado : 
-->
<% 
on error goto 0

Dim l_cm
Dim l_sql
Dim l_rs
Dim l_vdiapednro
Dim l_ternro
Dim l_vdiapeddesde
Dim l_vdiapedhasta
Dim l_vacnro

l_vdiapednro = request("vdiapednro")

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

'Busco las fechas del pedido

l_sql = "SELECT * "
l_sql = l_sql & " FROM  vacdiasped "
l_sql = l_sql & " WHERE vacdiasped.vdiapednro = " & l_vdiapednro

rsOpen l_rs, cn, l_sql, 0 

l_vacnro			= l_rs("vacnro")
l_vdiapeddesde		= l_rs("vdiapeddesde")
l_vdiapedhasta		= l_rs("vdiapedhasta")

l_rs.close

'Controlo si la licencia se puede modificar

l_sql = "SELECT elfechadesde,elfechahasta, elcantdias, vacnotifestado "
l_sql = l_sql & "FROM emp_lic INNER JOIN lic_vacacion ON lic_vacacion.emp_licnro = emp_lic.emp_licnro "
l_sql = l_sql & "LEFT JOIN vacnotif ON vacnotif.emp_licnro = emp_lic.emp_licnro "
l_sql = l_sql & "WHERE licestnro=2 AND empleado = " & l_ternro & " and emp_lic.tdnro = 2 AND lic_vacacion.vacnro = " & l_vacnro
l_sql = l_sql & " AND elfechadesde >= " & cambiafecha(l_vdiapeddesde,"YMD",true)
l_sql = l_sql & " AND elfechahasta <= " & cambiafecha(l_vdiapedhasta,"YMD",true)

rsOpen l_rs, cn, l_sql, 0 

if not l_rs.eof then
   l_rs.close
   cn.Close
%>
  <script>
     alert('El pedido tiene licencias asociadas y no se puede eliminar.');
	 window.close();
  </script>
<%
else
    l_rs.close

	'Borro el pedido
	set l_cm = Server.CreateObject("ADODB.Command")
	l_sql = "DELETE FROM vacdiasped WHERE vdiapednro = " & l_vdiapednro
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0

	cn.Close

     Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location.reload();window.close();</script>"

end if

%>
