<%option explicit%>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<%
'Archivo		: emp_est_formales_adp_04.asp
'Descripción	: relacion empleado estudios formales
'Autor			: lisandro moro
'Fecha			: 09/08/2003 
'Modificado:
'	Alvaro Bayon - 15-09-2003 - Se recibe como parámetro el nivel						
'								No existía el objeto comando
'	CCRossi - 10-10-2003 - Validar nulls antes de borrar
'
' de base de datos
dim l_rs
dim l_sql
dim l_cm
  
Dim l_ternro
Dim l_titnro
Dim l_instnro
Dim l_carredunro
Dim l_nivnro

l_ternro = l_ess_ternro
l_titnro = Request.QueryString("titnro")
l_instnro = Request.QueryString("instnro")
l_carredunro = Request.QueryString("carredunro")
l_nivnro = Request.QueryString("nivnro")

if l_titnro = "" or l_titnro = "0" then
	l_titnro = "null"
end if
if l_instnro = "" or l_instnro = "0" then
	l_instnro = "null"
end if
if l_carredunro = "" or l_carredunro = "0" then
	l_carredunro = "null"
end if

set l_cm = Server.CreateObject("ADODB.Command")
l_sql = "DELETE FROM cap_estformal WHERE ternro = " & l_ternro 
if l_titnro = "null"  then
l_sql = l_sql & " AND titnro IS NULL "
else
l_sql = l_sql & " AND titnro = " & l_titnro
end if
if l_instnro = "null"  then
l_sql = l_sql & " AND instnro IS NULL "
else
l_sql = l_sql & " AND instnro = " &l_instnro
end if
if l_nivnro = "null"  then
l_sql = l_sql & " AND nivnro IS NULL "  
else
l_sql = l_sql & " AND nivnro = " &l_nivnro
end if
if l_carredunro = "null"  then
l_sql = l_sql & " AND carredunro IS NULL "
else
l_sql = l_sql & " AND carredunro = " & l_carredunro
end if

'response.write l_sql
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0	

Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location.reload();window.close();</script>"
%>