<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<script>
window.returnValue='0';
</script>
<%
'Archivo	: tomarrevisor_eva_00.asp
'Descripción: buscar el revisor para evacabnro
'Autor		: CCRossi
'Fecha		: 03-06-2004
'Modificacion: 
'-------------------------------------------------------------------------------
 Dim l_cm
 Dim l_sql
 Dim l_rs
 
'variables locales
 dim l_evacabnro 
 
'uso local
 dim l_revisor
   
 l_evacabnro = Request.QueryString("evacabnro")

' ------------------------------------------------------------------------------------
'											BODY 
' ------------------------------------------------------------------------------------

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evaluador, empleg FROM evadetevldor "
l_sql = l_sql & " INNER JOIN empleado ON  evadetevldor.evaluador=empleado.ternro "
l_sql = l_sql & " WHERE  evatevnro = "  & cevaluador
l_sql = l_sql & " AND   evacabnro  = " & l_evacabnro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.EOF then
	l_revisor = l_rs("empleg")
else
	l_revisor = ""
end if		
l_rs.Close
set l_rs=nothing
%>
<script>
window.returnValue='<%=l_revisor%>';
window.close();
</script>
