<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->

<!--
'--------------------------------------------------------------------------
'Archivo       : borra_datosadm_eva_00.asp
'Descripcion   : borra datos de la tabla evadatosadm
'Creacion      : 23-12-2004    
'Autor         : Leticia Amadio.
'--------------------------------------------------------------------------
-->
<script>
returnValue = "";

</script>

<%
on error goto 0
Dim l_sql
Dim l_cm

Dim l_evldrnro

 dim l_evaevenro
 dim l_empleado

 l_evaevenro = Request.QueryString("evaevenro")
 l_empleado  = Request.QueryString("empleado")

'l_evldrnro = Request.QueryString("evldrnro")

Set l_cm = Server.CreateObject("ADODB.Command")

'verificar alguna dependencia?? ???
 l_sql = "DELETE FROM evadatosadm  "
 l_sql = l_sql & " WHERE EXISTS (SELECT * FROM evadetevldor  "
 l_sql = l_sql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
 l_sql = l_sql & " where evadetevldor.evldrnro=evadatosadm.evldrnro "
 l_sql = l_sql & " AND evacab.evaevenro = " & l_evaevenro
 l_sql = l_sql & " AND evacab.empleado  = " & l_empleado & ")"
 set l_cm = Server.CreateObject("ADODB.Command") 
 l_cm.activeconnection = Cn
 l_cm.CommandText = l_sql
 cmExecute l_cm, l_sql, 0
 
'l_sql = "DELETE FROM evadatosadm WHERE evldrnro = " & l_evldrnro

l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
 
response.write "<script>window.close(); </script>"
%>