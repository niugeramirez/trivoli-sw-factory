<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<script>
returnValue = "";
</script>

<%
dim l_sql
dim l_cm

dim l_evaevenro
dim l_empleado

l_evaevenro = Request.QueryString("evaevenro")
l_empleado  = Request.QueryString("empleado")

'BORRAR Resultado de Notas.

l_sql = "DELETE FROM evaplan WHERE evaplan.evldrnro IN "
l_sql = l_sql & " (SELECT evadetevldor.evldrnro FROM evadetevldor "
l_sql = l_sql & " INNER JOIN evacab ON evacab.evacabnro = evadetevldor.evacabnro "
l_sql = l_sql & " WHERE evacab.evaevenro = " & l_evaevenro
l_sql = l_sql & " AND   evacab.empleado = " & l_empleado & ")"
set l_cm = Server.CreateObject("ADODB.Command")
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
 
response.write "<script>window.close(); </script>"

%>