<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<script>
 returnValue = "";
</script>

<%
 dim l_sql
 dim l_cm
 dim l_rs

 dim l_evaevenro
 dim l_empleado

 l_evaevenro = Request.QueryString("evaevenro")
 l_empleado  = Request.QueryString("empleado")

 l_sql = "DELETE FROM evaluaestr  "
 l_sql = l_sql & " WHERE EXISTS (SELECT * FROM evadetevldor  "
 l_sql = l_sql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
 l_sql = l_sql & " where evadetevldor.evldrnro=evaluaestr.evldrnro "
 l_sql = l_sql & " AND evacab.evaevenro = " & l_evaevenro
 l_sql = l_sql & " AND evacab.empleado = " & l_empleado & ")"
 set l_cm = Server.CreateObject("ADODB.Command") 
 l_cm.activeconnection = Cn
 l_cm.CommandText = l_sql
 cmExecute l_cm, l_sql, 0

 l_sql = "DELETE FROM evaestructuras  "
 l_sql = l_sql & " WHERE EXISTS (SELECT * FROM evadetevldor  "
 l_sql = l_sql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
 l_sql = l_sql & " WHERE evacab.evacabnro=evaestructuras.evacabnro "
 l_sql = l_sql & " AND evacab.evaevenro = " & l_evaevenro
 l_sql = l_sql & " AND evacab.empleado = " & l_empleado & ")"
 set l_cm = Server.CreateObject("ADODB.Command") 
 l_cm.activeconnection = Cn
 l_cm.CommandText = l_sql
 cmExecute l_cm, l_sql, 0
 
 response.write "<script>window.close(); </script>" 
 
%>