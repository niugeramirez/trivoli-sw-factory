<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<script>
 returnValue = "";
</script>

<%
' Modificado: 28-09-2004 CCRossi- agregar la baja de evaarea
'-------------------------------------------------------------------------------------
dim l_sql
dim l_cm
dim l_rs

dim l_evaevenro
dim l_empleado

l_evaevenro = Request.QueryString("evaevenro")
l_empleado  = Request.QueryString("empleado")

dim l_lista

 Set l_rs = Server.CreateObject("ADODB.RecordSet") 
 l_sql =  " (select evadetevldor.evldrnro from evadetevldor "
 l_sql = l_sql & " INNER JOIN evacab ON evacab.evacabnro = evadetevldor.evacabnro "
 l_sql = l_sql & " WHERE evacab.evaevenro = " & l_evaevenro
 l_sql = l_sql & " AND   evacab.empleado = " & l_empleado & ")"
 rsOpen l_rs, cn, l_sql, 0 
 l_lista = "0"
 do until l_rs.eof
	l_lista= l_lista & "," &  l_rs("evldrnro")
	l_rs.movenext
 loop
 l_rs.close

 l_sql = "DELETE FROM evaresultado WHERE evaresultado.evldrnro IN "
 l_sql = l_sql & " (" & l_lista & ")"
' response.write(l_sql)

 set l_cm = Server.CreateObject("ADODB.Command") 
 l_cm.activeconnection = Cn
 l_cm.CommandText = l_sql
 cmExecute l_cm, l_sql, 0

 l_sql = "DELETE FROM evaarea WHERE evaarea.evldrnro IN "
 l_sql = l_sql & " (" & l_lista & ")"
' response.write(l_sql)

 set l_cm = Server.CreateObject("ADODB.Command") 
 l_cm.activeconnection = Cn
 l_cm.CommandText = l_sql
 cmExecute l_cm, l_sql, 0
 
 response.write "<script>window.close(); </script>"

%>