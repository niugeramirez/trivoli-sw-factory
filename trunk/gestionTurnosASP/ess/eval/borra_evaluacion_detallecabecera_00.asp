
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<script>
returnValue = "";
</script>

<%
' modificado: 06-10-2003 - agregar transaccion y rollback
'-------------------------------------------------------------------------------------
on error goto 0

dim l_sql
dim l_cm

dim l_evaevenro
dim l_empleado

l_evaevenro = Request.QueryString("evaevenro")
l_empleado  = Request.QueryString("empleado")

'BORRAR Resultado de Notas.
cn.BeginTrans

 l_sql = "DELETE  FROM evadetevldor WHERE evadetevldor.evacabnro IN "
 l_sql = l_sql & " (SELECT evacabnro FROM evacab WHERE "
 l_sql = l_sql & " evacab.evaevenro  = " & l_evaevenro
 l_sql = l_sql & " AND   evacab.empleado = " & l_empleado & ")"
 set l_cm = Server.CreateObject("ADODB.Command")
 response.write(l_sql)
 l_cm.activeconnection = Cn
 l_cm.CommandText = l_sql
 cmExecute l_cm, l_sql, 0

 l_sql = "DELETE FROM evadet WHERE evadet.evacabnro IN "  
 l_sql = l_sql & " (SELECT evacabnro FROM evacab WHERE "
 l_sql = l_sql & " evacab.evaevenro  = " & l_evaevenro
 l_sql = l_sql & " AND   evacab.empleado = " & l_empleado & ")"
 set l_cm = Server.CreateObject("ADODB.Command")
 l_cm.activeconnection = Cn
 l_cm.CommandText = l_sql
 cmExecute l_cm, l_sql, 0
	
'BORRAR cabecera
 l_sql = "DELETE "
 l_sql = l_sql & " FROM evacab WHERE " 
 l_sql = l_sql & " evaevenro= " &  l_evaevenro
 l_sql = l_sql & " AND empleado = " &  l_empleado
 set l_cm = Server.CreateObject("ADODB.Command")
 l_cm.activeconnection = Cn
 l_cm.CommandText = l_sql
 cmExecute l_cm, l_sql, 0

if cn.Errors.Count<>0 then
	cn.RollbackTrans
	response.write("<script>alert('Ha ocurrido un error. No se realiza la baja.')</script>")
else
	cn.CommitTrans
	response.write "<script>window.close();</script>"
end if	
%>