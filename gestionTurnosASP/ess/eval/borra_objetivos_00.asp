<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<script>
 returnValue = "";
</script>

<%
on error goto 0
 dim l_sql
 dim l_cm
 dim l_rs1
 Dim lista
 
 dim l_evaevenro
 dim l_empleado

 l_evaevenro = Request.QueryString("evaevenro")
 l_empleado  = Request.QueryString("empleado")

    Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
    l_sql = "select evaluaobj.evaobjnro from evadetevldor "
    l_sql = l_sql & " INNER JOIN evaluaobj ON evaluaobj.evldrnro= evadetevldor.evldrnro"
    l_sql = l_sql & " INNER JOIN evacab ON evacab.evacabnro = evadetevldor.evacabnro "
    l_sql = l_sql & " WHERE evacab.evaevenro = " & l_evaevenro
    l_sql = l_sql & " AND   evacab.empleado = " & l_empleado 
    rsOpen l_rs1, cn, l_sql, 0 
    lista = "0"
    Do Until l_rs1.EOF
        lista = lista & "," & l_rs1("evaobjnro")
        l_rs1.MoveNext
    Loop
    l_rs1.Close
    set l_rs1=nothing

    'borrar todos los resultados de objetivos (tiene un trnro asociado)
    l_sql = "DELETE FROM evaluaobj  "
    l_sql = l_sql & " WHERE evaluaobj.evaobjnro IN "
    l_sql = l_sql & " (" & lista & ")"
    set l_cm = Server.CreateObject("ADODB.Command") 
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
 
    'borrar todos los planes smart si hay alguno
    l_sql = "DELETE FROM evaplan WHERE evaplan.evaobjnro IN "
    l_sql = l_sql & " (" & lista & ")"
    set l_cm = Server.CreateObject("ADODB.Command") 
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
    
    
    'Borrar todos los comentarios del objetivo que ya de ols EVADETEVLDOR
    l_sql = "DELETE FROM evaobjsgto  "
    l_sql = l_sql & " WHERE evaobjsgto.evaobjnro IN "
    l_sql = l_sql & " (" & lista & ")"
    set l_cm = Server.CreateObject("ADODB.Command") 
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
    
    'Borrar todos los puntajes de la evaluacion, que son de objetivos obviamente.
    l_sql = "DELETE FROM evapuntaje  "
    l_sql = l_sql & " WHERE EXISTS (SELECT * FROM evadetevldor  "
    l_sql = l_sql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    l_sql = l_sql & " WHERE evadetevldor.evacabnro=evapuntaje.evacabnro "
    l_sql = l_sql & " AND evacab.evaevenro = " & l_evaevenro
    l_sql = l_sql & " AND evacab.empleado = " & l_empleado & ")"
    set l_cm = Server.CreateObject("ADODB.Command") 
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
    
    
    'Borrar el puntaje de objetivos General (Deloitte)
    l_sql = "DELETE FROM evagralobj  "
    l_sql = l_sql & " WHERE EXISTS (SELECT * FROM evadetevldor  "
    l_sql = l_sql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    l_sql = l_sql & " where evadetevldor.evldrnro=evagralobj.evldrnro "
    l_sql = l_sql & " AND   evacab.evaevenro = " & l_evaevenro
    l_sql = l_sql & " AND   evacab.empleado  = " & l_empleado & ")"
    set l_cm = Server.CreateObject("ADODB.Command") 
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
 
    'Borrar objetivos sin
    l_sql = "DELETE FROM evaobjetivo  WHERE evaobjetivo.evaobjnro IN "
    l_sql = l_sql & " (" & lista & ")"
    set l_cm = Server.CreateObject("ADODB.Command") 
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
    
    
    cn.close
    set cn=nothing
 response.write "<script>window.close(); </script>" 
 
%>