<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--#include virtual="/turnos/shared/inc/fecha.inc"-->
<!--#include virtual="/turnos/shared/inc/util.inc"-->
<!--#include virtual="/turnos/shared/inc/adovbs.inc"-->

<!--
-----------------------------------------------------------------------------
Archivo        : criterios_04.asp
Descripcion    : Modulo que se encarga del ABM de criterios - baja
Creador        : Scarpa D.
Fecha Creacion : 28/11/2003
Modificacion   :
  06/02/2004 - Scarpa D. - Correccion al borrar los empleados
-----------------------------------------------------------------------------
-->
<% 
Dim l_tipo
Dim l_cm
Dim l_rs
Dim l_sql

Dim l_selnro

Set l_rs = Server.CreateObject("ADODB.RecordSet")
set l_cm = Server.CreateObject("ADODB.Command")

l_selnro = request("selnro")

cn.beginTrans

l_sql = " SELECT * FROM sel_ter WHERE selnro=" & l_selnro

rsOpenCursor l_rs, cn, l_sql, 0, adOpenKeyset

do until l_rs.eof
   l_sql = " DELETE FROM sel_ter WHERE ternro=" & l_rs("ternro") & " AND selnro=" & l_selnro
   
   l_cm.activeconnection = Cn
   l_cm.CommandText = l_sql
   cmExecute l_cm, l_sql, 0
   
   l_rs.moveNext
loop 

l_rs.close

l_sql = " DELETE FROM seleccion WHERE selnro=" & l_selnro 

l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0

cn.commitTrans

Set cn = Nothing
Set l_cm = Nothing
%>

<script>
  alert('Operación Realizada.');
  window.opener.ifrmfiltros.location.reload();
  window.close();
</script>
