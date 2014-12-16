<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : vales_liq_06.asp
Descripcion    : Validar que la fecha de pedido pertenezca al periodo
Creador        : Scarpa D.
Fecha Creacion : 27/01/2004
  Modificado  : 12/09/2006 Raul Chinestra - se agregó Vales en Autogestión   
-----------------------------------------------------------------------------
-->
<% 
on error goto 0

Dim l_tipo
Dim l_rs
Dim l_sql
Dim l_fecha1
Dim l_fecha2
Dim l_pliqnro
Dim l_error

l_fecha1 = request.QueryString("fecha1")
l_fecha2 = request.QueryString("fecha2")
l_pliqnro = request.QueryString("pliqnro")

Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = "SELECT * "
l_sql = l_sql & " FROM periodo "
l_sql = l_sql & " WHERE pliqnro = " & l_pliqnro
l_sql = l_sql & " AND pliqdesde <= " & cambiafecha(l_fecha1,"YMD",true)
l_sql = l_sql & " AND pliqhasta >= " & cambiafecha(l_fecha1,"YMD",true)

rsOpen l_rs, cn, l_sql, 0

l_error = l_rs.eof

l_rs.close

l_sql = "SELECT * "
l_sql = l_sql & " FROM periodo "
l_sql = l_sql & " WHERE pliqnro = " & l_pliqnro
l_sql = l_sql & " AND pliqdesde <= " & cambiafecha(l_fecha2,"YMD",true)
l_sql = l_sql & " AND pliqhasta >= " & cambiafecha(l_fecha2,"YMD",true)

rsOpen l_rs, cn, l_sql, 0

l_error = l_error OR l_rs.eof

l_rs.close


Set l_rs = Nothing
%>

<script>
<%if l_error then %>
		parent.datosIncorrectos();
<%else%>
		parent.datosCorrectos();
<%end if%>
</script>
