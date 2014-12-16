<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : gen_select_emp_03.asp
Creador        : Scarpa D.
Fecha Creacion : 27/11/2003
Descripcion    : Lista de empleados derecha - empleados selectados
Modificacion   :
  23/12/2003 - Scarpa D. - Cambio en la forma de la pagina.
-----------------------------------------------------------------------------
-->
<% 
on error goto 0

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_lista
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_canttotal
Dim l_cantFiltro

Dim l_selalto
Dim l_selancho

'Modifica los tiempos de espera del IIS y de la BD
Server.ScriptTimeout = 1200
cn.close
cn.ConnectionTimeout = 1200 
cn.open
cn.CommandTimeout = 1200 

l_selalto    = request("selalto")
l_selancho   = request("selancho")

l_filtro = request("sqlfiltroder")
l_orden  = request("sqlordender")
l_lista  = request("seleccion")

if l_orden = "" then
   l_orden = " ORDER BY empleg"  'orden por defecto legajo
end if

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<link href="/turnos/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<html>
<head>
	<title><%= Session("Titulo")%>Untitled</title>
<script languaje="javascript">
function Cargar1(){
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = "SELECT DISTINCT ternro,empleg, terape, ternom "
l_sql = l_sql & "FROM v_empleado "
l_sql = l_sql & "WHERE ternro IN (" & l_lista & ") "

if l_filtro <> "" then
  l_sql = l_sql & " AND " & l_filtro & " "
end if

l_sql = l_sql & " " & l_orden

rsOpen l_rs, cn, l_sql, 0 
l_cantfiltro = -1
do until l_rs.eof
    response.write "newOp = new Option();" & vbCrLf
    response.write "newOp.value  = '" & l_rs("ternro") & "';" & vbCrLf
    response.write "newOp.text   = '" & l_rs("empleg") & " - " & l_rs("terape") & ", " & l_rs("ternom") & "';"  & vbCrLf
    l_cantfiltro = l_cantfiltro + 1
    response.write "document.registro.selfil.options[" & l_cantfiltro & "] = newOp;" & vbCrLf
	l_rs.MoveNext
loop
l_rs.Close
set l_rs = Nothing
cn.Close
set cn = Nothing
%>  
}

</script>	
</head>

<body topmargin="0" leftmargin="0" rightmargin="0" scroll=no>
<form name="registro">
<input type="Hidden" name="lista" value="<%= l_lista %>">
<select  size=20 style="width:<%= l_selancho%>px;height:<%= l_selalto%>px" name="selfil" ondblclick="parent.Uno(selfil,parent.nselfil.registro.nselfil, parent.document.datos.totalder, parent.document.datos.totalizq);"></select>
</form>
<script>
Cargar1();
parent.document.datos.filtroder.value = document.registro.selfil.length;
<%if l_filtro = "" then%>
parent.document.datos.totalder.value = document.registro.selfil.length;	
<%end if%>

</script>
</body>
</html>
