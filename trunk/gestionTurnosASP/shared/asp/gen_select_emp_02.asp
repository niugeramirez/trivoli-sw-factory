<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : gen_select_emp_02.asp
Creador        : Scarpa D.
Fecha Creacion : 27/11/2003
Descripcion    : Lista de empleados izquierda - empleados no selectados
Modificacion   :
  09/01/2004 - Scarpa D. - Restringir el conjunto de empleados que puede aparecer en el lado izq.
  09/01/2004 - Scarpa D. - Mantiene los empleados de la izquierda al cambiar el criterio
-----------------------------------------------------------------------------
-->
<% 
on error goto 0

Dim l_rs
Dim l_sql

Dim l_filtro
Dim l_lista
Dim l_lista2
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_canttotal
Dim l_cantFiltro
Dim l_sqlfiltroOnly

dim l_posicion
dim l_auxiliar

Dim l_seleccion
Dim l_sqlfiltroemp
Dim l_sqloperando
Dim l_hay_filtro

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

l_seleccion    = request("seleccion")
l_sqlfiltroemp = request("sqlfiltroemp")
l_sqloperando  = request("sqloperando")

l_filtro = request("sqlfiltroizq")
l_orden  = request("sqlordenizq")

l_sqlfiltroOnly = request("sqlfiltroonly")

if l_orden = "" then
   l_orden = " ORDER BY empleg"  'orden por defecto legajo
end if

'if l_lista = "" then
'  l_lista2 = " ternro = 0 "
'  else
'  l_lista2 = " ternro NOT IN (" & l_lista & ") "
'end if

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<link href="/serviciolocal/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<html>
<head>
	<title><%= Session("Titulo")%>Untitled</title>
<script languaje="javascript">
function Cargar1(){
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")

'response.write "parent.totalizq.value = " & rs(0) & ";" & vbCrLf

' Armo la consulta SQL
l_sql =         " SELECT DISTINCT * "
l_sql = l_sql & " FROM v_empleado "
l_sql = l_sql & " WHERE "

Dim l_arr
Dim l_i

l_arr = split(l_sqlfiltroemp,";")

l_sql = l_sql & " ( "

l_hay_filtro = false

' Busco todos los filtros
for l_i = 0 to UBound(l_arr)
    if l_hay_filtro then
       l_sql = l_sql & l_sqloperando & " ternro IN ( "
       l_sql = l_sql & l_arr(l_i) 
       l_sql = l_sql & " ) "
	else
	   l_hay_filtro = true
       l_sql = l_sql & " ternro IN ( "
       l_sql = l_sql & l_arr(l_i) 
       l_sql = l_sql & " ) "
	end if
next

if not l_hay_filtro then
   l_sql = l_sql & " 1=1 "
end if

l_sql = l_sql & " ) "

' Elimino de la seleccion los elementos seleccionados
if l_seleccion <> "" then
   l_sql = l_sql & " AND ternro NOT IN ( "
   l_sql = l_sql & l_seleccion
   l_sql = l_sql & " ) "
end if

' Solo muestro los empleados que estan en filtro Only
if l_sqlfiltroOnly <> "" then
   l_sql = l_sql & " AND ternro IN ( "
   l_sql = l_sql & l_sqlfiltroOnly
   l_sql = l_sql & " ) "  
   l_hay_filtro = true
end if


if l_hay_filtro then

    if l_filtro <> "" then
       l_sql = l_sql & "AND " & l_filtro & " "
    end if

    l_sql = l_sql & l_orden

	rsOpen l_rs, cn, l_sql, 0 
	l_cantfiltro = -1
	
	do until l_rs.eof
	
	    response.write "newOp = new Option();" & vbCrLf
	    response.write "newOp.value  = '" & l_rs("ternro") & "';" & vbCrLf
	    response.write "newOp.text   = '" & l_rs("empleg") & " - " & l_rs("terape") & ", " & l_rs("ternom") & "';"  & vbCrLf
	    l_cantfiltro = l_cantfiltro + 1
	    response.write "document.registro.nselfil.options[" & l_cantfiltro & "] = newOp;" & vbCrLf
	
		l_rs.MoveNext
	loop
	l_rs.Close
	
end if
	
set l_rs = Nothing
cn.Close
set cn = Nothing

%>  
}

</script>	
</head>

<body topmargin="0" leftmargin="0" rightmargin="0" scroll=no>
<form name="registro">
<select size=20 style="width:<%= l_selancho%>px;height:<%= l_selalto%>px" name=nselfil ondblclick="parent.Uno(nselfil,parent.selfil.registro.selfil, parent.document.datos.totalizq, parent.document.datos.totalder);"></select>
</form>
<script>
	Cargar1();
	parent.document.datos.filtroizq.value = document.registro.nselfil.length;
	<%if l_filtro = "" then%>
	parent.document.datos.totalizq.value = document.registro.nselfil.length;	
	<%end if%>
</script>
</body>
</html>
