<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
Archivo		: emp_est_formales_adp_01.asp
Descripción	: relacion empleado estudios formales
Autor			: lisandro moro
Fecha			: 09/08/2003 
Modificado:
	Alvaro Bayon - 15-09-2003 - Se recibe como parámetro el nivel						
								No existía el objeto comando
	Alvaro Bayon - 16-09-2003 - Se cambia la selección de campos por los de la tabla relacionada.
								Se usa Left join para la consulta.
								Se amplía la cantidad de columnas que se ven.
    Scarpa D. - 25-02-2004 - Se agrego la opcion de estudio actual								
    Scarpa D. - 05-03-2004 - Se cambio actual por futuro
	Gustavo Ring - 16-08-2007 - Se agrego el campo de validación de RR.HH
	Martin Ferraro - 30/08/2007 - Correccion de tipos en ORACLE
-->
<% 
on error goto 0

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_ternro
'Dim l_sqlfiltro
'Dim l_sqlorden
'Dim l_empleado
Dim l_estado 
Dim l_empleg
Dim l_capestrrhh

l_ternro = l_ess_ternro
l_empleg = l_ess_empleg

l_estado 	= Request.QueryString("estado")
l_filtro 	= Request.QueryString("filtro")
l_orden  	= Request.QueryString("orden")
'l_ternro 	= Request.QueryString("ternro")

if l_orden = "" then
	l_orden = "ORDER BY cap_estformal.capfecdes, cap_estformal.capfechas"
end if
%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Estudios Formales - RHPro &reg;</title>
</head>

<script>
var jsSelRow = null;

function Deseleccionar(fila)
{
 fila.className = "MouseOutRow";
}
function Seleccionar(fila,titu,nive,inst,carr){
 if (jsSelRow != null) {
  Deseleccionar(jsSelRow);
 };
 document.datos.cabnro.value = fila;
 document.datos.titnro.value = titu;
 document.datos.nivnro.value = nive;
 document.datos.instnro.value = inst;
 document.datos.carredunro.value = carr;
 fila.className = "SelectedRow";
 jsSelRow		= fila;
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th>Nivel</th>
        <th>Titulo</th>
		<th>Institución</th>
		<th>Carrera</th>
		<th nowrap>Fecha&nbsp;Desde</th>
		<th nowrap>Fecha&nbsp;Hasta</th>
 		<th>RRHH</th>		
    </tr>

<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")

 l_sql =  "SELECT cap_estformal.titnro, titulo.titdesabr, cap_estformal.nivnro, nivest.nivdesc, cap_estformal.instnro"
 l_sql = l_sql & " , cap_carr_edu.carredudesabr,institucion.instdes,cap_estformal.carredunro, capfecdes, capfechas"', capactual "
 l_sql = l_sql & " , capestrrhh "
  l_sql = l_sql & " FROM cap_estformal "
 l_sql = l_sql & " LEFT JOIN titulo       ON cap_estformal.titnro = titulo.titnro "
 l_sql = l_sql & " LEFT JOIN nivest       ON cap_estformal.nivnro = nivest.nivnro "
 l_sql = l_sql & " LEFT JOIN institucion  ON cap_estformal.instnro = institucion.instnro "
 l_sql = l_sql & " LEFT JOIN cap_carr_edu ON cap_estformal.carredunro = cap_carr_edu.carredunro "
 l_sql = l_sql & " WHERE cap_estformal.ternro = " & l_ternro 
 'l_sql = l_sql & " ORDER BY cap_estformal.capfecdes, cap_estformal.capfechas"

if l_filtro <> "" then
  l_sql = l_sql & " AND " & l_filtro & " "
end if

if l_estado <> "" then
   l_sql = l_sql & " AND capactual = " & l_estado & " "
end if

l_sql = l_sql & " "& l_orden

rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	do until l_rs.eof
	%>
	    <tr onClick="Javascript:Seleccionar(this,'<%= l_rs("titnro")%>','<%= l_rs("nivnro")%>','<%= l_rs("instnro")%>','<%= l_rs("carredunro")%>')">
	        <td nowrap align="left"><%= l_rs("nivdesc")%></td>
	        <td nowrap align="left"><%= l_rs("titdesabr")%></td>		
			<td nowrap align="left"><%= l_rs("instdes")%></td>
			<td nowrap align="left"><%= l_rs("carredudesabr")%></td>
			<td nowrap align="left"><%= l_rs("capfecdes")%></td>
			<td nowrap align="left"><%= l_rs("capfechas")%></td>
	        <%if not isnull(l_rs("capestrrhh")) then%>
				<%if clng(l_rs("capestrrhh")) = -1 then%> 
  				   <td nowrap align="center">Aceptado</td>			
				<%else%>
  				   <td nowrap align="center">Pendiente</td>						
				<%end if
			Else%>	
				   <td nowrap align="center">Pendiente</td>						
			<%end if 
			%>
	    </tr><%
		l_rs.MoveNext
	loop
else
%> <tr>
        <td nowrap  colspan="8" align="left">No posee Estudios Formales</td>
   </tr><%
end if
l_rs.Close
set l_rs = Nothing
cn.Close
set cn = Nothing
%>
</table>
<form name="datos" method="post">
<input type="Hidden" name="ternro" value="">
<input type="Hidden" name="cabnro" value="">
<input type="Hidden" name="titnro" value="">
<input type="Hidden" name="nivnro" value="">
<input type="Hidden" name="instnro" value="">
<input type="Hidden" name="carredunro" value="">
<input type="Hidden" name="orden" value="<%= l_orden %>">
<input type="Hidden" name="filtro" value="<%= l_filtro %>">
</form>

</body>
</html>