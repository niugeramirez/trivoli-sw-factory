<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<% Response.AddHeader "Content-Disposition", "attachment;filename=Estudio Formales.xls" %>

<%
on error goto 0 
'----------------------------------------------------------------------------------------------------------------
' Modificado: 08-10-2003 - CCRossi - Cambiar el select para que muestre los registros
'						             que no tiene titulol ni institucion
' Modificado: 25-02-2004 - Scarpa D. - Se agrego la opcion de estudio actual
' Modificado: 05-03-2004 - Scarpa D. - Cambio de actual por futuro
'			: 17-10-2005 - Leticia A. - Agregar FechaISO y sacar el campo futuro que no se muestra en el modulo
'			: 16-08-2007 - Gustavo Ring - Se agrego el campo de validación de RR.HH
'----------------------------------------------------------------------------------------------------------------
Dim rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_ternro
Dim l_estado 
Dim l_capestrrhh

l_estado 	= Request.QueryString("estado")
l_filtro 	= Request.QueryString("filtro")
l_orden  	= Request.QueryString("orden")
l_ternro 	= l_ess_ternro

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Estudios Formales - RHPro &reg;</title>
</head>


<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th>Nivel</th>
        <th>Titulo</th>
		<th>Institución</th>
		<th>Carrera</th>
		<th>Fecha Desde</th>
		<th>Fecha Hasta</th>
 		<th>RRHH</th>		
		<!-- <th>Futuro</th> -->
    </tr>

<%
 Set rs = Server.CreateObject("ADODB.RecordSet")
 l_sql =  "SELECT cap_estformal.titnro, titulo.titdesabr, cap_estformal.nivnro, nivest.nivdesc, cap_estformal.instnro"
 l_sql = l_sql & " , cap_carr_edu.carredudesabr,institucion.instdes,cap_estformal.carredunro, capfecdes, capfechas, capactual "
 l_sql = l_sql & " , capestrrhh "
 l_sql = l_sql & " FROM cap_estformal "
 l_sql = l_sql & " LEFT JOIN titulo       ON cap_estformal.titnro = titulo.titnro "
 l_sql = l_sql & " LEFT JOIN nivest       ON cap_estformal.nivnro = nivest.nivnro "
 l_sql = l_sql & " LEFT JOIN institucion  ON cap_estformal.instnro = institucion.instnro "
 l_sql = l_sql & " LEFT JOIN cap_carr_edu ON cap_estformal.carredunro = cap_carr_edu.carredunro "
 l_sql = l_sql & " WHERE cap_estformal.ternro = " & l_ternro

if l_filtro <> "" then
  l_sql = l_sql & " AND " & l_filtro & " "
end if

if l_estado <> "" then
  l_sql = l_sql & " AND capactual = " & l_estado & " "
end if

l_sql = l_sql & " "& l_orden

rsOpen rs, cn, l_sql, 0 
if not rs.eof then
	do until rs.eof
	%>
	    <tr>
	        <td nowrap align="left"><%= rs("nivdesc")%></td>
	        <td nowrap align="left"><%= rs("titdesabr")%></td>		
			<td nowrap align="left"><%= rs("instdes")%></td>
			<td nowrap align="left"><%= rs("carredudesabr")%></td>
			<td nowrap align="left"><%= fechaISO(rs("capfecdes"))%></td>
			<td nowrap align="left"><%= fechaISO(rs("capfechas"))%></td>
		    <%if not isnull(rs("capestrrhh")) then%>
				<%if rs("capestrrhh") = -1 then%> 
  				   <td nowrap align="center">Aceptado</td>			
				<%else%>
  				   <td nowrap align="center">Pendiente</td>						
				<%end if
			Else%>	
				   <td nowrap align="center">Pendiente</td>						
			<%end if 
			%>
	   			
			<%'if CInt(rs("capactual")) = -1 then%>
  			   <!-- <td nowrap align="center">No</td> -->
			<%'else%>
  			   <!-- <td nowrap align="center">Si</td> -->
			<%'end if%>
	    </tr><%
		rs.MoveNext
	loop
else
%> <tr>
        <td nowrap  colspan="8" align="center"><b>No hay registros</b></td>
   </tr><%
end if
rs.Close
cn.Close
%>
</table>

</body>
</html>