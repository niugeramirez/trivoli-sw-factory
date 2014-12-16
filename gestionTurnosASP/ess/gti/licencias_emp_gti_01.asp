<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
-------------------------------------------------------------------------------------------------
Archivo       : licencias_emp_gti_01.asp
Descripcion   : Listado licencias
Creacion      : 24/03/2004
Autor         : Scarpa D.
Modificacion  :
  06/05/2004 - Scarpa D. - Se quitarin los campos de licencias parciales
  18/10/2004 - Scarpa D. - Mostrar la cantidad de dias corridos y habiles
  09-11-05 - Leticia A. - Si se configuro el ConfRep, mostrar los tipo de licencias configuradas.
  27-11-2006 - Diego Rosso - Cambio para que no se puedan modificar ni eliminar vacaciones que ya tengan pago/desc
-------------------------------------------------------------------------------------------------
-->
<% 
on error goto 0

Dim l_rs
Dim l_rs2
Dim l_sql
Dim l_elhoradesde
Dim l_elhorahasta

Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_estado
Dim l_ternro
 Dim l_repnro 
 Dim l_sql_confrep
 
 Dim l_ModEli 'Diego Rosso
 
 ' ************
 l_repnro = 151


Set l_rs  = Server.CreateObject("ADODB.RecordSet")


l_ternro = l_ess_ternro
l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY elfechadesde,elhoradesde"
end if


' ____________________________________________________________________
' Verificar si se cargaron Tipo de Licencias a mostrar en el ConfRep  
 l_sql = " SELECT repnro FROM confrep WHERE repnro=" & l_repnro
 rsOpen l_rs, cn, l_sql, 0 
 
 l_sql_confrep = ""
 if not l_rs.eof then  	' AND confrep.conftipo = 'TD' ?va
 	 l_sql_confrep = " INNER JOIN confrep ON confrep.confval=tipdia.tdnro  AND confrep.repnro="& l_repnro
 end if 
 l_rs.Close
 
 ' __________________________________________________________________
 
%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%= c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Licencias - Gesti&oacute;n de Tiempos - RHPro &reg;</title>
</head>

<script>
var jsSelRow = null;

function Deseleccionar(fila)
{
 fila.className = "MouseOutRow";
}
function Seleccionar(fila,cabnro,ternro,modelin,estado)
{
 if (jsSelRow != null)
 {
  Deseleccionar(jsSelRow);
 };

 document.datos.cabnro.value = cabnro;
 document.datos.ternro.value = ternro;
 document.datos.estado.value = estado;
 document.datos.ModEli.value = modelin;
 
 fila.className = "SelectedRow";
 jsSelRow		= fila;
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th align="left">Licencia</th>
        <th align="left">Apellido y Nombre</th>
        <th align="center">Desde</th>
        <th align="center">Hasta</th>
        <th align="center">D&iacute;as</th>
        <th align="center">H&aacute;b.</th>		
        <th align="center">Estado</th>
    </tr>
<%
l_sql = " SELECT emp_licnro, tddesc, empleado.ternro, empleado.terape, empleado.ternom, elfechadesde, elfechahasta, licestdesabr, emp_lic.licestnro "
l_sql = l_sql & " , elcantdias "
l_sql = l_sql & " FROM emp_lic INNER JOIN empleado ON emp_lic.empleado=empleado.ternro "
l_sql = l_sql & " INNER JOIN tipdia ON emp_lic.tdnro=tipdia.tdnro "
if l_sql_confrep <> "" then
	l_sql = l_sql & l_sql_confrep
end if
l_sql = l_sql & " LEFT JOIN lic_estado ON emp_lic.licestnro = lic_estado.licestnro"
l_sql = l_sql & " WHERE emp_lic.empleado= " & l_ternro 

if l_filtro <> "" then
  l_sql = l_sql & " AND " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden

'response.write l_sql
rsOpen l_rs, cn, l_sql, 0 

if not l_rs.eof then
 l_modeli = 0
 
	do until l_rs.eof
		l_estado = l_rs("licestnro")
		if isNull(l_estado) then
			l_estado = 0
		end if
		
		'Diego Rosso
 			 Set l_rs2  = Server.CreateObject("ADODB.RecordSet")
			 l_sql = " SELECT vacpdnro FROM vacpagdesc WHERE emp_licnro =" & l_rs("emp_licnro")
			rsOpen l_rs2, cn, l_sql, 0 

			if not l_rs2.eof then
	   			l_modeli = l_rs2("vacpdnro")
			end if
			'----------
%>
    	<tr onclick="Javascript:Seleccionar(this,<%= l_rs("emp_licnro")%>,<%= l_rs("ternro")%>,<%= l_modeli %>,<%= l_estado%>)">
        	<td nowrap align="left"><%= l_rs("tddesc")%></td>
	        <td nowrap align="left"><%= l_rs("terape") & " "& l_rs("ternom") %></td>
    	    <td nowrap align="center"><%= l_rs("elfechadesde")%></td>
        	<td nowrap align="center"><%= l_rs("elfechahasta")%></td>
	        <td nowrap align="center"><%= DateDiff("d",CDate(l_rs("elfechadesde")),CDate(l_rs("elfechahasta"))) + 1 %></td>				
    	    <td nowrap align="center"><%= l_rs("elcantdias")%></td>
        	<td nowrap align="center"><%= l_rs("licestdesabr")%></td>
		</tr>
<%
		l_rs.MoveNext
	loop
else
%>
	<td colspan="7">No se encontraron datos. </td>
<%
end if

l_rs.Close
set l_rs = Nothing


cn.Close
set cn = Nothing
%>
</table>

<form name="datos" method="post">
<input type="Hidden" name="cabnro" value="0">
<input type="Hidden" name="ternro" value="0">
<input type="Hidden" name="estado" value="0">
<input type="Hidden" name="orden" value="<%= l_orden %>">
<input type="Hidden" name="filtro" value="<%= l_filtro %>">
<input type="Hidden" name="ModEli" value="0"><!-Diego Rosso->
</form> 

</body>
</html>
