<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<%
'================================================================================
'Archivo		: proyecto_eva_ag_01.asp
'Descripción	: Abm de Proyectos
'Autor			: CCRossi
'Fecha			: 30-08-2004
'Modificado		: 15-12-2004 CCRossi
' Segun el perfil (empleado o no empleado, se mostrarán los proyectos...
'				: 28-07-2005 - LA. - cambio codigo de proyecto por cod de evento.
'				: 19-08-2005 - muestre proyectos si no tiene evaluac generada o si al menos una evaluacion (cabaprobada, evacab)no se termino
'================================================================================
on error goto 0

Dim l_rs
Dim l_rs1
Dim l_sql
Dim l_sqlfiltro
Dim l_sqlorden

Dim l_mostrar
Dim l_evaluacs 

'parametros
Dim l_filtro
Dim l_orden
Dim l_ternro
Dim l_perfil

l_ternro = request("ternro") ' ternro del logeado
l_perfil = request("perfil")

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY evaproyecto.evaproynro "
end if
%>

<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<head>
<link href="/serviciolocal/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Proyectos - Gesti&oacute;n de Desempeño - RHPro &reg;</title>
</head>

<script>
var jsSelRow = null;

function Deseleccionar(fila)
{
 fila.className = "MouseOutRow";
}
function Seleccionar(fila,cabnro)
{
 if (jsSelRow != null)
 {
  Deseleccionar(jsSelRow);
 };

 document.datos.cabnro.value = cabnro;
 fila.className = "SelectedRow";
 jsSelRow		= fila;
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table width="100%">
    <tr>
        <th>C&oacute;digo de Evento</th>
		<!-- <th>C&oacute;digo</th> -->
        <th>Descripci&oacute;n</th>
        <th>Cliente</th>
        <th>Engagement</th>
        <th>Per&iacute;odo</th>
        <th>Fecha Desde</th>		
        <th>Fecha Hasta</th>		
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT DISTINCT evaproyecto.evaproynro, evaproynom,  "
l_sql = l_sql & " evaproyfdd, evaproyfht, evaclinom, evaengdesabr, evaperdesabr, evaevenro "
l_sql = l_sql & " FROM evaproyecto "
l_sql = l_sql & " INNER JOIN evaevento ON evaevento.evaproynro = evaproyecto.evaproynro"
l_sql = l_sql & " LEFT  JOIN evaproyemp ON evaproyemp.evaproynro = evaproyecto.evaproynro "
l_sql = l_sql & " LEFT  JOIN evaperiodo ON evaproyecto.evapernro = evaperiodo.evapernro "
  'l_sql = l_sql & " LEFT JOIN empleado ON empleado.ternro = evaproyemp.ternro "
l_sql = l_sql & " INNER JOIN evaengage  ON evaengage.evaengnro = evaproyecto.evaengnro "
l_sql = l_sql & " INNER JOIN evacliente ON evacliente.evaclinro = evaengage.evaclinro "
l_sql = l_sql & " WHERE ( evaproyecto.proyrevisor =  " & l_ternro
l_sql = l_sql & "   OR  evaproyecto.proygerente =  " & l_ternro
l_sql = l_sql & "   OR  evaproyecto.proysocio   =  " & l_ternro
l_sql = l_sql & "   OR  evaproyecto.proyaux1 =  " & l_ternro
l_sql = l_sql & "   OR  evaproyecto.proyaux2 =  " & l_ternro
l_sql = l_sql & "	OR  evaproyemp.ternro =  " & l_ternro & ")"
if l_filtro <> "" then
  l_sql = l_sql & " AND " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
	'Response.Write l_sql
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="7">No hay Proyectos.</td>
</tr>
<%
else
	l_evaluacs = 0
	do until l_rs.eof
		l_mostrar=0 
		'mostrar si aun no tiene evaluaciones asociadas o si las evaluaciones
		'NO estan aprobadas
		Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT cabaprobada FROM evacab "
		l_sql = l_sql & " INNER JOIN evaevento ON evaevento.evaevenro=evacab.evaevenro "
		l_sql = l_sql & "        AND evaevento.evaproynro=" & l_rs("evaproynro")
		rsOpen l_rs1, cn, l_sql, 0 
		if not l_rs1.eof then
			do while not l_rs1.eof
				if l_rs1("cabaprobada")<>-1 then
				  l_mostrar=-1
				end if 
			l_rs1.MoveNext
			loop
		else
			l_mostrar=-1
		end if
		l_rs1.Close
		set l_rs1=nothing
		if l_mostrar=-1 then 
			l_evaluacs = -1 	%>
	    <tr ondblclick="Javascript:parent.abrirVentana('proyecto_eva_ag_02.asp?Tipo=M&cabnro=' + datos.cabnro.value,'',500,500)" onclick="Javascript:Seleccionar(this,<%= l_rs("evaproynro")%>)">
			<td width="10%" align="right"><%= l_rs("evaevenro")%>&nbsp;</td>
	        <!-- <td width="10%" align="right"><%'= l_rs("evaproynro")%> &nbsp;</td>  -->
	        <td width="20%" nowrap><%= l_rs("evaproynom")%> </td>
	        <td width="20%" nowrap><%= l_rs("evaclinom")%> </td>
	        <td width="20%" nowrap><%= l_rs("evaengdesabr")%> </td>
	        <td width="20%" nowrap><%= l_rs("evaperdesabr")%> </td>
	        <td width="10%" nowrap align="center"><%= l_rs("evaproyfdd")%> </td>
	        <td width="10%" nowrap align="center"><%= l_rs("evaproyfht")%> </td>
	    </tr>
<%		end if
		l_rs.MoveNext
	loop
	
	if l_evaluacs = 0  then %>
	<tr>
		 <td colspan="7"> No existen evaluaciones asociadas o finalizaron las evaluaciones.</td>
	</tr>
<% 	end if
end if

l_rs.Close
set l_rs = Nothing
cn.Close
set cn = Nothing
%>
</table>
<form name="datos" method="post">
<input type="Hidden" name="cabnro" value="0">
<input type="Hidden" name="orden" value="<%=l_orden%>">
<input type="Hidden" name="filtro" value="<%=l_filtro%>">
</form>
</body>
</html>

