<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--
Modificado: 12-10-2005 - Leticia A. - Adaptarlo a Autogestion.
Modificado: 31-10-2005 - CCR. - Adaptarlo a Autogestion.
-->
<%
'variables 
 Dim l_rs
 Dim l_sql
 Dim l_ternrologeado
 Dim l_ternroactual
'parametros entrada

'de uso local
 dim l_filtro
 dim l_filtro2
 dim l_orden
 dim l_emplegactual
 dim l_empleglogeado
 
 
l_filtro = request("filtro")
l_orden  = request("orden")
l_emplegactual   = request("empleg")
l_empleglogeado  = Session("empleg")

if l_orden = "" then
	l_orden = " ORDER BY evaevento.evaevenro"
end if

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT ternro  from empleado WHERE empleg = " 	& l_empleglogeado
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_ternrologeado= l_rs("ternro")
end if
l_rs.close
l_sql = "SELECT ternro  from empleado WHERE empleg = " 	& l_emplegactual
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_ternroactual= l_rs("ternro")
end if
l_rs.close
set	 l_rs=nothing%>

<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" http-equiv="refresh" content="text/html; charset=iso-8859-1">
<title><%if ccodelco=-1 then%>Eventos del Ciclo de Gesti&oacute;n del Desempeño<%else%>Evento de Evaluaci&oacute;n<%end if%> - Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
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
<table>
    <tr>
        <th>C&oacute;digo</th>
        <th>Descripci&oacute;n</th>
        <th>Tipo de Evento</th>
        <th>Formulario</th>
        <th>Fecha Eval.</th>
        <th>Fecha Desde</th>
        <th>Fecha Hasta</th>
    </tr>
<%


Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT DISTINCT evaevento.evaevenro, evaevento.evaevedesabr, evatipoevento.evatipevedabr, evatipoeva.evatipdesabr, evaevento.evaevefecha, evaevento.evaevefdesde, evaevento.evaevefhasta FROM  evaevento "
l_sql = l_sql & " INNER JOIN evatipoeva    ON evatipoeva.evatipnro = evaevento.evatipnro INNER JOIN evatipoevento ON evatipoevento.evatipevenro = evaevento.evatipevenro "
l_sql = l_sql & " WHERE "
' no exista ninguna evaluacion asociada
l_sql = l_sql & " ( NOT EXISTS (select * from evacab where evacab.evaevenro=evaevento.evaevenro) )"
' que exista al menos una cuyo revisor (supervisor) sea el logeado
l_sql = l_sql & " OR  "
l_sql = l_sql & " ( EXISTS (select * from evacab inner join evadetevldor on evadetevldor.evacabnro=evacab.evacabnro and evadetevldor.evatevnro = " & cevaluador &" and evadetevldor.evaluador = " & l_ternrologeado & " where evacab.evaevenro=evaevento.evaevenro  and evacab.empleado=" & l_ternroactual & ") )"
if l_filtro <> "" then
 l_sql = l_sql & " WHERE " & l_filtro 
end if
	
'response.write l_sql
l_sql = l_sql & " " & l_orden	
rsOpen l_rs, cn, l_sql, 0 

if l_rs.eof then%>
<tr>
	 <td colspan="7">No hay eventos de evaluaci&oacute;n.</td>
</tr>
<%else
	do until l_rs.eof
	%>
	<%if ccodelco=-1 then%>
	<tr onclick="Javascript:Seleccionar(this,<%=l_rs("evaevenro")%>)" ondblclick="Javascript:parent.abrirVentanaVerif('evento_evaluacion_02.asp?Tipo=M&evaevenro=<%=l_rs("evaevenro")%>','',505,400)">
	<%else%>
	<tr onclick="Javascript:Seleccionar(this,<%=l_rs("evaevenro")%>)" ondblclick="Javascript:parent.abrirVentanaVerif('evento_evaluacion_02.asp?Tipo=M&evaevenro=<%=l_rs("evaevenro")%>','',505,450)">
	<%end if%>
	    <td nowrap width="10%"><%=l_rs("evaevenro")%></td>
		<td nowrap width="20%"><%=l_rs("evaevedesabr")%></td>
		<td nowrap width="20%"><%=l_rs("evatipevedabr")%></td>
		<td nowrap width="20%"><%=l_rs("evatipdesabr")%></td>
		<td nowrap width="10%"><%=l_rs("evaevefecha")%></td>
		<td nowrap width="10%"><%=l_rs("evaevefdesde")%></td>
		<td nowrap width="10%"><%=l_rs("evaevefhasta")%></td>
	</tr>
	<%l_rs.MoveNext
	loop
end if ' del if l_rs.eof
l_rs.Close
set l_rs=nothing

cn.Close	

%>
</table>

<form name="datos" method="post">
<input type="Hidden" name="cabnro" value="0" >
<input type="Hidden" name="orden" value="<%= l_orden %>">
<input type="hidden" name="filtro" value="<%= l_filtro2 %>">
</form>

</body>
</html>
