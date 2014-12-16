<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : noved_horarias_gti_01.asp
Descripcion    : Modulo que se encarga de listar las nov horarias
Modificacion   :
    07/10/2003 - Scarpa D. - Considerar el motivo nulo en la consulta SQL
    10/11/2003 - Scarpa D. - Cambio en el filtro de la SQL
	06/10/2005 - Leticia A. - 
-----------------------------------------------------------------------------
-->
<% 
on error goto 0

Dim l_rs
Dim l_sql
Dim l_gnovhoradesde
Dim l_gnovhorahasta

Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden

Dim l_fechadesde
Dim l_fechahasta
Dim l_ternro


l_filtro = request("filtro")
l_orden  = request("orden")

l_ternro = l_ess_ternro
l_fechadesde = request.querystring("fechadesde")
l_fechahasta  = request.querystring("fechahasta")

if l_orden = "" then
  l_orden = 3
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Licencias - Gesti&oacute;n de Tiempos - RHPro &reg;</title>
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
    <tr nowrap>
        <th align="left">Descripci&oacute;n</th>
        <th align="left">Tipo de Novedad</th>
        <th align="center">Fecha Desde</th>
        <th align="center">Hora</th>
        <th align="center">Fecha Hasta</th>
        <th align="center">Hora</th>
        <th align="left">Motivo</th>
    </tr>
<%
select case l_orden
  case "-1"  l_sqlorden = "ORDER BY gnovdesabr DESC"
  case "-2"  l_sqlorden = "ORDER BY gtnovdesabr DESC"
  case "-3"  l_sqlorden = "ORDER BY gnovdesde DESC"
  case "-4"  l_sqlorden = "ORDER BY gnovhasta DESC"
  case "1"  l_sqlorden = "ORDER BY gnovdesabr "
  case "2"  l_sqlorden = "ORDER BY gtnovdesabr "
  case "3"  l_sqlorden = "ORDER BY gnovdesde "
  case "4"  l_sqlorden = "ORDER BY gnovhasta "
  case else l_sqlorden = ""
end select

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT gnovnro, gnovdesabr, empleado.terape, empleado.ternom, gti_tiponovedad.gtnovdesabr, gti_motivo.motdesabr "
l_sql = l_sql & ", gnovdiacompleto, gnovdesde, gnovhasta, gnovhoradesde, gnovhorahasta, gnovorden, gnovmaxhoras "
l_sql = l_sql & "FROM gti_novedad "
l_sql = l_sql & "INNER JOIN gti_tiponovedad ON gti_novedad.gtnovnro=gti_tiponovedad.gtnovnro "
l_sql = l_sql & "LEFT  JOIN gti_motivo ON gti_novedad.motnro=gti_motivo.motnro "
l_sql = l_sql & "INNER JOIN empleado ON gti_novedad.gnovotoa=empleado.ternro "
l_sql = l_sql & "WHERE ternro= " & l_ternro  & " AND "

Dim des
Dim has

des = cambiafecha(l_fechadesde, "YMD", true)
has = cambiafecha(l_fechahasta, "YMD", true)

'Nomenclatura:
'  [ ] = intervalo de la novedad
'  ( ) = caso del intervalo que controlo

' -----[--(_)--]------
l_sql = l_sql & " ( ( gnovdesde <= " & des & " AND gnovhasta >= " & has & " ) "

' --(__[___)---]------
l_sql = l_sql & " OR ( gnovdesde <= " & has & " AND gnovhasta >= " & has & " ) "

' -----[---(___]__)---
l_sql = l_sql & " OR ( gnovdesde <= " & des & " AND gnovhasta >= " & des & " ) "

' --(__[_______]__)---
l_sql = l_sql & " OR ( gnovdesde >= " & des & " AND gnovhasta <= " & has & " ) ) "

if l_filtro <> "" then
  l_sql = l_sql & " AND " & l_filtro & " "
end if
l_sql = l_sql & l_sqlorden
'response.write l_sql
rsOpen l_rs, cn, l_sql, 0 
do until l_rs.eof 
	l_gnovhoradesde = l_rs("gnovhoradesde")
	l_gnovhorahasta = l_rs("gnovhorahasta")
%>
    <tr onclick="Javascript:Seleccionar(this,<%= l_rs("gnovnro")%>)">
        <td nowrap align="left"><%= l_rs("gnovdesabr")%></td>
		<td nowrap align="left"><%= l_rs("gtnovdesabr")%></td>
        <td  nowrap align="center"><%= l_rs("gnovdesde")%></td>
        <td nowrap align="center"><%= mid(l_gnovhoradesde,1,2)&":"& mid(l_gnovhoradesde,3,2)%></td>    
        <td nowrap align="center"><%= l_rs("gnovhasta")%></td>
        <td nowrap align="center"><%= mid(l_gnovhorahasta,1,2)&":"& mid(l_gnovhorahasta,3,2)%></td>    
        <td nowrap align="center"><%= l_rs("motdesabr")%></td>
	</tr>
<%
	l_rs.MoveNext
loop
l_rs.Close

set l_rs = Nothing
cn.Close
set cn = Nothing
%>
</table>
<% if l_fechadesde<>"" then %>
	<script> parent.datos.fechadesde.value="<%= l_fechadesde %>";</script>
<% end if %>
<% if l_fechahasta<>"" then %>
	<script> parent.datos.fechahasta.value="<%= l_fechahasta %>";</script>
<% end if %>

<form name="datos" method="post">
<input type="Hidden" name="cabnro" value="0">
<input type="Hidden" name="orden" value="<%= l_orden %>">
<input type="Hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
