<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<%
'=====================================================================================
'Archivo  : historicas_eva_ag_01.asp
'Objetivo : lista de formularios de evaluacion - cerrados-aprobados - rol:evaluador
'Fecha	  : 12-05-2006
'Autor	  : Leticia A.
'Modificación: 04/10/2006 - Leticia Amadio - Adecuarlo a autogestion
'=====================================================================================
on error goto 0
'variables 
 Dim l_rs
 Dim l_rs1
 Dim l_sql
 dim l_filtro2
 dim l_horaapro 
 dim l_empleado
 
'parametros
 dim l_filtro
 dim l_orden
 dim l_ternro  
 Dim l_tipoeval ' si es deloitte viene tipo de evaluación


l_filtro = request("filtro")
l_orden  = request("orden")
'l_ternro  = request("ternro") ' viene el ternro del empleg de autogestion
l_ternro =  l_ess_ternro
l_tipoeval = request("tipoeval")

if len(l_filtro) <> 0 then
	if left(l_filtro,1) <> "'" then
		l_filtro2 = "'" & l_filtro & "'"
	else
		l_filtro2 =  mid(l_filtro,2,len(request("filtro")) - 1)
	end if	
end if	

if l_orden = "" then
	l_orden = " ORDER BY evaevento.evaevenro " 'evaevedesabr
end if

	
	
%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" http-equiv="refresh" content="text/html; charset=iso-8859-1">
<title>Hist&oacute;ricos - Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
</head>
<script>
var jsSelRow = null;

function Deseleccionar(fila)
{
 	fila.className = "MouseOutRow";
}

function Seleccionar(fila,cabnro,evaevenro,empleg)
{
 if (jsSelRow != null)
 {
  Deseleccionar(jsSelRow);
 };

 document.datos.cabnro.value = cabnro;
 document.datos.evaevenro.value = evaevenro;
 document.datos.empleg.value = empleg;
 fila.className = "SelectedRow";
 
 jsSelRow		= fila;
}

</script>


<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
		<th>C&oacute;digo <br> de Evento</th>
		<% if cint(cdeloitte) <> -1 then %>
        <th>Evento</th>
		<% end if %>
		<th> Rol </th>
		<th> Evaluado </th>
        <th>Fecha Aprobaci&oacute;n</th>
        <th>Hora Aprobaci&oacute;n</th>
    </tr>
<%


Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT DISTINCT evacab.evacabnro, evaevento.evaevenro,evaevento.evaevedesabr,evatipoeva.evatipdesabr, "
l_sql = l_sql & " empleado.empleg,empleado.terape,empleado.terape2,empleado.ternom,empleado.ternom2, "  
l_sql = l_sql & " evacab.cabaprobada, fechaapro, horaapro, evatipevalua.evatevdesabr "  
l_sql = l_sql & " FROM evacab "
l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro = evacab.evacabnro"
l_sql = l_sql & " INNER JOIN evatipevalua ON evatipevalua.evatevnro = evadetevldor.evatevnro"
l_sql = l_sql & " INNER JOIN evaevento    ON evaevento.evaevenro = evacab.evaevenro "
l_sql = l_sql & " INNER JOIN evatipoeva   ON evatipoeva.evatipnro = evaevento.evatipnro "
l_sql = l_sql & " INNER JOIN empleado   ON  empleado.ternro = evacab.empleado "
'l_sql = l_sql & " WHERE evacab.empleado = " & l_ternro
l_sql = l_sql & " WHERE evadetevldor.evaluador = " & l_ternro
l_sql = l_sql & "    AND evadetevldor.evatevnro <> " & cint(cautoevaluador) 
l_sql = l_sql & "    AND evacab.cabaprobada = -1 "
if l_filtro <> "" then
 l_sql = l_sql & " AND " & l_filtro 
end if
l_sql = l_sql & " " & l_orden
'response.Write l_sql & "<br>"
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="3">No hay<%if cint(ccodelco)=-1 then%>&nbsp;Procesos Aprobados&nbsp;para el Supervisado<%else%>&nbsp;Evaluaciones Aprobadas para el Empleado&nbsp;<%end if%> o Filtro Ingresado.</td>
</tr>
<%else
	do until l_rs.eof
		l_empleado = l_rs("empleg") & " - " & l_rs("terape") & ", " & l_rs("ternom")
		if trim(l_rs("horaapro")) <>"" then
			l_horaapro=left(l_rs("horaapro"),2)&":"&right(l_rs("horaapro"),2)
		end if%>
		<tr onclick="Javascript:Seleccionar(this,<%=l_rs("evacabnro")%>,<%=l_rs("evaevenro")%>,<%=l_rs("empleg")%>)" ondblclick="Javascript:parent.abrirVentanaVerif('form_carga_eva_ag_00.asp?evacabnro=<%=l_rs("evacabnro")%>&evaevenro=<%=l_rs("evaevenro")%>&empleg=<%=l_rs("empleg")%>','',800,600);">
			<td align="right"><%=l_rs("evaevenro")%>&nbsp;</td>
			<% if cint(cdeloitte) <> -1 then %>
			<td nowrap><%=l_rs("evaevedesabr")%></td>
			<% end if %>
			<td nowrap><%=l_rs("evatevdesabr")%></td>
			<td nowrap>&nbsp;<%=l_empleado %> </td>
			<td nowrap align=center><%=l_rs("fechaapro")%></td>
			<td nowrap align=center><%=l_horaapro%></td>
			
			</tr>
		<%
	l_rs.MoveNext
	loop
	
end if ' del if l_rs.eof
l_rs.Close
cn.Close	
%>
</table>

<form name="datos" method="post">
<input type="Hidden" name="cabnro" value="0" >
<input type="Hidden" name="evaevenro" value="0" >
<input type="Hidden" name="empleg" value="0" >
<input type="Hidden" name="orden"  value="<%= l_orden %>">
<input type="hidden" name="filtro" value="<%= l_filtro2 %>">
</form>

</body>
</html>
