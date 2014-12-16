<% Response.AddHeader "Content-Disposition", "attachment;filename=proyectos.xls" %>
<% Response.ContentType = "application/vnd.ms-excel" %>


<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<% 'Option Explicit %>
<%

'================================================================================
'Archivo		: proyecto_eva_ag_excel.asp
'Descripción	: Abm de Proyectos
'Autor			: CCRossi
'Fecha			: 30-08-2004
'Modificado		: 08-05-2005 LAmadio - arreglo para que funcione
' 				: 29-07-2005 - LA. - cambio de codigo proyecto por evento
'================================================================================
on error goto 0

Dim l_rs
Dim l_sql
Dim l_sqlfiltro
Dim l_sqlorden

'parametros
Dim l_filtro
Dim l_orden
Dim l_ternro
Dim l_perfil

Dim l_mostrar

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
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Proyectos - Gesti&oacute;n de Desempeño - RHPro &reg;</title>
</head>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th colspan=7>Proyectos</th>
    </tr>
    <tr>
        <th>C&oacute;digo de Evento</th>
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
l_sql = l_sql & " INNER JOIN evaevento ON evaevento.evaproynro = evaproyecto.evaproynro "
l_sql = l_sql & " LEFT  JOIN evaproyemp ON evaproyemp.evaproynro = evaproyecto.evaproynro "
l_sql = l_sql & " LEFT JOIN evaperiodo ON evaproyecto.evapernro = evaperiodo.evapernro "
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

rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="7">No hay Proyectos.</td>
</tr>
<%else
	do until l_rs.eof
	
		l_mostrar=0 
		'mostrar si aun no tiene evaluaciones asociadas o si las evaluaciones
		'NO estan aprobadas
		Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT cabaprobada  "
		l_sql = l_sql & " FROM evacab "
		l_sql = l_sql & " INNER JOIN evaevento ON evaevento.evaevenro=evacab.evaevenro "
		l_sql = l_sql & "        AND evaevento.evaproynro=" & l_rs("evaproynro")
		rsOpen l_rs1, cn, l_sql, 0 
		if not l_rs1.eof then
			if l_rs1("cabaprobada")<>-1 then
				l_mostrar=-1
			end if
		else
			l_mostrar=-1		
		end if
		l_rs1.Close
		set l_rs1=nothing
		if l_mostrar=-1 then %>
	    <tr>
			<td width="10%" align="right"><%= l_rs("evaevenro")%>&nbsp;</td>
	        <td width="20%" nowrap><%= l_rs("evaproynom")%> </td>
	        <td width="45%" nowrap><%= l_rs("evaclinom")%></td>
	        <td width="45%" nowrap><%= l_rs("evaengdesabr")%></td>
	        <td width="20%" nowrap><%= l_rs("evaperdesabr")%></td> 
	        <td width="45%" nowrap align="center"><%= fechaISO(l_rs("evaproyfdd"))%></td>
	        <td width="45%" nowrap align="center"><%= fechaISO(l_rs("evaproyfht"))%></td>
	    </tr>
	<%
		end if
		l_rs.MoveNext
	loop
end if
l_rs.Close
set l_rs = Nothing
cn.Close
set cn = Nothing
%>
</table>
</body>
</html>
