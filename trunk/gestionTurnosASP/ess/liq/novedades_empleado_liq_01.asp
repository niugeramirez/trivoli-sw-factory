<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
Archivo: novedades_empleado_liq_01.asp
Descripción: abm de novedades del empleado
Autor: Fernando Favre
Fecha: 10-10-03
Modificado: 
	13-11-03 FFavre Se agrego la sigla al valor
	15-11-03 FFavre Se saco que la novedad sea 'auto'
	25-11-03 FFavre Se agregaron periodos retroactivos.
	12-02-04 FFavre Se muestra el valor con la cant. de decimales definidos para el concepto.
    03-09-04 Scarpa D. Pasar como parametro el nenro	
	25-10-05 - Leticia A. - Adecuacion a Autogestion 
	27-10-05 - Leticia A. - Si se configuro el ConfRep, mostrar los conceptos configurados.
-->
<%
 Dim l_rs
 Dim l_rs1
 Dim l_sql
 Dim l_filtro
 Dim l_orden
 Dim l_ternro
 Dim l_empleg
 Dim l_repnro 
 Dim l_sql_confrep
 
 ' ************
 l_repnro = 150
 
 l_filtro = request("filtro")
 l_orden  = request("orden")
 
 l_empleg = request("empleg")
 l_ternro = l_ess_ternro  
 
 if l_orden = "" then
 	l_orden = " ORDER BY concepto.conccod ASC "
 end if
 
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
 
 
' ______________________________________________________________
' Verificar si se cargaron Conceptos a mostrar en el ConfRep    
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 l_sql = " SELECT repnro FROM confrep WHERE repnro=" & l_repnro
 rsOpen l_rs, cn, l_sql, 0 
 
 l_sql_confrep = ""
 if not l_rs.eof then  	' AND confrep.conftipo = 'CO' ?? va
	 l_sql_confrep = " INNER JOIN confrep ON UPPER(confrep.confval2)=UPPER(concepto.conccod) AND confrep.confval = tipopar.tpanro"
	 l_sql_confrep = l_sql_confrep & " AND confrep.repnro="& l_repnro
 end if 
 l_rs.Close
 
 'set l_rs = Nothing
 ' _____________________________________________________________
%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Novedades por Empleado - Liquidaci&oacute;n de Haberes - RHPro &reg;</title>
</head>

<script>
var jsSelRow = null;

function Deseleccionar(fila){
	fila.className = "MouseOutRow";
}

function Seleccionar(fila, cabnro, concnro, tpanro){
	if (jsSelRow != null){
  		Deseleccionar(jsSelRow);
	};
	document.datos.cabnro.value = cabnro;
	document.datos.concnro.value = concnro;
	document.datos.tpanro.value = tpanro;
	fila.className = "SelectedRow";
	jsSelRow		= fila;
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos" method="post">

<table>
	<tr>
		<th><b>C&oacute;digo</b></th>
		<th><b>Concepto</b></th>
		<th><b>Par&aacute;metro</b></th>
		<th><b>Valor</b></th>
		<th><b>Unidad</b></th>
		<th><b>Depurable</b></th>
		<th><b>Vigencia</b></th>
		<th><b>Desde</b></th>
		<th><b>Hasta</b></th>
		<th><b>Retroactivo</b></th>
		<th><b>Desde</b></th>
		<th><b>Hasta</b></th>
	</tr>
	<%
	if l_ternro <> "" then
		l_sql = "SELECT novemp.nenro, novemp.empleado, novemp.concnro, novemp.tpanro, concepto.conccod, "
		l_sql = l_sql & "concepto.concabr, concepto.conccantdec, tipopar.tpadabr, novemp.nevalor, "
		l_sql = l_sql & "periodo.pliqdesc AS nepliqdesdedesc, periodo2.pliqdesc as nepliqhastadesc, "
		l_sql = l_sql & "novemp.nevigencia, novemp.nedesde, novemp.nehasta, novemp.neretro, concepto.fornro, unidad.unisigla "
		l_sql = l_sql & " FROM novemp INNER JOIN concepto ON novemp.concnro = concepto.concnro "
		l_sql = l_sql & " INNER JOIN tipopar ON novemp.tpanro = tipopar.tpanro "
		l_sql = l_sql & " INNER JOIN unidad ON tipopar.uninro = unidad.uninro "
		if l_sql_confrep <> "" then
			l_sql = l_sql & l_sql_confrep
		end if
		l_sql = l_sql & " LEFT JOIN periodo ON periodo.pliqnro=novemp.nepliqdesde "
		L_SQL = L_SQL & " LEFT JOIN periodo periodo2 ON periodo2.pliqnro=novemp.nepliqhasta "
		l_sql = l_sql & " WHERE novemp.empleado = " & l_ternro & " "
						  
		if l_filtro <> "" then
		  l_sql = l_sql & "AND " & l_filtro & " "
		end if
		
		l_sql = l_sql & l_orden
		
		rsOpen l_rs, cn, l_sql, 0 
		
		if l_rs.eof then
		%>
			<td colspan="12">No tiene novedades</td>
		<%
		else
			do until l_rs.eof
			%>
			<!-- ondblclick="Javascript:parent.abrirVentanaVerif('novedades_empleado_liq_02.asp?tipo=M&empleg=<%'=l_empleg %>&concnro=<%'=l_rs("concnro")%>&tpanro=<%'=l_rs("tpanro")%>&nenro=<%'=l_rs("nenro")%>','',630,300)" -->
			<tr onclick="Javascript:Seleccionar(this,<%=l_rs("nenro")%>, <%=l_rs("concnro")%>, <%=l_rs("tpanro")%>)">
				<td><%= l_rs("conccod")%></td>
				<td><%= l_rs("concabr")%></td>
				<td><%= l_rs("tpadabr")%></td>
				<td align="right"><%= formatnumber(l_rs("nevalor"), l_rs("conccantdec"))%></td>
				<td><%= l_rs("unisigla")%></td>
				<%
				l_sql = "SELECT * "
				l_sql = l_sql & "FROM con_for_tpa "
				l_sql = l_sql & "WHERE con_for_tpa.concnro = " & l_rs("concnro") & " "
				l_sql = l_sql & "AND con_for_tpa.tpanro = " & l_rs("tpanro") & " "
				l_sql = l_sql & "AND con_for_tpa.fornro = " & l_rs("fornro") & " "
				l_sql = l_sql & "AND con_for_tpa.depurable = -1"
				rsOpen l_rs1, cn, l_sql, 0
				%>
				<td align="center"><% if not l_rs1.eof then%>S&iacute;<%else%>No<%end if%></td>
				<td align="center"><% if l_rs("nevigencia") then%>S&iacute;<%else%>No<%end if%></td>
				<td align="center"><%= l_rs("nedesde")%></td>
				<td align="center"><%= l_rs("nehasta")%></td>
				<td align="center"><%if (l_rs("nepliqdesdedesc") <> "" and l_rs("nepliqdesdedesc") <> "") then%>S&iacute;<%else%>No<%end if%></td>
				<td align="center"><%= l_rs("nepliqdesdedesc")%></td>
				<td align="center"><%= l_rs("nepliqhastadesc")%></td>
				<%
				l_rs1.close
				%>
			</tr>
				<%
				l_rs.MoveNext
			loop
		end if
		
		l_rs.Close
	else
		%>
		<td colspan="12">No tiene novedades</td>
		<%
	end if
	set l_rs = Nothing
	set l_rs1 = Nothing
	cn.Close
	set cn = Nothing
	%>
</table>

<input type="Hidden" name="cabnro" value="">
<input type="Hidden" name="concnro" value="">
<input type="Hidden" name="tpanro" value="">
<input type="Hidden" name="orden" value="<%= l_orden %>">
<input type="Hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
