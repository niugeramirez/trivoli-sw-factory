<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--
Archivo: help_conceptos_liq_01.asp
Descripción: Ayuda de conceptos
Autor : FFavre
Fecha: 10/2003
Modificado:
	25-11-03 FFavre Se agregaron periodos retroactivos.
	04-02-04 FFavre Actualiza la cantidad de decimales definidos para el concepto.
	25-10-05 - Leticia A. - Adecuacion a Autogestion 
	26-10-05 - Leticia A. - Si se config el ConfRep, mostrar los conceptos configurados
-->
<% 
 Dim l_rs
 Dim l_sql
 Dim l_filtro
 Dim l_orden
 Dim l_repnro 
 Dim l_sql_confrep
 
 ' ********
 l_repnro = 150
 
 l_filtro = request("filtro")
 l_orden  = request("orden")
 
 if l_orden = "" then
 	l_orden = " ORDER BY concepto.conccod "
 end if

' ____________________________________________________________
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
 
 set l_rs = Nothing
 
%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Conceptos - Liquidaci&oacute;n de Haberes - RHPro &reg;</title>
</head>

<script>
var jsSelRow = null;

function Deseleccionar(fila){
	fila.className = "MouseOutRow";
}

function Seleccionar(fila, cabnro, conccod, concabr, concretro, tpanro, tpadabr, unisigla, conccantdec){
	if (jsSelRow != null)
		Deseleccionar(jsSelRow);

	document.datos.cabnro.value = cabnro;
	document.datos.conccod.value = conccod;
	document.datos.concabr.value = concabr;
	document.datos.concretro.value = concretro;
	document.datos.tpanro.value = tpanro;
	document.datos.tpadabr.value = tpadabr;
	document.datos.unisigla.value = unisigla;
	document.datos.conccantdec.value = conccantdec;
	fila.className = "SelectedRow";
	jsSelRow	   = fila;
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th>C&oacute;digo</th>
        <th>Concepto</th>
        <th>F&oacute;rmula</th>
		<th>Par&aacute;metro</th>
    </tr>
<%
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 l_sql = "SELECT concepto.concnro, concepto.conccod, concepto.concabr, concepto.concretro, concepto.conccantdec, "
 l_sql = l_sql & " formula.fordabr, tipopar.tpadabr, tipopar.tpanro, unidad.unisigla "
 l_sql = l_sql & " FROM concepto INNER JOIN cft_resumen ON concepto.concnro = cft_resumen.concnro "
 l_sql = l_sql & " INNER JOIN formula ON concepto.fornro = formula.fornro "
 l_sql = l_sql & " INNER JOIN tipopar ON cft_resumen.tpanro = tipopar.tpanro "
 l_sql = l_sql & " INNER JOIN unidad ON tipopar.uninro = unidad.uninro "
 if l_sql_confrep <> "" then
 	l_sql = l_sql & l_sql_confrep
 end if
 l_sql = l_sql & " WHERE cft_resumen.carind = -1 "
 if l_filtro <> "" then
 	l_sql = l_sql & " AND " & l_filtro & " "
 end if
 l_sql = l_sql & l_orden
 
 rsOpen l_rs, cn, l_sql, 0 
 
  if l_rs.eof then%>
	<tr>
		<td colspan="4">No existen Conceptos</td>
	</tr>
<%
 else
	do until l_rs.eof
		
	%>
	    <tr ondblclick="Javascript:parent.Pasar_Valor()" onclick="Javascript:Seleccionar(this, <%= l_rs("concnro")%>, '<%= l_rs("conccod")%>', '<%= l_rs("concabr")%>', <%= l_rs("concretro")%>, <%= l_rs("tpanro")%>, '<%= l_rs("tpadabr")%>', '<%= l_rs("unisigla")%>', <%= l_rs("conccantdec")%>)">
	        <td width="15%" nowrap align="right"><%= l_rs("conccod")%></td>
	        <td width="35%" nowrap><%= l_rs("concabr")%></td>
	        <td width="25%" nowrap><%= l_rs("fordabr")%></td>
			<td width="25%" nowrap><%= l_rs("tpadabr")%></td>
	    </tr>
	<%
		l_rs.MoveNext
	loop
 end if
 l_rs.Close
 set l_rs = Nothing
 cn.Close
 set cn = Nothing
%>
</table>
<form name="datos" method="post">
<input type="Hidden" name="cabnro" value="0">
<input type="Hidden" name="conccod" value="0">
<input type="Hidden" name="concabr" value="0">
<input type="Hidden" name="concretro" value="0">
<input type="Hidden" name="tpanro" value="0">
<input type="Hidden" name="tpadabr" value="0">
<input type="Hidden" name="unisigla" value="0">
<input type="Hidden" name="conccantdec" value="0">
<input type="Hidden" name="orden" value="<%= l_orden %>">
<input type="Hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
