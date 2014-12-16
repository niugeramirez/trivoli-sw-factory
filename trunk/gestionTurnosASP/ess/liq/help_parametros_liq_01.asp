<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--
Archivo: help_parametros_liq_01.asp
Descripción: Ayuda de parametros
Autor : FFavre
Fecha: 10/2003
Modificado: 25-10-05 - Leticia A. - Adecuacion a Autogestion 
			26-10-05 - Leticia A. - Si se config el ConfRep, mostrar los parametros configurados
-->
<% 
 Dim l_rs
 Dim l_sql
 Dim l_filtro
 Dim l_orden
 Dim l_concnro
 Dim l_repnro 
 Dim l_sql_confrep
 
 ' **************
 l_repnro = 150
 
 
 l_concnro = request("concnro")
 l_filtro = request("filtro")
 l_orden  = request("orden")
 
 if l_orden = "" then
 	l_orden = " ORDER BY tipopar.tpanro "
 end if
 
' __________________________________________________________________
' Verificar si se cargaron Conceptos/param a mostrar en el ConfRep  
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 l_sql = " SELECT repnro FROM confrep WHERE repnro=" & l_repnro
 rsOpen l_rs, cn, l_sql, 0 
 
 l_sql_confrep = ""
 if not l_rs.eof then  	' AND confrep.conftipo = 'CO' ?? va
 	 l_sql_confrep = " INNER JOIN confrep  ON confrep.confval = tipopar.tpanro AND confrep.repnro="& l_repnro
	 l_sql_confrep = l_sql_confrep & " INNER JOIN concepto ON UPPER(confrep.confval2)=UPPER(concepto.conccod) AND concepto.concnro="& l_concnro 
 end if 
 l_rs.Close
 
 set l_rs = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Conceptos - Liquidación de Haberes - RHPro &reg;</title>
</head>

<script>
var jsSelRow = null;

function Deseleccionar(fila){
	fila.className = "MouseOutRow";
}

function Seleccionar(fila, cabnro, tpadabr, unisigla){
	if (jsSelRow != null)
		Deseleccionar(jsSelRow);

	document.datos.cabnro.value = cabnro;
	document.datos.tpadabr.value = tpadabr;
	document.datos.unisigla.value = unisigla;
	fila.className = "SelectedRow";
	jsSelRow	   = fila;
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th>C&oacute;digo</th>
        <th>Descripci&oacute;n</th>
    </tr>
<%
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 l_sql = "SELECT DISTINCT tipopar.tpanro, tipopar.tpadabr, unidad.unisigla "
 l_sql = l_sql & " FROM cft_resumen INNER JOIN tipopar ON cft_resumen.tpanro = tipopar.tpanro "
 l_sql = l_sql & " INNER JOIN unidad ON tipopar.uninro = unidad.uninro "
  if l_sql_confrep <> "" then
 	l_sql = l_sql & l_sql_confrep
 end if
 l_sql = l_sql & " WHERE cft_resumen.carind = -1 AND cft_resumen.concnro = " & l_concnro
 if l_filtro <> "" then
 	l_sql = l_sql & " AND " & l_filtro & " "
 end if
 l_sql = l_sql & l_orden

 rsOpen l_rs, cn, l_sql, 0 
 if l_rs.eof then%>
	<tr>
		<td colspan="4">No existen Par&aacute;metros</td>
	</tr>
<%
 else
	do until l_rs.eof
	%>
	    <tr ondblclick="Javascript:parent.Pasar_Valor()" onclick="Javascript:Seleccionar(this, <%= l_rs("tpanro")%>, '<%= l_rs("tpadabr")%>', '<%= l_rs("unisigla")%>')">
	        <td width="20%" nowrap align="right"><%= l_rs("tpanro")%></td>
	        <td width="80%" nowrap><%= l_rs("tpadabr")%></td>
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
<input type="Hidden" name="tpadabr" value="0">
<input type="Hidden" name="unisigla" value="0">
<input type="Hidden" name="orden" value="<%= l_orden %>">
<input type="Hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
