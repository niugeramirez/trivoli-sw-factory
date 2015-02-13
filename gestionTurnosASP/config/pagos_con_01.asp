<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 
'Archivo: companies_con_01.asp
'Descripción: ABM de Companies
'Autor : Raul Chinestra
'Fecha: 26/11/2007

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden

Dim l_idobrasocial

Dim l_idpracticarealizada
Dim l_sumapago

l_idpracticarealizada = request("idpracticarealizada") 

l_filtro = request("filtro")
l_orden  = request("orden")
l_idobrasocial = request("idobrasocial")

if l_orden = "" then
  l_orden = " ORDER BY fecha "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_confirm.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<head>
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Detalle de Pagos</title>
</head>

<script>
var jsSelRow = null;

function Deseleccionar(fila){
	fila.className = "MouseOutRow";
}

function Seleccionar(fila,cabnro){
	if (jsSelRow != null){
		Deseleccionar(jsSelRow);
	};
	document.datos.cabnro.value = cabnro;
	fila.className = "SelectedRow";
	jsSelRow = fila;
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" >
<table>
    <tr>
      <th nowrap>Fecha</th>			
      <th>Medio Pago</th>
	  <th>Obra Social</th>
	  <th>Importe</th>
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT pagos.id, pagos.* , mediosdepago.titulo , obrassociales.descripcion"
l_sql = l_sql & " FROM pagos "
l_sql = l_sql & " LEFT JOIN  mediosdepago ON mediosdepago.id = pagos.idmediodepago "
l_sql = l_sql & " LEFT JOIN  obrassociales ON obrassociales.id = pagos.idobrasocial "

l_sql = l_sql & " WHERE idpracticarealizada = " & l_idpracticarealizada
if l_filtro <> "" then
  l_sql = l_sql & " AND " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="4" >No existen Pagos cargados.</td>
</tr>
<%else
	l_sumapago = 0
	do until l_rs.eof
		l_sumapago = l_sumapago + cdbl(l_rs("importe"))
	%>
	    <tr ondblclick="Javascript:parent.abrirVentana('pagos_con_02.asp?Tipo=M&idpracticarealizada=<%= l_idpracticarealizada %>&cabnro=' + datos.cabnro.value,'',520,200)" onclick="Javascript:Seleccionar(this,<%= l_rs("id")%>)">
	        <td width="15%" align="left" nowrap><%= l_rs("fecha")%></td>
	        <td width="30%" nowrap><%= l_rs("titulo")%></td>		
			<td width="30%" nowrap><%= l_rs("descripcion")%></td>		
			<td width="25%" align="right" ><%= l_rs("importe")%></td>			
	    </tr>
	<%
		l_rs.MoveNext
	loop
	%>
	    <tr>
	        <td width="15%" align="left" nowrap>&nbsp;</td>
	        <td width="30%" nowrap>&nbsp;</td>		
			<td width="30%" nowrap><b>TOTAL</b></td>		
			<td width="25%" align="right" ><b><%= l_sumapago %></b></td>			
	    </tr>
	<%	
end if
l_rs.Close
set l_rs = Nothing
cn.Close
set cn = Nothing
%>
</table>
<form name="datos" method="post">
<input type="hidden" name="cabnro" value="0">
<input type="hidden" name="orden" value="<%= l_orden %>">
<input type="hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
