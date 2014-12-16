<% Option Explicit %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<% 
'Archivo: productos_con_02.asp
'Descripción: Abm de Vendedores habilitados a descargar sin contrato
'Autor : Lisandro Moro
'Fecha: 10/02/2005

'Datos del formulario
'
on error goto 0

Dim l_tipo
Dim l_vencornro
Dim l_vencordes
Dim l_vencorrazsoc
Dim l_pronro

Dim l_claseVen
Dim l_clasePro
Dim l_vensel
Dim l_prosel
'ADO
Dim l_sql
Dim l_rs

l_tipo = request.querystring("tipo")
l_vencornro = request.querystring("cabnro")
l_pronro = request.querystring("pronro")
%>
<html>
<head>
<link href="/ticket/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Vendedores Habilitados a Descargar sin contrato - Ticket</title>
</head>
<script src="/ticket/shared/js/fn_ayuda.js"></script>
<script src="/ticket/shared/js/fn_windows.js"></script>
<script>
function Valida(){
	<% If l_tipo = "A" OR l_tipo = "M"then%>
	if (document.datos.vencornro.value == "0"){
		alert('Debe seleccionar un Vendedor.');
		document.datos.vencornro.selected;
		document.datos.vencornro.focus();
		return;
	}
	if (document.datos.pronro.value == "0"){
		alert('Debe seleccionar un Producto.');
		document.datos.pronro.selected;
		document.datos.pronro.focus();		
		return;
	}
	abrirVentanaH('vendedores_sin_contrato_con_03.asp?tipo=<%= l_tipo %>&cabnro=' + document.datos.vencornro.value + '&pronro='+ document.datos.pronro.value ,'',520,160);
	<% Else  %>
	window.close();
	<% End If %>
}

function Setear(){
	var tipo = '<%= l_tipo %>';
	if (tipo == 'A'){
		document.datos.vencornro.focus();
	}else{
		document.datos.pronro.focus();
	}
}
</script>
<% 
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT vencornro, vencordes, vencorrazsoc"
l_sql = l_sql & " FROM tkt_vencor "
l_sql = l_sql & " WHERE venact = -1 AND vencornro = " & l_vencornro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_vencornro = l_rs("vencornro")
	l_vencordes = l_rs("vencordes")
	l_vencorrazsoc = l_rs("vencorrazsoc")
end if
l_rs.Close
Select case l_tipo
	case "A"
		l_claseVen = ""
		l_clasePro = ""
		l_vensel = 0
		l_prosel = 0
	case "M"		
		l_claseVen = " class=""deshabinp"" disabled"
		l_clasePro = ""
		l_vensel = l_vencornro
		l_prosel = l_pronro
	case "C"
		l_claseVen = " class=""deshabinp"" disabled"
		l_clasePro = " class=""deshabinp"" disabled"
		l_vensel = l_vencornro
		l_prosel = l_pronro
end Select

%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="Setear();">
<form name="datos">
	<input type="Hidden" name="tipo" value="<%= l_tipo %>">
<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2"  nowrap>Vendedores Habilitados a Descargar sin Contrato</td>
	<td class="th2" align="right">
		<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</td>
</tr>
<tr>
	<td colspan="2" height="100%">
		<table border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td width="50%"></td>
				<td>
					<table cellspacing="0" cellpadding="0" border="0">
						<tr>
						    <td height="100%" align="right" nowrap><b>Vendedor:</b></td>
							<td>
								<select name="vencornro" size="1" style="width:300;" <%= l_claseVen %>>
									<option value=0 selected>&laquo; Seleccione un Vendedor &raquo;</option>
									<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
									l_sql = "SELECT vencornro, vencorcod,vencordes "
									l_sql  = l_sql  & " FROM tkt_vencor "
									l_sql  = l_sql  & " WHERE venhab = -1 and vencortip='V'"
									l_sql  = l_sql  & " ORDER BY vencordes "
									rsOpen l_rs, cn, l_sql, 0
									do until l_rs.eof		%>	
									<option value=<%= l_rs("vencornro") %> > 
									<%= l_rs("vencordes") %> (<%=l_rs("vencorcod")%>)</option>
									<%	l_rs.Movenext
									loop
									l_rs.Close %>
								</select>
								<script> document.datos.vencornro.value=<%= l_vensel %>;</script>
							</td>
						</tr>
						<tr>
						    <td height="100%" align="right" nowrap><b>Producto:</b></td>
							<td height="100%">
								<select name="pronro" size="1" style="width:300;" <%= l_clasePro %>>
									<option value=0 selected>&laquo; Seleccione un Producto &raquo;</option>
									<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
									l_sql = "SELECT pronro, prodes, procod "
									l_sql  = l_sql  & " FROM tkt_producto "
  								    l_sql  = l_sql  & " WHERE proest <> 0 "							
									l_sql  = l_sql  & " ORDER BY prodes "
									rsOpen l_rs, cn, l_sql, 0
									do until l_rs.eof		%>	
									<option value=<%= l_rs("pronro") %> > 
									<%= l_rs("prodes") %> (<%=l_rs("procod")%>) </option>
									<%	l_rs.Movenext
									loop
									l_rs.Close %>
								</select>
								<script> document.datos.pronro.value= "<%= l_prosel %>"</script>
							</td>
						</tr>
					</table>
				</td>
				<td width="50%"></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
    <td colspan="2" align="right" class="th2">
		<% call MostrarBoton ("sidebtnABM", "Javascript:Valida();","Aceptar")%>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>
</table>
</form>
<%
set l_rs = nothing
Cn.Close
'Cn = nothing
%>
</body>
</html>