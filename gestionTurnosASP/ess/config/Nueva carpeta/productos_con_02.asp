<% Option Explicit %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<% 
'Archivo: productos_con_02.asp
'Descripción: Abm de Cámaras
'Autor : Lisandro Moro
'Fecha: 09/02/2005

'Modificada por: Javier Posadas
'Fecha: 05/04/2005
'Descripción: Se agregó la posibilidad de habilitar/deshabilitar Productos

'Datos del formulario
'pronro, procod, prodes, tippronro, proenv, procla, provercon, promez, proest
on error goto 0

Dim l_pronro
Dim l_procod
Dim l_prodes
Dim l_tipprodes
Dim l_proenv
Dim l_procla
Dim l_provercon
Dim l_promez
Dim l_proest
Dim l_envase
Dim l_proacuanu
Dim l_proacumen
Dim l_proacuaux

'ADO
Dim l_sql
Dim l_rs

%>
<html>
<head>
<link href="/ticket/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Productos - Ticket</title>
</head>
<style type="text/css">
.none{
	padding : 0;
	padding-left : 0;
}
</style>
<script src="/ticket/shared/js/fn_ayuda.js"></script>
<script src="/ticket/shared/js/fn_windows.js"></script>
<script>

function Validar(){

	if (document.datos.proacuaux.value == ""){
		document.datos.proacuaux.select();
		alert('Debe ingresar un valor numérico en el Acumulador Auxiliar.');
		document.datos.proacuaux.focus();
		return;
	}	

	if (isNaN(document.datos.proacuaux.value)){
		document.datos.proacuaux.select();
		alert('Debe ingresar un valor numérico en el Acumulador Auxiliar.');
		document.datos.proacuaux.focus();
		return;
	}	

	document.datos.submit();
}


</script>
<% 
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_pronro = request.querystring("cabnro")
l_sql = "SELECT pronro, procod, prodes, tipprodes, proenv, procla, provercon, promez, proest, proacuanu, proacumen, proacuaux "
l_sql = l_sql & " FROM tkt_producto "
l_sql = l_sql & " INNER JOIN tkt_tipoproducto ON tkt_tipoproducto.tippronro =  tkt_producto.tippronro "
l_sql  = l_sql  & " WHERE pronro = " & l_pronro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_pronro    = l_rs("pronro")
	l_procod    = l_rs("procod")
	l_prodes    = l_rs("prodes")
	l_tipprodes = l_rs("tipprodes")
	l_proenv    = l_rs("proenv")
	l_procla    = l_rs("procla")
	l_provercon = l_rs("provercon")
	l_promez    = l_rs("promez")
	l_proest    = l_rs("proest")
	
	select case UCase(l_proenv)
		case "G"
			l_envase = "Granel"
		case "B"
			l_envase = "Bolsa"
		case else
			l_envase = "Ninguno"
	end select
	
	l_proacuanu  = l_rs("proacuanu")
	l_proacumen  = l_rs("proacumen")
	l_proacuaux  = l_rs("proacuaux")
	
	
end if
l_rs.Close
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos" action="productos_con_04.asp" method="post" target="valida">
	<input type="Hidden" name="pronro" value="<%= l_pronro %>">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2"  nowrap>Productos</td>
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
						    <td height="100%" align="right" nowrap><b>Código:</b></td>
							<td height="100%">
								<input type="text" readonly class="deshabinp" name="procod" size="12" maxlength="20" value="<%= l_procod %>">
							</td>
						</tr>
						<tr>
						    <td height="100%" align="right" nowrap><b>Descripción:</b></td>
							<td height="100%">
								<input type="text" readonly class="deshabinp" name="prodes" size="50" maxlength="50" value="<%= l_prodes %>">
							</td>
						</tr>
						<tr>
						    <td height="100%" align="right" nowrap><b>Tipo Producto:</b></td>
							<td height="100%">
								<input type="text" readonly class="deshabinp" name="tippro" size="50" maxlength="50" value="<%= l_tipprodes %>">
							</td>
						</tr>
						<tr>
						    <td height="100%" align="right" nowrap><b>Clasificación:</b></td>
							<td height="100%">
								<input type="text" readonly class="deshabinp" name="procla" size="50" maxlength="50" value="<%= l_procla %>">
							</td>
						</tr>
						<tr>
						    <td height="100%" align="right" nowrap><b>Envase:</b></td>
							<td height="100%">
								<input type="text" readonly class="deshabinp" name="proenv" size="50" maxlength="50" value="<%= l_envase %>">
							</td>
						</tr>
						<tr>
						    <td height="100%" align="right" nowrap><b>Acumulado Anual:</b></td>
							<td height="100%">
								<input type="text" readonly class="deshabinp" name="proacuanu" size="30" maxlength="50" value="<%= l_proacuanu %>">
							</td>
						</tr>						
						<tr>
						    <td height="100%" align="right" nowrap><b>Acumulado Mensual:</b></td>
							<td height="100%">
								<input type="text"  readonly class="deshabinp" name="proacumen" size="30" maxlength="50" value="<%= l_proacumen %>">
							</td>
						</tr>												
						<tr>
						    <td height="100%" align="right" nowrap><b>Acumulado Auxiliar:</b></td>
							<td height="100%">
								<input type="text" name="proacuaux" size="30" maxlength="50" value="<%= l_proacuaux %>">
							</td>
						</tr>																		
						<tr>
						    <td height="100%" align="right" nowrap><b>Verifica Contrato:</b></td>
							<td height="100%">
								<input type="Checkbox" disabled readonly  name="provercon" value="<%= l_provercon %>" <% If l_provercon = -1 then %>Checked<% End If %>>
							</td>
						</tr>
						<tr>
						    <td height="100%" align="right" nowrap><b>Producto Mezcla:</b></td>
							<td height="100%">
								<input type="Checkbox" disabled readonly  name="promez" value="<%= l_promez %>" <% If l_promez = -1 then %>Checked<% End If %>>
							</td>
						</tr>
						<tr>
						    <td height="100%" align="right" nowrap><b>Habilitado:</b></td>
							<td height="100%">
								<input type="Checkbox" disabled readonly  name="proest" value="<%= l_proest %>" <% If l_proest = -1 then %>Checked<% End If %>>
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
		<% call MostrarBoton ("sidebtnABM", "Javascript:Validar();","Aceptar")%>	
		<a class=sidebtnABM href="Javascript:window.close()">Salir</a>
	</td>
</tr>
</table>
<iframe name="valida" style="visibility=hidden;" src="" width="100%" height="100%"></iframe> 
</form>
<%
set l_rs = nothing
Cn.Close
'Cn = nothing
%>
</body>
</html>
