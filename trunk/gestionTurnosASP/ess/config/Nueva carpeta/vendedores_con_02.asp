<% Option Explicit %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<% 
'Archivo: vendedores_con_02.asp
'Descripción: Abm de Vendedores
'Autor : Gustavo Manfrin
'Fecha: 08/02/2005
'Modificado: 04/01/2006 Raul Chinestra Se agrego el combo de Camaras Arbitrales

'Datos del formulario
Dim l_vencornro
Dim l_vencorcod
Dim l_vencordes
Dim l_vencorrazsoc
Dim l_venhab
Dim l_vendom
Dim l_venloc
Dim l_camnro

'ADO
Dim l_sql
Dim l_rs

%>
<html>
<head>
<link href="/ticket/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Vendedores - Ticket</title>
</head>
<script src="/ticket/shared/js/fn_ayuda.js"></script>
<script src="/ticket/shared/js/fn_windows.js"></script>
<script>

function validar(){
	document.datos.submit();
}	

</script>
<% 
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_vencornro = request.querystring("cabnro")

l_sql = "SELECT vencornro, vencorcod, vencordes, vencorrazsoc, venhab, vencordom, vencorlocnro, locdes, camnro "
l_sql = l_sql & " FROM tkt_vencor "
l_sql = l_sql & " LEFT JOIN tkt_localidad ON tkt_localidad.locnro = tkt_vencor.vencorlocnro "
l_sql = l_sql & " WHERE venact = -1 AND vencortip = 'V' "
l_sql = l_sql & " AND vencornro = " & l_vencornro

rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_vencornro = l_rs("vencornro")
	l_vencorcod = l_rs("vencorcod")
	l_vencordes = l_rs("vencordes")
	l_vencorrazsoc = l_rs("vencorrazsoc")
	l_venhab = l_rs("venhab")
	l_vendom = l_rs("vencordom")
	l_venloc = l_rs("locdes")
	l_camnro = l_rs("camnro")
end if
l_rs.Close
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos" action="vendedores_con_03.asp" method="post" target="ifrm">
	<input type="hidden" name="vencornro" value="<%= l_vencornro %>">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2"  nowrap>Vendedores</td>
	<td class="th2" align="right">
		<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</td>
</tr>
<tr>
	<td colspan="2" width="100%" height="100%">	
		<table cellpadding="0" cellspacing="0">
			<tr>
				<td width="50%"></td>
				<td>
					<table cellpadding="0" cellspacing="0">
						<tr>
						    <td align="right"><b>Código:</b></td>
							<td>
								<input type="text" readonly class="deshabinp" name="vencod" size="20" maxlength="15" value="<%= l_vencorcod %>">
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Nombre clave:</b></td>
							<td>
								<input type="text" readonly class="deshabinp" name="vendes" size="60" maxlength="50" value="<%= l_vencordes %>">
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Razón social:</b></td>
							<td>
								<input type="text" readonly class="deshabinp" name="venrzso" size="60" maxlength="50" value="<%= l_vencorrazsoc %>">
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Domicilio:</b></td>
							<td>
								<input type="text" readonly class="deshabinp" name="vendom" size="60" maxlength="50" value="<%= l_vendom %>">
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Localidad:</b></td>
							<td>
								<input type="text" readonly class="deshabinp" name="locdes" size="60" maxlength="50" value="<%= l_venloc %>">
							</td>
						</tr>												
						<tr>
						    <td align="right" nowrap><b>Habilitado:</b></td>
							<td>
								<input type="Checkbox" readonly disabled name="venhab" <% If l_rs("venhab") = -1 then %>Checked<% end if %>>				
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Cámara Arbitral:</b></td>
							<td colspan="3">
								<select name="camnro" size="1" style="width:300;" <%'= l_claseCombo %>>
									<option value=0 selected>&laquo; Cámara definida en Configuración &raquo;</option>
								<%	l_sql = "SELECT camnro, camdes, camcod "
									l_sql  = l_sql  & " FROM tkt_camara "
									l_sql  = l_sql  & " ORDER BY camdes "
									rsOpen l_rs, cn, l_sql, 0
									do until l_rs.eof %>	
									<option value=<%= l_rs("camnro") %> > 
									<%= l_rs("camdes") %> (<%=l_rs("camcod")%>) </option>
									<%	l_rs.Movenext
									loop
									l_rs.Close %>
								</select>
								<% If l_camnro = "0" or l_camnro = "" or IsNull(l_camnro) then
									l_camnro = 0
								end if %>
									<script> document.datos.camnro.value= "<%= l_camnro %>"</script>
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
		<iframe name="ifrm" style="visibility=hidden;" src="" width="0" height="0"></iframe> 	
		<a class=sidebtnABM href="Javascript:validar()">Aceptar</a>	
		<a class=sidebtnABM href="Javascript:window.close()">Salir</a>
	</td>
</tr>
</table>
</form>
<%
set l_rs = nothing
Cn.Close
Cn = nothing
%>
</body>
</html>
