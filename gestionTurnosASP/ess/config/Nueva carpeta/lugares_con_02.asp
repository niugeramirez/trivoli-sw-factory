<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: lugares_con_02.asp
'Descripción: Consulta de lugares
'Autor : Rayl Chinestra
'Fecha: 08/02/2005
'Modificado : Raul Chinestra 02/03/2006 Se eliminaron los campos lugpro, lugbaj y se agregó el campo lugzon que indica la 
' zona comercial a la que pertenece el lugar y que se va a usar para bajar los cupos, contratos y ordenes de trabajo.

'Datos del formulario
'lugnro lugcod lugdes planro locnro pronro lugpro
on error goto 0
Dim l_lugnro
Dim l_lugcod
Dim l_lugdes
Dim l_plades
Dim l_locdes
Dim l_prodes
Dim l_lugzon
Dim l_lugdir
Dim l_lugestacion
Dim l_lugdesvio

'ADO
Dim l_sql
Dim l_rs

%>
<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Lugares - Ticket</title>
</head>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script>

function Validar_Formulario(){
	document.datos.submit();
}

</script>
<% 
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_lugnro = request.querystring("cabnro")
l_sql = "SELECT  lugnro, lugcod, lugdes, locdes, prodes, lugzon, lugdir, estacion, desvio "
l_sql = l_sql & " FROM tkt_lugar "
l_sql = l_sql & " INNER JOIN tkt_localidad ON tkt_localidad.locnro = tkt_lugar.locnro "
l_sql = l_sql & " INNER JOIN tkt_provincia ON tkt_provincia.pronro = tkt_lugar.pronro "
l_sql = l_sql  & " WHERE lugnro = " & l_lugnro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_lugnro = l_rs("lugnro")
	l_lugcod = l_rs("lugcod")
	l_lugdes = l_rs("lugdes")
	l_locdes = l_rs("locdes")
	l_prodes = l_rs("prodes")
	l_lugzon = l_rs("lugzon")
	l_lugdir = l_rs("lugdir")
	l_lugestacion = l_rs("estacion")
	l_lugdesvio = l_rs("desvio")
end if

l_rs.Close
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos" action="lugares_con_03.asp" method="post" target="valida">
<input type="hidden" name="lugnro" value="<%= l_lugnro %>">
<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
	<tr>
	    <td class="th2"  nowrap>Lugares</td>
		<td class="th2" align="right">
			<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
		</td>
	</tr>
	<tr>
		<td height="100%" colspan="2">
			<table>
				<tr>
					<td width="50%"></td>
					<td>
						<table>
							<tr>
							    <td height="100%" align="right"><b>Código:</b></td>
								<td height="100%">
									<input type="text" readonly class="deshabinp" name="loccod" size="10" maxlength="10" value="<%= l_lugcod %>">
								</td>
							</tr>
							<tr>
							    <td height="100%" align="right"><b>Descripción:</b></td>
								<td height="100%">
									<input type="text" readonly class="deshabinp" name="prodes" size="60" maxlength="50" value="<%= l_lugdes %>">
								</td>
							</tr>
							<tr>
						    	<td align="right" nowrap><b>Dirección:</b></td>
								<td>
									<input type="text" name="lugdir" size="60" maxlength="50" value="<%= l_lugdir %>">
								</td>
							</tr>
							<tr>
							    <td height="100%" align="right"><b>Localidad:</b></td>
								<td height="100%">
									<input type="text" readonly class="deshabinp" name="locdes" size="60" maxlength="50" value="<%= l_locdes %>">
								</td>
							</tr>
							<tr>
							    <td height="100%" align="right"><b>Provincia:</b></td>
								<td height="100%">
									<input type="text" readonly class="deshabinp" name="prodes" size="60" maxlength="50" value="<%= l_prodes %>">
								</td>
							</tr>
							<tr>
							    <td height="100%" align="right"><b>Zona:</b></td>
								<td height="100%">
									<input type="text" class="deshabinp" name="lugzon" size="10" maxlength="10" value="<%= l_lugzon %>">
								</td>
							</tr>

							<tr>
							    <td height="100%" align="right"><b>Estacion:</b></td>
								<td height="100%">
									<input type="text" name="lugestacion" size="60" maxlength="50" value="<%= l_lugestacion %>">
								</td>
							</tr>

							<tr>
							    <td height="100%" align="right"><b>Desvio:</b></td>
								<td height="100%">
									<input type="text" name="lugdesvio" size="18" maxlength="15" value="<%= l_lugdesvio %>">
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
  		   <a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>
    		<% call MostrarBoton ("sidebtnABM", "Javascript:window.close();","Salir")%>			
		</td>
	</tr>
</table>
<iframe name="valida" style="visibility=hidden;" src="" width="100%" height="100%"></iframe> 
</form>
<%
set l_rs = nothing
cn.Close
set cn = nothing
%>
</body>
</html>
