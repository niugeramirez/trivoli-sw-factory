<% Option Explicit %>
<!--#include virtual="/ticket/shared/inc/sec.inc"-->
<!--#include virtual="/ticket/shared/inc/const.inc"-->
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<% 
'Archivo: param_configuracion_con_00.asp
'Descripción: Abm de parametros de configuracion
'Autor : Lisandro Moro
'Fecha: 01/03/2005
%>
<html>
<head>
<link href="/ticket/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<title><%= Session("Titulo")%>Parámetros de Configuración - Ticket</title>
<script src="/ticket/shared/js/fn_windows.js"></script>
<script src="/ticket/shared/js/fn_confirm.js"></script>
<script src="/ticket/shared/js/fn_ayuda.js"></script>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
	<table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%">
		<tr style="border-color :CadetBlue;">
			<td align="left" class="barra">Parámetros de Configuración</td>
			<td nowrap align="right" class="barra">
				<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
			</td>
		</tr>
		<tr valign="top" height="100%">
			<td colspan="2" style="" width="100%">
				<iframe name="ifrm" src="param_configuracion_con_01.asp" frameborder="0" width="100%" height="100%" scrolling="no" ></iframe> 
			</td>
		</tr>
		<tr>
		    <td colspan="2" align="right" class="th2">
				<% call MostrarBoton ("sidebtnABM", "Javascript:document.ifrm.Valida();","Aceptar")%>
				<a class=sidebtnABM href="Javascript:window.parent.close()">Salir</a>
				
			</td>
		</tr>
	</table>
</body>
</html>
