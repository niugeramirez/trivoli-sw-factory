<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
Archivo     : rep_recibo_ess_00.asp
Autor       : GdeCos
Creacion    : 14/04/2005
Descripcion : Modulo que se encarga de generar el listado en el IFRAME del 00,
              recibos de sueldos.
Modificacion: 14/08/2006 - Martin Ferraro - Paso como parametro el asp del recibo de sueldo a mostrar
											Correccion de funcion atras que perdia el legajo de mss
											ahora el recibo se toma de la carpeta \rhpro\liq
-->
<%
on error goto 0

Dim l_rs
Dim l_sql
Dim l_recibo
Dim l_parametro
Dim l_salida
Dim leg_ant

l_recibo = request("rec")

if l_recibo = "" then
   l_recibo = 0
end if

l_parametro = request.querystring("salida")
if trim(l_parametro) = "" then
	l_salida = "rep_recibo_liq_03.asp"
else
	l_salida = l_parametro
end if

%>
<html>
<head>
<link href="../<%= c_estilo %>" rel="StyleSheet" type="text/css">
<title>Recibo de Sueldos - Liquidación de Haberes - Autogestion</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_ay_generica.js"></script>
<SCRIPT SRC="/serviciolocal/shared/js/menu_def.js"></SCRIPT>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script src="/serviciolocal/shared/js/fn_help_emp.js"></script>
<script src="/serviciolocal/shared/js/fn_buscar_emp.js"></script>
<script>
var indice = 1;
var ifrmListo = 0;

function Imprimir(){
	document.ifrm.focus();
	window.print();
}


function ConfPagina(){
    var WebBrowser = '<OBJECT ID="WebBrowser' + indice + '" WIDTH=0 HEIGHT=0 CLASSID="CLSID:8856F961-340A-11D0-A96B-00C04FD705A2"></OBJECT>';
    document.body.insertAdjacentHTML('beforeEnd', WebBrowser);

    execScript("on error resume next: WebBrowser" + indice + ".ExecWB 8, -1", "VBScript");
	indice++;
}

function Atras()
{
	<% 
	if Session("empleg") = l_ess_empleg then
		leg_ant = ""
	else
		leg_ant = l_ess_empleg
	end if
	%>
	<% if l_parametro = "" then %>
	window.location = "recibosueldo_ess_00.asp?empleg=<%= leg_ant %>";
	<% else %>
	window.location = "recibosueldo_ess_00.asp?salida=<%= l_salida %>?empleg=<%= leg_ant %>";
	<% end if %>

}

</script>

</head>

<body>
<form name=datos>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" align="center">
	<tr>
	    <th height="10">
			Recibos de Sueldos
		</th>
		<th align="center">		  
		    <a class="sidebtnSHW" href="Javascript:Atras()">Atrás</a>			  			
	    	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
       	    <a class="sidebtnSHW" href="Javascript:Imprimir()">Imprimir</a>			  
		</th>
	</tr>	
    <tr valign="top" height="100%">
        <td colspan="2" align="center" style="">
			  <iframe name="ifrm" src="..\..\liq\<%= l_salida %>?bpronro=<%= l_recibo %>&empleg=<%= l_ess_empleg%>" width="100%" height="100%"></iframe> 
	    </td>
    </tr>
</table>
</form>

</body>
</html>
