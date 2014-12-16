<%Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : conf_x_empr_02.asp
Descripcion    : Modulo que se encarga de mostrar la config de una empresa.
Creador        : Scarpa D.
Fecha Creacion : 21/08/2003
Modificacion   :
   01/10/2003 - Scarpa D. - Modificaciones en la estructura HTML
-----------------------------------------------------------------------------
-->
<% 
' Variables
Dim l_confnro

Dim	l_confdesc
Dim	l_confint 
Dim	l_confactivo 

Dim	l_confdescant
Dim	l_confintant 
Dim	l_confactivoant

dim l_tipo

dim l_rs
dim l_rs1
dim l_sql

l_tipo        = Request.QueryString("tipo")
l_confnro      = Request.QueryString("confnro")

select Case l_tipo
	Case "A":

		l_confdesc      = ""
		l_confint       = ""
		l_confactivo    = ""
		
		l_confdescant   = ""
		l_confintant    = ""
		l_confactivoant = ""

	Case "M":

        Set l_rs = Server.CreateObject("ADODB.RecordSet")	
		
		l_sql = "SELECT * "
		l_sql = l_sql & " FROM  confper"
		l_sql = l_sql & " WHERE confnro = " & l_confnro
		
		l_rs.MaxRecords = 1
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_confdesc      = l_rs("confdesc")
			l_confint       = l_rs("confint")
			l_confactivo    = l_rs("confactivo")
			
			l_confdescant   = l_rs("confdesc")
			l_confintant    = l_rs("confint")
			l_confactivoant = l_rs("confactivo")
		end if
		l_rs.Close
		set l_rs = nothing
end select
%>

<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Configuraci&oacute;n del reporte</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_hora.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script>
function Validar_Formulario()
{

if (document.datos.confdesc.value == "" )
	alert("Ingrese una descripción.");
else	
if (document.datos.confint.value == "" )
	alert("Ingrese un valor.");
else	
if (isNaN(document.datos.confint.value))
	alert("Ingrese un valor numerico.");
else{	
    abrirVentanaH('','vent_oculta',200,200); 
	document.datos.submit();
}

}

</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">

<form name="datos" action="conf_x_empr_03.asp?Tipo=<%=l_tipo%>"  target="vent_oculta" method="post">
<input type="hidden" name="tipo" value="<%=l_tipo%>">

<input type="hidden" name="confintant" value="<%=l_confintant%>">
<input type="hidden" name="confdescant" value="<%=l_confdescant%>">
<input type="hidden" name="confactivoant" value="<%=l_confactivoant%>">
<table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
<tr style="border-color :CadetBlue;">
<td colspan="2" align="left" class="barra">Datos de la Configuraci&oacute;n</td>
<td colspan="2" align="right" class="barra">
    <a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
</td>	
</tr>
<tr>
	<td align="right"><b>C&oacute;digo:</b></td>
	<td colspan=3><input type="text" name="confnro" size="4" maxlength="4" value="<%=l_confnro%>" class="deshabinp" readonly="true">
	</td>
</tr>
<tr>
	<td align="right"><b>Descripci&oacute;n:</b></td>
	<td colspan=3><input type="text" name="confdesc" size="30" maxlength="30" value="<%=l_confdesc%>">
	</td>
</tr>
<tr>
	<td align="right"><b>Valor:</b></td>
	<td colspan=3><input type="text" name="confint" size="16" maxlength="16" value="<%=l_confint%>">
	</td>
</tr>
<tr>
	<td align="right"><b>Activo:</b></td>
	<td colspan=3>
	  <input type="checkbox" name="confactivo" <% if CInt(l_confactivo) = -1 then response.write "checked" end if %>>
	</td>
</tr>
<tr>
    <td align="right" class="th2" colspan=4>
		<a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>
</table>
</form>

</body>
</html>
