<%Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<%
'Archivo        : conf_mil_02.asp
'Descripcion    : Modulo que se encarga de admin. los servidores de mail
'Creador        : Lisandro Moro
'Fecha Creacion : 08/03/2005
'Modificacion   :

on error goto 0

' Variables
Dim l_cfgemailnro
Dim l_cfgemailhost
Dim l_cfgemaildesc
Dim l_cfgemailfrom
Dim l_cfgemailest
Dim l_cfgemailport

Dim l_cfgemailhostant
Dim l_cfgemaildescant
Dim l_cfgemailfromant
Dim l_cfgemailestant
Dim l_cfgemailportant

dim l_tipo

dim l_rs
dim l_rs1
dim l_sql

l_tipo        = Request.QueryString("tipo")
l_cfgemailnro = Request.QueryString("cfgemailnro")

select Case l_tipo
	Case "A":

        l_cfgemailhost = ""
        l_cfgemaildesc = ""
        l_cfgemailfrom = ""
        l_cfgemailest  = "0"
        l_cfgemailport = "0"

        l_cfgemailhostant = ""
        l_cfgemaildescant = ""
        l_cfgemailfromant = ""
        l_cfgemailestant  = "0"
        l_cfgemailportant = "0"

	Case "M":

        Set l_rs = Server.CreateObject("ADODB.RecordSet")	
		
		l_sql = "SELECT * "
		l_sql = l_sql & " FROM  conf_email"
		l_sql = l_sql & " WHERE cfgemailnro = " & l_cfgemailnro
		
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_cfgemailhost = l_rs("cfgemailhost")
			l_cfgemaildesc = l_rs("cfgemaildesc")
			l_cfgemailfrom = l_rs("cfgemailfrom")
			l_cfgemailest  = l_rs("cfgemailest")
			l_cfgemailport = l_rs("cfgemailport")
			
			l_cfgemailhostant = l_rs("cfgemailhost")
			l_cfgemaildescant = l_rs("cfgemaildesc")
			l_cfgemailfromant = l_rs("cfgemailfrom")
			l_cfgemailestant  = l_rs("cfgemailest")
			l_cfgemailportant = l_rs("cfgemailport")
		end if 
		l_rs.Close
		set l_rs = nothing
end select
%>

<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Configuraci&oacute;n Servicio de Email</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_hora.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script>

function Validar_Formulario(){

	if (document.datos.cfgemaildesc.value == "" )
		alert("Debe ingresar una descripción.");
	else
	if (document.datos.cfgemailhost.value == "" )
		alert("Debe ingresar un nombre de host.");
	else
	if (document.datos.cfgemailport.value == "" )
		alert("Debe ingresar un número de puerto.");
	else	
	if (isNaN(document.datos.cfgemailport.value))
		alert("El puerto debe ser númerico.");
	else	
	if (document.datos.cfgemailfrom.value == "" )	
		alert("Debe ingresar un email de origen.");	
	else{	
	    abrirVentanaH('','vent_oculta',200,200); 
		document.datos.submit();
	}

}

</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">

<form name="datos" action="conf_mail_03.asp?Tipo=<%=l_tipo%>"  target="vent_oculta" method="post">
<input type="hidden" name="tipo" value="<%=l_tipo%>">

<input type="hidden" name="cfgemailportant" value="<%=l_cfgemailportant%>">
<input type="hidden" name="cfgemailestant"  value="<%=l_cfgemailestant%>">
<input type="hidden" name="cfgemailfromant" value="<%=l_cfgemailfromant%>">
<input type="hidden" name="cfgemaildescant" value="<%=l_cfgemaildescant%>">
<input type="hidden" name="cfgemailhostant" value="<%=l_cfgemailhostant%>">

<table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
<tr style="border-color :CadetBlue;">
<td colspan="2" align="left" class="barra">Datos de la Configuraci&oacute;n</td>
<td colspan="2" align="right" class="barra">
    <a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
</td>	
</tr>
<tr>
	<td align="right"><b>C&oacute;digo:</b></td>
	<td colspan=3><input type="text" name="cfgemailnro" size="10" maxlength="10" value="<%=l_cfgemailnro%>" class="deshabinp" readonly="true">
	</td>
</tr>
<tr>
	<td align="right"><b>Descripci&oacute;n:</b></td>
	<td colspan=3><input type="text" name="cfgemaildesc" size="30" maxlength="100" value="<%=l_cfgemaildesc%>">
	</td>
</tr>
<tr>
	<td align="right"><b>EMail Origen:</b></td>
	<td colspan=3><input type="text" name="cfgemailfrom" size="25" maxlength="50" value="<%=l_cfgemailfrom%>">
	</td>
</tr>
<tr>
	<td align="right"><b>Host:</b></td>
	<td colspan=3><input type="text" name="cfgemailhost" size="25" maxlength="50" value="<%=l_cfgemailhost%>">
	</td>
</tr>
<tr>
	<td align="right"><b>Puerto:</b></td>
	<td colspan=3><input type="text" name="cfgemailport" size="6" maxlength="6" value="<%=l_cfgemailport%>">
	</td>
</tr>
<tr>
	<td align="right"><b>Activo:</b></td>
	<td colspan=3><input type="checkbox" name="cfgemailest" <%if CInt(l_cfgemailest) = -1 then response.write "checked" end if%>>
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
