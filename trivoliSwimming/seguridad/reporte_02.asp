<%Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/inc/sec.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/const.inc"-->
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<% 

'Modificado
'24-11-04 - Alvaro Bayon - Validaciones - Botón de ayuda, foco.

' Variables
Dim l_repnro
Dim l_repdesc
dim	l_repagr

dim l_tipo

dim l_rs
dim l_rs1
dim l_sql

l_tipo      = Request.QueryString("tipo")
l_repnro   = Request.QueryString("repnro")

select Case l_tipo
	Case "A":
		 l_repdesc = ""
		 l_repagr = 0

	Case "M":
		If len(trim(l_repnro)) = 0 then
			response.write("<script>alert('Debe seleccionar un reporte');window.close();</script>")
		end if

		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT repnro, "  
		l_sql = l_sql & " repdesc,  "
		l_sql = l_sql & " repagr   "
		l_sql = l_sql & " FROM  reporte"
		l_sql = l_sql & " WHERE reporte.repnro = " & l_repnro

		l_rs.MaxRecords = 1
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_repdesc = l_rs("repdesc")
			l_repagr = l_rs("repagr")
		end if
		l_rs.Close
		set l_rs = nothing
end select
%>

<html>
<head>
<link href="/trivoliSwimming/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Reportes</title>
<script src="/trivoliSwimming/shared/js/fn_windows.js"></script>
<script src="/trivoliSwimming/shared/js/fn_confirm.js"></script>
<script src="/trivoliSwimming/shared/js/fn_hora.js"></script>
<script src="/trivoliSwimming/shared/js/fn_valida.js"></script>
<script src="/trivoliSwimming/shared/js/fn_ayuda.js"></script>
<script>
function Validar_Formulario()
{
if (Trim(document.datos.repdesc.value)==""){
	alert("Ingrese una descripción.")
	document.datos.repdesc.focus()
	}
else	
if (!stringValido(document.datos.repdesc.value)){
	alert("La descripción contiene caracteres no válidos.");
	document.datos.repdesc.focus()
	}
else	
		document.datos.submit();
}

</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onload="document.datos.repdesc.focus()">

<table border="0" cellpadding="0" cellspacing="0" height="100%">
<tr style="border-color :CadetBlue;">
	<td class="th2" align="left" class="barra">Datos del Reporte</td>
	<td class="th2" align="right">		  
		<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</td>

</tr>

<form name="datos" action="reporte_03.asp?Tipo=<%=l_tipo%>" method="post">
<input type="hidden" name="tipo" value="<%=l_tipo%>">
<%if l_tipo = "M" then%> 
<tr>
	<td align="right"><b>C&oacute;digo:</b></td>
	<td ><input type="text" style="deshabinp" name="repnro" size="10" maxlength="10" value="<%=l_repnro%>" readonly>
	</td>
</tr>
<%end if%>
<tr>
	<td align="right"><b>Descripci&oacute;n:</b></td>
	<td ><input type="text" name="repdesc" size="30" maxlength="30" value="<%=l_repdesc%>">
	</td>
</tr>
<tr>
	<td align="right">
	<%if len(trim(l_repagr)) <> 0 then
		if l_repagr then %>
			<input type="checkbox" checked id=checkbox1 name=repagr>
		<%else%>
			<input type="checkbox" id=checkbox1 name=repagr >	
		<%end if
	else%>
		<input type="checkbox" id=checkbox1 name=repagr >
	<%end if%>	
	</td>
	<td align="left" ><b>Agrupado</b></td>
</tr>


</form>
<tr>
    <td colspan=2 align="right" class="th2">
		<a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>
</table>

</body>
</html>
