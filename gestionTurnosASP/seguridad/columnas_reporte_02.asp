<%Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->

<% 

'Archivo: columnas_reporte_02.asp
'Descripción: Módulo que se encarga de mostrar los datos
'             de una columna del confrep.
'Modificado:
'    29/07/2003 - Scarpa D. - Agregado de la columna confsuma   
'    25/08/2003 - Scarpa D. - Agregado de la columna confval2		
'    01/10/2003 - Scarpa D. - Cambio en etiquetas
'    14/04/2004 - Alvaro Bayon - Al campo valor le di 9 lugares
'    23/11/2004 - Alvaro Bayon - Validaciones

' Variables
Dim l_repnro
dim l_repdesc
Dim l_confnrocol
Dim l_confetiq
Dim l_conftipo
Dim	l_confval
Dim	l_confval2
Dim	l_confaccion

Dim l_confnrocolant
Dim l_confetiqant
Dim l_conftipoant
Dim	l_confvalant
Dim	l_confval2ant
Dim l_confaccionant

dim l_tipo

dim l_rs
dim l_rs1
dim l_sql

l_tipo        = Request.QueryString("tipo")
l_repnro      = Request.QueryString("repnro")
l_confnrocol  = Request.QueryString("confnrocol")
l_conftipo    = Request.QueryString("conftipo")
l_confetiq    = Request.QueryString("confetiq")
l_confval     = Request.QueryString("confval")
l_confval2    = Request.QueryString("confval2")
l_confaccion  = Request.QueryString("confaccion")

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT reporte.repnro, reporte.repdesc "
l_sql = l_sql & " FROM  reporte "
l_sql = l_sql & " WHERE reporte.repnro =  " & l_repnro
l_rs.Maxrecords = 50
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_repdesc = l_rs("repdesc")
end if

select Case l_tipo
	Case "A":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT conftipo, "  
		l_sql = l_sql & " confetiq,  "
		l_sql = l_sql & " confval, confnrocol, confval2, confaccion "
		l_sql = l_sql & " FROM  confrep"
		l_sql = l_sql & " WHERE confrep.repnro = " & l_repnro
		l_sql = l_sql & " ORDER BY confrep.confnrocol DESC "
		
		l_rs.MaxRecords = 1
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_confnrocol = l_rs("confnrocol") + 1
		else
			l_confnrocol =  1
		end if
		l_rs.Close
		set l_rs = nothing
		l_confetiq = ""
		l_conftipo = ""
		l_confval  = 0

	Case "M":
		If len(trim(l_confnrocol)) = 0 then
			response.write("<script>alert('Debe seleccionar una columna');window.close();</script>")
		end if

		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT * FROM  confrep"
		l_sql = l_sql & " WHERE confrep.repnro = " & l_repnro
		l_sql = l_sql & " AND   confrep.confnrocol = " & l_confnrocol
		l_sql = l_sql & " AND   confrep.conftipo = '" & l_conftipo & "'"
		l_sql = l_sql & " AND   confrep.confetiq = '" & l_confetiq & "'"
		l_sql = l_sql & " AND   confrep.confval = " & l_confval
		l_sql = l_sql & " AND   confrep.confval2 = '" & l_confval2 & "'"
		l_sql = l_sql & " AND   confrep.confaccion = '" & l_confaccion & "'"		
		
		l_rs.MaxRecords = 1
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_confnrocol = l_rs("confnrocol")
			l_conftipo = l_rs("conftipo")
			l_confetiq = l_rs("confetiq")
			l_confval  = l_rs("confval")
			l_confval2  = l_rs("confval2")			
			l_confaccion  = l_rs("confaccion")			
			
			l_confnrocolant = l_rs("confnrocol")
			l_conftipoant = l_rs("conftipo")
			l_confetiqant = l_rs("confetiq")
			l_confvalant  = l_rs("confval")
			l_confval2ant  = l_rs("confval2")			
			l_confaccionant  = l_rs("confaccion")						
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
<script src="/serviciolocal/shared/js/fn_valida.js"></script>
<script>
function Validar_Formulario()
{
if (document.datos.confnrocol.value == "" ){
	alert("Ingrese un Nro de Columna.");
	document.datos.confnrocol.focus();
	}
else
if (isNaN(document.datos.confnrocol.value)){
	alert("Ingrese un número de Columna.");
	document.datos.confnrocol.focus();
	}	
else
if (Trim(document.datos.confetiq.value) == "" ){
	alert("Ingrese una Etiqueta.");
	document.datos.confetiq.focus();
	}
else	
if (!stringValido(document.datos.confetiq.value)){
	alert("La Etiqueta contiene caracteres no válidos.");
	document.datos.confetiq.focus();
	}
else	
if (Trim(document.datos.conftipo.value) == "" ){
	alert("Ingrese un Tipo.");
	document.datos.conftipo.focus();
	}
else	
if (!stringValido(document.datos.conftipo.value)){
	alert("El Tipo contiene caracteres no válidos.");
	document.datos.conftipo.focus();
	}
else	
if ((document.datos.confval.value == "" ) || (document.datos.confval.value == 0 )){
	alert("Ingrese un Valor.");
	document.datos.confval.focus();
	}
else	
if (isNaN(document.datos.confval.value)){
	alert("Ingrese un Valor numérico.");
	document.datos.confval.focus();
	}	
else
if (!stringValido(document.datos.confval2.value)){
	alert("El Valor Alfanumérico contiene caracteres no válidos.");
	document.datos.confval2.focus();
	}
else	

	document.datos.submit();
}

</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onload="javascript:document.datos.confnrocol.focus();">
<form name="datos" action="columnas_reporte_03.asp?Tipo=<%=l_tipo%>" method="post">
<input type="hidden" name="tipo" value="<%=l_tipo%>">

<input type="hidden" name="confnrocolant" value="<%=l_confnrocolant%>">
<input type="hidden" name="conftipoant" value="<%=l_conftipoant%>">
<input type="hidden" name="confetiqant" value="<%=l_confetiqant%>">
<input type="hidden" name="confvalant" value="<%=l_confvalant%>">
<input type="hidden" name="confval2ant" value="<%=l_confval2ant%>">
<input type="hidden" name="confaccionant" value="<%=l_confaccionant%>">
<table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
<tr style="border-color :CadetBlue;">
<td colspan="2" align="left" class="barra">Datos de la Configuraci&oacute;n</td>
<td colspan="2" class="th2" align="right">
	  <a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
</td>
</tr>
<tr>
	<td align="right"><b>Reporte:</b></td>
	<td colspan=3>
	<input type="text" class="deshabinp" name="repnro" size="4" maxlength="4" value="<%=l_repnro%>" readonly>
	<input type="text" class="deshabinp" name="repdesc" size="30" maxlength="30" value="<%=l_repdesc%>" readonly>
	</td>
</tr>
<tr>
	<td align="right"><b>Nro Columna:</b></td>
	<td colspan=3><input type="text" name="confnrocol" size="4" maxlength="4" value="<%=l_confnrocol%>">
	</td>
</tr>
<tr>
	<td align="right"><b>Etiqueta:</b></td>
	<td colspan=3><input type="text" name="confetiq" size="50" maxlength="50" value="<%=l_confetiq%>">
	</td>
</tr>

<tr>
	<td align="right"><b>Tipo:</b></td>
	<td colspan=3><input type="text" name="conftipo" size="3" maxlength="3" value="<%=l_conftipo%>">
	</td>
</tr>

<tr>
	<td align="right"><b>Valor Num&eacute;rico:</b></td>
	<td colspan=3><input type="text" name="confval" size="10" maxlength="9" value="<%=l_confval%>">
	</td>
</tr>
<tr>
	<td align="right"><b>Valor Alfanum&eacute;rico:</b></td>
	<td colspan=3><input type="text" name="confval2" size="10" maxlength="10" value="<%=l_confval2%>">
	</td>
</tr>
<tr>
	<td align="right"><b>Acci&oacute;n:</b></td>
	<td colspan=3>
	   <select name="confaccion" size="1">
 	     <option value="sumar" <%if l_confaccion = "sumar" then response.write "selected" end if%> > Sumar	   
 	     <option value="restar" <%if l_confaccion = "restar" then response.write "selected" end if%> > Restar
	   </select> 
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
