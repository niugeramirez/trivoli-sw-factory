<%Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<% on error goto 0
'---------------------------------------------------------------------------------
'Archivo	: nota_adp_02.asp
'Descripción: Form de datos de nota
'Autor		: Claudia Cecilia Rossi
'Fecha		: 30-08-2003
'Modificado	: 08-11-05 - Leticia A. - Adecuarlo para Autogestion.
'----------------------------------------------------------------------------------

'Variables
 Dim l_empleg
 Dim l_notanro
 Dim l_tnonro
 Dim l_tnonroant
 Dim l_notatxt
 Dim l_notmotivo
 Dim l_notremitente
 Dim l_nothoravenc_m
 Dim l_nothoravenc_h
 Dim l_notfecvenc
 Dim l_nothoraalta_m
 Dim l_nothoraalta_h
 Dim l_notfecalta 
 'Dim l_ternro
 Dim l_tipo
 
 Dim l_rs
 Dim l_rs1
 Dim l_sql
 
' tomar parametros de entrada
 l_tipo   = Request.QueryString("tipo")
 l_notanro= Request.QueryString("notanro")
 
 l_empleg	= request.Querystring("empleg")   ' l_ess_empleg - NO 
 'l_ternro = l_ess_ternro
 
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 
 select Case l_tipo
	Case "A":
		 l_tnonro 		 = ""
		 l_notatxt 		 = ""
		 l_notfecalta	 = date()
		 l_nothoraalta_h = mid(time,1,2)
		 l_nothoraalta_m = mid(time,4,2)
		 l_notfecvenc	 = ""
		 l_nothoravenc_h = ""
		 l_nothoravenc_m = ""
		 l_notremitente	 = ""
		 l_notmotivo	 = ""
	Case "M":
		l_sql = "SELECT notatxt, ternro, notfecalta, nothoraalta, notfecvenc, nothoravenc, notremitente, notmotivo, tnonro "
		l_sql = l_sql & " FROM  notas_ter "
		l_sql = l_sql & " WHERE notanro = " & l_notanro
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			 l_tnonro		 = l_rs("tnonro")
			 'l_tnonroant     = l_rs("tnonroant")
			 l_notatxt		 = l_rs("notatxt")
			 l_notfecalta	 = l_rs("notfecalta")
			 l_nothoraalta_h = left(l_rs("nothoraalta"), 2)
			 l_nothoraalta_m = right(l_rs("nothoraalta"), 2)
			 l_notfecvenc	 = l_rs("notfecvenc")
			 l_nothoravenc_h = left(l_rs("nothoravenc"), 2)
			 l_nothoravenc_m = right(l_rs("nothoravenc"), 2)
			 l_notremitente	 = l_rs("notremitente")
			 l_notmotivo	 = l_rs("notmotivo")
		end if
		l_rs.Close
 end select
%>

<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Notas  - Administraci&oacute;n de Personal - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_hora.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script src="/serviciolocal/shared/js/fn_valida.js"></script>
<script>
function Validar_Formulario(){
	if (document.datos.tnonro.value == "" ){
		alert("Seleccione un Tipo de Nota.");
		document.datos.tnonro.focus();
		return;
	}
	if (document.datos.notfecalta.value == "" ){
		alert("Debe ingresar una Fecha de Alta.");
		document.datos.notfecalta.focus();
		return;
	}
	if (!validarfecha(document.datos.notfecalta)){
		document.datos.notfecalta.focus();
		document.datos.notfecalta.select();
		return;
	}
	if (document.datos.nothoraalta_h.value == ""){
		alert('Debe ingresar la Hora de Ingreso.');
		document.datos.nothoraalta_h.focus();
		return;
	}
	if (!h_correcta(document.datos.nothoraalta_h.value, '00')){
		alert('La Hora de Ingreso no es válida.');
		document.datos.nothoraalta_h.focus();
		document.datos.nothoraalta_h.select();
		return;
	}
	if (document.datos.nothoraalta_m.value == ""){
		alert('Debe ingresar los minutos de Hora de Ingreso.');
		document.datos.nothoraalta_m.focus();
		return;
	}
	if (!h_correcta(document.datos.nothoraalta_h.value, document.datos.nothoraalta_m.value)){
		alert('La Hora de Ingreso no es válida.');
		document.datos.nothoraalta_h.focus();
		document.datos.nothoraalta_h.select();
		return;
	} 
	if (!h_correcta('00', document.datos.nothoraalta_m.value)){
		alert('Los minutos de la Hora de Ingreso no son válidos.');
		document.datos.nothoraalta_m.focus();
		document.datos.nothoraalta_m.select();
		return;
	} 
	if (document.datos.notfecvenc.value != "" && !validarfecha(document.datos.notfecvenc)){
		document.datos.notfecvenc.focus();
		document.datos.notfecvenc.select();
		return;
	}
	if (document.datos.notfecalta.value != "" && document.datos.notfecvenc.value != "" && menor(document.datos.notfecvenc.value,document.datos.notfecalta.value)){
		alert('La Fecha de Ingreso deber ser menor que la Fecha de Vencimiento.');
		document.datos.notfecvenc.focus();
		document.datos.notfecvenc.select();
		return;
	}
	if (document.datos.notfecvenc.value != "" && document.datos.nothoravenc_h.value == ""){
		alert('Si ingresa la Fecha de Vencimiento Debe ingresar la Hora de Vencimiento');
		document.datos.nothoravenc_h.focus();
		document.datos.nothoravenc_h.select();
		return;
	}	
	if (document.datos.notfecvenc.value != "" && document.datos.nothoravenc_h.value != "" && !h_correcta(document.datos.nothoravenc_h.value, '00')){
		alert('La Hora de Vencimiento no es válida.');
		document.datos.nothoravenc_h.focus();
		document.datos.nothoravenc_h.select();
		return;
	}
	if (document.datos.notfecvenc.value != "" && document.datos.nothoravenc_m.value != "" && !h_correcta('00', document.datos.nothoravenc_m.value)){
		alert('Los minutos de la Hora de Vencimiento no son válidos.');
		document.datos.nothoravenc_m.focus();
		document.datos.nothoravenc_m.select();
		return;
	}
	if (document.datos.notfecvenc.value != "" && document.datos.nothoravenc_m.value != "" && document.datos.nothoravenc_h.value == ""){
		alert('Debe Ingresar la Hora de Vencimiento.');
		document.datos.nothoravenc_h.focus();
		document.datos.nothoravenc_h.select();
		return;
	}
	if (document.datos.notfecvenc.value != "" && document.datos.nothoravenc_h.value != "" && document.datos.nothoravenc_m.value == ""){
		alert('Debe Ingresar los minutos de la Hora de Vencimiento.');
		document.datos.nothoravenc_m.focus();
		document.datos.nothoravenc_m.select();
		return;
	}
	if (document.datos.notfecvenc.value == "" && document.datos.nothoravenc_h.value != ""){
		alert('Si ingresa la Hora de Vencimiento Debe ingresar la Fecha de Vencimiento');
		document.datos.notfecvenc.focus();
		document.datos.notfecvenc.select();
		return;
	}	
	if ( document.datos.notfecvenc.value != "" && document.datos.nothoravenc_h.value != "" && (!h_correcta(document.datos.nothoravenc_h.value, document.datos.nothoravenc_m.value))){
		alert('La Hora de Vencimiento no es válida.');
		document.datos.nothoravenc_h.focus();
		document.datos.nothoravenc_h.select();
		return;
	} 
	if (!stringValido(document.datos.notremitente.value)) {
		alert("El Remitente contiene caracteres no válidos.");
		document.datos.notremitente.focus();
		document.datos.notremitente.select();
		return;
	}
	if (!stringValido(document.datos.notmotivo.value)) {
		alert("El Motivo contiene caracteres no válidos.");
		document.datos.notmotivo.focus();
		document.datos.notmotivo.select();
		return;
	}
	if (document.datos.notatxt.value.length >  1000 ){
		alert("La Nota no puede superar 1000 caracteres.");
		document.datos.notatxt.focus();
		document.datos.notatxt.select();
		return;
	}
/*	if (!stringValido(document.datos.notatxt.value)) {
		alert("La Nota contiene caracteres no válidos.");
		document.datos.notatxt.focus();
		document.datos.notatxt.select();
		return;
	}*/
	document.datos.submit(); 
}

function Nuevo_Dialogo(w_in, pagina, ancho, alto){
	return w_in.showModalDialog(pagina,'', 'help:0;status:0;resizable:0;center:1;scroll:0;yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}

function Ayuda_Fecha(txt){
	var jsFecha = Nuevo_Dialogo(window, '/serviciolocal/shared/js/calendar.html', 16, 14);
	if (jsFecha == null) txt.value = ''
	else txt.value = jsFecha;
}
</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onload="document.datos.tnonro.focus();">
<form name="datos" action="nota_adp_03.asp?empleg=<%=l_empleg%>" method="post">
<input type="hidden" name="tipo" value="<%=l_tipo%>">
<!-- <input type="hidden" name="ternro" value="<%'=l_ternro%>"> -->
<input type="hidden" name="notanro" value="<%=l_notanro%>">
<input type="hidden" name="tnonroant" value="<%=l_tnonroant%>">

<table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
 <tr style="border-color :CadetBlue;">
	<th align="left" class="th2" colspan="3" height="1">Datos de la Nota</th>
	<th class="th2" align="right">&nbsp;</th>
 </tr>
 <tr>
    <td align="right">&nbsp;<br><b>Tipo de Nota:</b></td>
	<td colspan="3">&nbsp;<br>
		<select name="tnonro" style="width:270px;">
		<option value="">&laquo;Seleccione una opci&oacute;n&raquo;</option>
		<%
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT tiponota.tnonro, tnodesabr "
		l_sql = l_sql & " FROM tiponota "
		l_sql = l_sql & " WHERE tiponota.tnoconfidencial = 0 "
		rsOpen l_rs, cn, l_sql, 0
		do while not l_rs.eof
			%><option value=<%=l_rs("tnonro")%>><%=l_rs("tnodesabr") & " (" & l_rs("tnonro") & ")"%></option><%
			l_rs.MoveNext
		loop
		l_rs.Close
		set l_rs = nothing
		%>
		</select>
		
		<script>document.datos.tnonro.value='<%=l_tnonro%>'</script>
	</td>
 </tr>
 <tr>
    <td align="right"><b>Fecha Ingreso:</b></td>
	<td>
		<input type="Text" name="notfecalta" size="10" maxlength="10" value="<%= l_notfecalta %>">
		<a href="Javascript:Ayuda_Fecha(document.datos.notfecalta);"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
	</td>
    <td align="right"><b>Hora Ingreso:</b></td>
	<td>
		<input type="Text" name="nothoraalta_h" size="2" maxlength="2" value="<%= l_nothoraalta_h %>">&nbsp;<b>:</b>&nbsp;
		<input type="Text" name="nothoraalta_m" size="2" maxlength="2" value="<%= l_nothoraalta_m %>">
	</td>
 </tr>
 <tr>
    <td align="right"><b>Fecha Vencimiento:</b></td>
	<td>
		<input type="Text" name="notfecvenc" size="10" maxlength="10" value="<%= l_notfecvenc %>">
		<a href="Javascript:Ayuda_Fecha(document.datos.notfecvenc);"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
	</td>
    <td align="right"><b>Hora Vencimiento:</b></td>
	<td>
		<input type="Text" name="nothoravenc_h" size="2" maxlength="2" value="<%= l_nothoravenc_h %>">&nbsp;<b>:</b>&nbsp;
		<input type="Text" name="nothoravenc_m" size="2" maxlength="2" value="<%= l_nothoravenc_m %>">
	</td>
 </tr>
 <tr>
    <td align="right"><b>Remitente:</b></td>
	<td colspan="3">
		<input type="Text" name="notremitente" size="50" maxlength="50" value="<%= l_notremitente %>">
	</td>
 </tr>
 <tr>
    <td align="right"><b>Motivo:</b></td>
	<td colspan="3"><input type="Text" name="notmotivo" size="50" maxlength="50" value="<%= l_notmotivo %>"></td>
 </tr>

 <tr>
    <td align="right"><b>Nota:</b></td>
	<td colspan="3"><textarea name="notatxt" rows="10" cols="50"><%= l_notatxt %></textarea></td>
 </tr>
 <tr>
    <td align="right" class="th2" colspan="4" height="1">
		<a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
 </tr>
</table>
</form>

</body>
</html>
