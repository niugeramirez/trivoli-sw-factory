<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 

'Archivo: pol_cuenta_seg_02.asp
'Descripción: ABM de Políticas de cuenta
'Autor: Alvaro Bayon
'Fecha: 21/02/2005

 Dim l_rs
 Dim l_sql
 
 Dim l_pass_expira_dias
 Dim l_pass_camb_dias
 Dim l_pass_longitud
 Dim l_pass_historia
 Dim l_pass_int_fallidos
 Dim l_pass_dias_log
 Dim l_pass_cambiar
 Dim l_pol_desc
 
 Dim l_pol_nro
 Dim l_tipo
 
 l_tipo = request.QueryString("tipo")
 l_pol_nro = request.QueryString("pol_nro")
 
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
%>
<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Pol&iacute;ticas de Cuentas - Ticket</title>
</head>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_valida.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_numeros.js"></script>
<script>
function Validar_Formulario(){
	if (Trim(document.datos.pol_desc.value) == ""){
		alert("Debe ingresar una Descripción.");
		document.datos.pol_desc.focus();
		return false;
	}
	if(!stringValido(document.datos.pol_desc.value)){		
		alert("La Descripción contiene caracteres inválidos.");
		document.datos.pol_desc.focus();
		return false;
	}

	if (Trim(document.datos.pass_expira_dias.value) == ""){
		alert("Debe ingresar la cantidad de días en que la Contraseña expira.");
		document.datos.pass_expira_dias.focus();
		return false;
	}
	if (!validanumero(document.datos.pass_expira_dias, 4, 0)){
		alert("La cantidad de días no es un número válido.");
		document.datos.pass_expira_dias.focus();
		return false;
	}
	if ((!document.datos.exp_dias[0].checked)&&(document.datos.pass_expira_dias.value<=0)){
		alert("La cantidad de días debe ser mayor que cero.");
		document.datos.pass_expira_dias.focus();
		return false;
	}

	if (Trim(document.datos.pass_camb_dias.value) == ""){
		alert("Debe ingresar la cantidad de días en que se permite cambiar la Contraseña.");
		document.datos.pass_camb_dias.focus();
		return false;
	}
	if (!validanumero(document.datos.pass_camb_dias, 4, 0)){
		alert("La cantidad de días no es un número válido.");
		document.datos.pass_camb_dias.focus();
		return false;
	}
	if ((!document.datos.camb_dias[0].checked)&&(document.datos.pass_camb_dias.value<=0)){
		alert("La cantidad de días debe ser mayor que cero.");
		document.datos.pass_camb_dias.focus();
		return false;
	}

	if (Trim(document.datos.pass_int_fallidos.value) == ""){
		alert("Debe ingresar la cantidad de Intentos fallidos.");
		document.datos.pass_int_fallidos.focus();
		return false;
	}
	if (!validanumero(document.datos.pass_int_fallidos, 3, 0)){
		alert("La cantidad de Intentos fallidos no es un número válido.");
		document.datos.pass_int_fallidos.focus();
		return false;
	}
	if (document.datos.pass_int_fallidos.value==0){
		alert("La cantidad de Intentos fallidos debe ser mayor a cero.");
		document.datos.pass_int_fallidos.focus();
		return false;
	}

	if (Trim(document.datos.pass_dias_log.value) == ""){
		alert("Debe ingresar la cantidad de Días.");
		document.datos.pass_dias_log.focus();
		return false;
	}
	if (!validanumero(document.datos.pass_dias_log, 4, 0)){
		alert("La cantidad de Días no es un número válido.");
		document.datos.pass_dias_log.focus();
		return false;
	}
	if (document.datos.pass_dias_log.value.parseInt<0){
		alert("La cantidad de Días debe ser mayor o igual que cero.");
		document.datos.pass_dias_log.focus();
		return false;
	}

	if (Trim(document.datos.pass_longitud.value) == ""){
		alert("Debe ingresar la longitud mínima de la Contraseña.");
		document.datos.pass_longitud.focus();
		return false;
	}
	if (!validanumero(document.datos.pass_longitud, 4, 0)){
		alert("La longitud no es un número válido.");
		document.datos.pass_longitud.focus();
		return false;
	}
	if ((document.datos.longitud[1].checked)&&(document.datos.pass_longitud.value<=0)){
		alert("La longitud debe ser mayor que cero.");
		document.datos.pass_longitud.focus();
		return false;
	}
	
	if (Trim(document.datos.pass_historia.value) == ""){
		alert("Debe ingresar la cantidad de Contraseñas a recordar.");
		document.datos.pass_historia.focus();
		return false;
	}
	if (!validanumero(document.datos.pass_historia, 4, 0)){
		alert("La cantidad de Contraseñas no es un número válido.");
		document.datos.pass_historia.focus();
		return false;
	}
	if ((document.datos.historia[1].checked)&&(document.datos.pass_historia.value<=0)){
		alert("La cantidad de Contraseñas debe ser mayor que cero.");
		document.datos.pass_historia.focus();
		return false;
	}


	document.datos.submit();
	
}

function radioclick(nombre){
if (nombre=='exp_dias') {
	document.datos.pass_expira_dias.value = 0; 
	document.datos.pass_expira_dias.readOnly = true; 
	document.datos.pass_expira_dias.className = "deshabinp"; 
	}
if (nombre=='camb_dias') {
	document.datos.pass_camb_dias.value = 0; 
	document.datos.pass_camb_dias.readOnly = true; 
	document.datos.pass_camb_dias.className = "deshabinp"; 
	}
if (nombre=='longitud') {
	document.datos.pass_longitud.value = 0; 
	document.datos.pass_longitud.readOnly = true; 
	document.datos.pass_longitud.className = "deshabinp"; 
	}
if (nombre=='historia') {
	document.datos.pass_historia.value = 0; 
	document.datos.pass_historia.readOnly = true; 
	document.datos.pass_historia.className = "deshabinp"; 
	}
}
function hab_texto(nombre){
if (nombre=='exp_dias') {
	document.datos.pass_expira_dias.readOnly = false; 
	document.datos.pass_expira_dias.className = "habinp"; 
	}
if (nombre=='camb_dias') {
	document.datos.pass_camb_dias.readOnly = false; 
	document.datos.pass_camb_dias.className = "habinp"; 
	}
if (nombre=='longitud') {
	document.datos.pass_longitud.readOnly = false; 
	document.datos.pass_longitud.className = "habinp"; 
	}
if (nombre=='historia') {
	document.datos.pass_historia.readOnly = false; 
	document.datos.pass_historia.className = "habinp"; 
	}
}

function Nuevo_Dialogo(w_in, pagina, ancho, alto){
	return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}

function Ayuda_Fecha(txt){
 var jsFecha = Nuevo_Dialogo(window, '/serviciolocal/shared/js/calendar.html', 16, 15);

	if (jsFecha == null) txt.value = ''
	else txt.value = jsFecha;
}

</script>
<%
 if l_tipo = "A" then
 	l_pol_nro			= ""
 	l_pol_desc 			= ""
	l_pass_expira_dias 	= 0
	l_pass_camb_dias	= 0
	l_pass_int_fallidos = 3
	l_pass_dias_log 	= 1
	l_pass_cambiar 		= 0
	l_pass_longitud 	= 5
	l_pass_historia 	= 3
 else
	l_sql = 		"SELECT pol_desc, pass_expira_dias, pass_camb_dias, pass_int_fallidos, pass_dias_log, "
	l_sql = l_sql & "pass_cambiar, pass_longitud, pass_historia "
	l_sql = l_sql & "FROM pol_cuenta "
	l_sql = l_sql & "WHERE pol_nro = " & l_pol_nro
	rsOpen l_rs, cn, l_sql, 0 
	if not l_rs.eof then
		l_pol_desc 			= l_rs("pol_desc")
		l_pass_expira_dias 	= l_rs("pass_expira_dias")
		l_pass_camb_dias 	= l_rs("pass_camb_dias")
		l_pass_int_fallidos = l_rs("pass_int_fallidos")
		l_pass_dias_log 	= l_rs("pass_dias_log")
		l_pass_cambiar 		= l_rs("pass_cambiar")
		l_pass_longitud 	= l_rs("pass_longitud")
		l_pass_historia 	= l_rs("pass_historia")
	end if
	l_rs.Close
 end if

%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="Javascript:document.datos.pol_desc.focus();">
<form target="valida" name="datos" action="pol_cuenta_seg_03.asp?Tipo=<%= l_tipo %>" method="post">
<input type="Hidden" name="pol_nro" value="<%= l_pol_nro %>">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
	<td class="th2" height="1" colspan="2">Datos de la Pol&iacute;tica</td>
	<td class="barra" align="right"><a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a></td>
</tr>
<tr>
	<td colspan="3" height="100%" width="100%">
		<table cellpadding="0" cellspacing="0" border="0">
			<tr>
				<td width="50%"></td>
				<td align="center">
					<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td colspan="2" align="center">
							<table cellpadding="0" cellspacing="0" width="1">
								<tr>
									<td align="right"><b>Descripci&oacute;n:</b></td>
									<td colspan="3" align="left">
										<input type="text" name="pol_desc" size="55" maxlength="50" value="<%= l_pol_desc %>">
									</td>
								</tr>
							</table>
						</td>
					</tr>

					<tr>
						<td nowrap>
						<input type="radio" name="exp_dias"  value="r1" <% if l_pass_expira_dias = 0 then%>Checked<%end if %>  onclick="Javascript:radioclick('exp_dias');"><b>La Contraseña nunca expira</b>
						</td>														   
						<td nowrap>
						<input type="radio" name="camb_dias"  value="0" <% if l_pass_camb_dias = 0 then%>Checked<%end if %> onclick="Javascript:radioclick('camb_dias');"><b>No exigir cambios de Contraseña</b>
						</td>
					</tr>
					<tr>
						<td  nowrap>
						<input type="radio" name="exp_dias"  value="1" <% if l_pass_expira_dias <> 0 then%>Checked<%end if%> onclick="Javascript:hab_texto('exp_dias');">
						<b>Expira en</b> <input type="text" name="pass_expira_dias" <% if l_pass_expira_dias = 0 then %>class="deshabinp" readonly<% End If %> size="3" maxlength="3" value="<%= l_pass_expira_dias %>">
						<b>d&iacute;as</b>
						</td>														   
						<td  nowrap>
						<input type="radio" name="camb_dias"  value="1" <% if l_pass_camb_dias <> 0 then%>Checked<%end if %> onclick="Javascript:hab_texto('camb_dias');">
						<b>Exigir cambios en</b> <input type="text" name="pass_camb_dias" <% if l_pass_camb_dias = 0 then %>class="deshabinp" readonly<% End If %> size="3" maxlength="3" value="<%= l_pass_camb_dias %>">
						<b>d&iacute;as</b>
						</td>
					</tr>
					<tr>
						<td  colspan="2"><hr></td>
					</tr>
					<tr>
						<td colspan="2" nowrap align="center" ><b>Bloquear despu&eacute;s de</b> <input type="text" name="pass_int_fallidos" size="3" maxlength="3" value="<%= l_pass_int_fallidos %>"> <b>intentos fallidos</b></td>
					</tr>
					<tr>
						<td colspan="2" nowrap  align="center"><b>Bloquear si en</b> <input type="text" name="pass_dias_log" size="3" maxlength="3" value="<%= l_pass_dias_log %>"> <b>d&iacute;as no se logueo</b></td>
					</tr>

					<tr>
						<td colspan="2" nowrap  align="center">
						<input type="checkbox" name="pass_cambiar" <% if CInt(l_pass_cambiar) = -1 then%>Checked<%end if%>><b>Cambiar la Contraseña al primer logueo</b>
						</td>
					</tr>
					
					<tr>
						<td  colspan="2"><hr></td>
					</tr>
					
					<tr>
						<td nowrap>
						<input type="radio" name="longitud"  value="0" <% if l_pass_longitud = 0 then%>Checked<%end if %> onclick="Javascript:radioclick('longitud');"><b>Permitir Contraseña en blanco</b>
						</td>														   
						<td nowrap>
						<input type="radio" name="historia"  value="0" <% if l_pass_historia = 0 then%>Checked<%end if %> onclick="Javascript:radioclick('historia');"><b>No mantener histórico de contraseñas</b>
						</td>
					</tr>
					<tr>
						<td nowrap>
						<input type="radio" name="longitud"  value="1" <% if l_pass_longitud <> 0 then%>Checked<%end if %>  onclick="Javascript:hab_texto('longitud');">
						<b>Longitud m&iacute;nima</b> 
						<input type="text" name="pass_longitud" size="3" maxlength="3" value="<%= l_pass_longitud %>" <% if l_pass_longitud = 0 then%>class="deshabinp" readonly<%end if %> > <b>caracteres</b>
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						</td>														   
						<td nowrap>
						<input type="radio" name="historia"  value="1" <% if l_pass_historia <> 0 then%>Checked<%end if %> onclick="Javascript:hab_texto('historia');">
						<b>Recordar</b>
						<input type="text" name="pass_historia" size="3" maxlength="3" value="<%= l_pass_historia %>" <% if l_pass_historia = 0 then%>class="deshabinp" readonly<%end if %>> <b>contraseñas</b>	   
						</td>
					</tr>
					<tr>
						<td height="10" colspan="2"></td>
					</tr>
					</table>
				</td>
				<td width="50%"></td>
			</tr>
		</table>
	</td>
</tr>

<tr>
    <td align="right" class="th2" colspan="3">
		<a class=sidebtnABM style="cursor:hand" onclick="Javascript:Validar_Formulario()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close();">Cancelar</a>
	</td>
</tr>
</table>
<iframe name="valida" style="visibility=hidden;" src="" width="100%" height="100%"></iframe>
</form>
<%
	set l_rs = nothing
	cn.close
	set cn = nothing
%>
</body>
</html>
