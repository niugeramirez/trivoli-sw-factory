<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--#include virtual="/turnos/shared/inc/password.inc"-->
<!-----------------------------------------------------------------------------------
Archivo     : cambio_clave_seg_00.asp
Descripcion : Permite cambiar la clave del usuario.
Autor       : Favre F.
Creacion    : 26/07/2004
Modificar	:
-------------------------------------------------------------------------------------
-->
<%
'on error goto 0
 Dim l_iduser
 
 Dim l_pass_expira_dias
 Dim l_pass_longitud
 Dim l_pass_int_fallidos
 Dim l_pass_dias_log
 Dim l_pass_historia
 Dim l_aux
 Dim l_txt
 Dim l_ctabloqueada
 Dim l_pol_nro
 
 Dim l_hpassfecini
 Dim l_hlogfecini
 
 l_iduser = Session("UserName")
 
 
'-------------------------------------------------------------------------------------------------
' Verifico si la cuenta del usuario esta bloqueada
'-------------------------------------------------------------------------------------------------
 l_ctabloqueada = ctabloqueada(l_iduser)
 
 
 If CInt(l_ctabloqueada) = -1 then
 	l_txt = "Cuenta Bloqueada. Consulte con el administrador."
 else
	' Verifico los valores de la politicas de cuentas
	l_pol_nro			= valoruser_pol_cuenta(l_iduser, "pol_nro")
	l_pass_expira_dias	= CInt(valorpol_cuenta(l_pol_nro, "pass_expira_dias"))
	l_pass_longitud		= CInt(valorpol_cuenta(l_pol_nro, "pass_longitud"))
	l_pass_dias_log	 	= CInt(valorpol_cuenta(l_pol_nro, "pass_dias_log"))
	l_pass_historia		= CInt(valorpol_cuenta(l_pol_nro, "pass_historia"))
	
	
	' Verifico si se expiro(vencio) la contraseña
	l_hpassfecini = valorhistpass (l_iduser, "hpassfecini")
	if l_hpassfecini = "" then
		l_hpassfecini = date()
	end if
	l_aux = datediff("d", l_hpassfecini, date())
	if l_pass_expira_dias <> 0 and l_pass_expira_dias <= l_aux then
		' Bloqueo la cuenta del usuario
		bloquearcuenta l_iduser, -1
		
		' Doy de baja el password
		bajapass l_iduser
		
		l_ctabloqueada = -1
		l_txt = "Cuenta Bloqueada. Expiro su contraseña."
	else
		' Verifico si se pasaron los dias sin loguearse permitidos por el sistema.
		l_hlogfecini = valorhistlog (l_iduser, "hlogfecini")
		if l_hlogfecini = "" then
			l_hlogfecini = date()
		end if
		l_aux = datediff("d", l_hlogfecini, date())
		if l_pass_dias_log <> 0 and l_pass_dias_log <= l_aux then
			' Bloqueo la cuenta del usuario
			bloquearcuenta l_iduser, -1
			
			' Doy de baja el password
			bajapass l_iduser
			
			l_ctabloqueada = -1
			l_txt = "Cuenta Bloqueada. Plazo exedido sin loguearse."
		end if
	end if
 end if
 
%>
<html>
<head>
<link href="/turnos/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Contraseñas - Ticket</title>
</head>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script src="/turnos/shared/js/fn_numeros.js"></script>
<script>
function Validar_Formulario(){
	if (document.datos.ctabloqueada.value == -1)
		alert("Cuenta Bloqueada. Consulte con el administrador.");
	else{
		document.datos.action = "cambio_clave_seg_01.asp?menu=pass";
		document.datos.submit();
	}
}


function MostrarError(nro){
	switch (nro){
		case "1":
			// Se bloqueo la cuenta por superar la cantidad de intentos fallidos permitidos.
			// Usuario no valido.
			document.datos.usrpass.className = "deshabinp";
			document.datos.usrpass.readOnly = true;
			document.datos.usrpassnuevo.className = "deshabinp";
			document.datos.usrpassnuevo.readOnly = true;
			document.datos.usrconfirm.className = "deshabinp";
			document.datos.usrconfirm.readOnly = true;
			break;
		case "2":
			// Contraseña incorrecta.
			document.datos.usrpass.focus();
			document.datos.usrpass.select();
			break;
		case "3":
			// Contraseña en blanco	no esta permitida.
			// Longitud del nuevo password no es valida
			// El nuevo password coindice con un historico.
			document.datos.usrpassnuevo.focus();
			document.datos.usrpassnuevo.select();
			break;
		case "4":
			// Confirmacion no es coincidente
			document.datos.usrconfirm.focus();
			document.datos.usrconfirm.select();
			break;
	}
}
</script>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="document.datos.usrpass.focus();" onunload="window.returnValue = document.datos.passvalido.value;">
<form name="datos" target="ifrm_oculto" action="" method="post">
<input type="Hidden" name="passvalido" value="0">
<input type="Hidden" name="ctabloqueada" value="<%= l_ctabloqueada %>">
<input type="Hidden" name="pass_longitud" value="<%= l_pass_longitud %>">


<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr style="border-color :CadetBlue;">
	<td align="left" colspan=3 class="barra" height="1">Datos de la Contraseña</td>
	<td class="barra" align="right"><a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a></td>
</tr>
<tr>
    <td align="right" ><b>Id.Usuario:</b></td>
	<td colspan=2><input type="text" name="iduser" size="21" class="deshabinp" readonly maxlength="20" value="<%= l_iduser %>"></td>
	<td></td>
</tr>
<tr>
    <td align="right" ><b>Contrase&ntilde;a:</b></td>
	<td colspan=3><input type="password" name="usrpass" size="31" maxlength="30" value="" <% If l_ctabloqueada then %>class="deshabinp" readonly<% End If %>></td>
</tr>
<tr>
    <td colspan=4 height="5"></td>
</tr>
<tr>
    <td align="right" ><b>Nueva:</b></td>
	<td colspan=3><input type="Password" name="usrpassnuevo" size="31" maxlength="30" <% If l_ctabloqueada then %>class="deshabinp" readonly<% End If %>></td>
</tr>
<tr>
    <td align="right" ><b>Confirmaci&oacute;n:</b></td>
	<td colspan=3><input type="Password" name="usrconfirm" size="31" maxlength="30" <% If l_ctabloqueada then %>class="deshabinp" readonly<% End If %>></td>
</tr>

<!--tr>
	<td colspan=4 height="30%">
		<iframe name="ifrm_oculto" src="#" width="100%" height="100%"></iframe>
	</td>
</tr-->

<tr>
    <td align="right" class="th2"  height="1" colspan="4">
		<a class=sidebtnABM style="cursor:hand" onclick="Javascript:Validar_Formulario()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
		<iframe name="ifrm_oculto" src="#" width="0" height="0" style="visibility: hidden;"></iframe>
	</td>
</tr>
</table>
</form>
<% If CInt(l_ctabloqueada) = -1 then %>
	<script>
	alert('<%= l_txt %>');
	</script>
<% End If %>
</body>
</html>
