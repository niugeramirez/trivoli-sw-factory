<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<%

'Archivo: usuarios_seg_02.asp
'Descripción: ABM de usuarios 
'Autor: Alvaro Bayon
'Fecha: 21/02/2005
on error goto 0
 Dim l_iduser
 Dim l_nombre
 Dim l_perfnro
 Dim l_pol_nro
 Dim l_ctabloqueada
 Dim l_usrpasscambiar
 Dim l_usrdb
 Dim l_usrdbpass
 Dim l_usremail
 
 Dim l_rs
 Dim l_sql
 
 Dim tipo
 
 tipo = request("tipo")
 
%>
<html>
<head>
<link href="/turnos/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Usuarios - Ticket</title>
</head>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script src="/turnos/shared/js/fn_numeros.js"></script>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_valida.js"></script>
<script>
function Validar_Formulario(){
	if (Trim(document.datos.iduser.value) == "") {
		alert("Debe ingresar el ID de Usuario.");
		document.datos.iduser.focus();
		return false;
	}

	if(!stringValido(document.datos.iduser.value)){
		alert("El ID contiene caracteres inválidos.");
		document.datos.iduser.focus();
		return false;
	}
	
	if (Trim(document.datos.nombre.value) == ""){
		alert("Debe ingresar el Nombre de Usuario.");
		document.datos.nombre.focus();
		return false;
	}

	if(!stringValido(document.datos.nombre.value)){
		alert("El Nombre contiene caracteres inválidos.");
		document.datos.nombre.focus();
		return false;
	}

	if (document.datos.perfnro.value == "") {
		alert("Debe Seleccionar un Perfil.");
		document.datos.perfnro.focus();
		return false;
	}

	if (document.datos.pol_nro.value == ""){
		alert('Debe seleccionar una Política de Cuenta.');
		document.datos.pol_nro.focus();
		return false;
	}
	
	if (<% If tipo <> "A" then %>document.datos.cambiapass.checked && <% End If %>document.datos.ctabloqueada.checked){
		alert('No se puede bloquear la Cuenta y cambiar la Contraseña al mismo tiempo.');
		return false;
	}
	
	if (document.datos.ctabloqant.value == -1 && !document.datos.ctabloqueada.checked && !document.datos.cambiapass.checked){
		alert("Desbloqueó la Cuenta. Debe ingresar una Contraseña nueva.");
		document.datos.cambiapass.focus();
		return false;
	}
	<% If tipo <> "A" then %>
	if (!document.datos.cambiapass.checked && '<%= tipo %>' == 'A') {
		alert("Debe ingresar una Contraseña.");
		document.datos.cambiapass.focus();
		return false;
	}
	<% End If %>
	
	if (<% If tipo <> "A" then %>document.datos.cambiapass.checked && <% End If %>document.datos.huserpass.value != document.datos.confirmacion.value) {
		alert("La Confirmación no coincide con la Contraseña.");
		document.datos.confirmacion.focus();
		document.datos.confirmacion.select();
		return false;
	}

	if(document.datos.huserpass.readOnly==false){
		if(!stringValido(document.datos.huserpass.value)){
			alert("La Contraseña contiene caracteres inválidos.");
			document.datos.huserpass.focus();
			return false;
		}
		if(!stringValido(document.datos.confirmacion.value)){
			alert("La Confirmación contiene caracteres inválidos.");
			document.datos.confirmacion.focus();
			return false;
		}
	}

	if ((Trim(document.datos.usremail.value) != "")&&(!emailValido(document.datos.usremail.value))) {
		alert("Debe ingresar un email válido.");
		document.datos.usremail.focus();
		return false;
	}
	
	abrirVentanaH('','bblank',0,0);	
	document.datos.submit();
}


function Nuevo_Dialogo(w_in, pagina, ancho, alto){
	return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}


function habilitarPass(){
	if (document.datos.cambiapass.checked){
		document.datos.huserpass.className="habinp"; 
		document.datos.huserpass.readOnly=false;
		document.datos.confirmacion.className="habinp"; 
		document.datos.confirmacion.readOnly=false;
	}else{
		document.datos.huserpass.className="deshabinp"; 
		document.datos.huserpass.readOnly=true;
		document.datos.confirmacion.className="deshabinp"; 
		document.datos.confirmacion.readOnly=true;
	}
}


function habilitarPassBD(){
	if (!document.datos.crearusrbase.checked){
		document.datos.usrdb.className="habinp"; 
		document.datos.usrdb.readOnly=false;
		document.datos.usrdbpass.className="habinp"; 
		document.datos.usrdbpass.readOnly=false;
		document.datos.usrdbconfirmacion.className="habinp"; 
		document.datos.usrdbconfirmacion.readOnly=false;
	}else{
		document.datos.usrdb.className="deshabinp"; 
		document.datos.usrdb.readOnly=true;
		document.datos.usrdbpass.className="deshabinp"; 
		document.datos.usrdbpass.readOnly=true;
		document.datos.usrdbconfirmacion.className="deshabinp"; 
		document.datos.usrdbconfirmacion.readOnly=true;
	}
}


function cambiaperfil(){
	document.datos.pol_nro.value = "" + document.datos.perfnro[document.datos.perfnro.selectedIndex].pol_nro + "";
}

function cambiapolnro(){
	if (document.datos.pol_nro[document.datos.pol_nro.selectedIndex].pass_cambiar == -1)
		document.datos.usrpasscambiar.checked = true;
	else
		document.datos.usrpasscambiar.checked = false;
//	document.datos.usrpasscambiar.value = "" + document.datos.pol_nro[document.datos.pol_nro.selectedIndex].pass_cambiar + "";
}

</script>

<%
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 
 select Case tipo
	Case "A":
		l_iduser = ""
		l_nombre = ""
		l_perfnro = ""
'		l_MRUOrden = 1		
'		l_MRUCant = "0"
		l_pol_nro = ""
		l_ctabloqueada = 0
		l_usrpasscambiar = 0
		l_usrdb = ""
		l_usrdbpass = ""
		l_usremail = ""
	Case "M":
		l_iduser = request("iduser")
		l_sql = 		"SELECT user_per.usrnombre, user_per.perfnro, user_per.MRUOrden, user_per.MRUCant, "
		l_sql = l_sql & "usr_pol_cuenta.pol_nro, user_per.ctabloqueada, user_per.usrpasscambiar, "
		l_sql = l_sql & "user_per.usrdb, user_per.usrdbpass, user_per.usremail "
		l_sql = l_sql & "FROM user_per LEFT JOIN usr_pol_cuenta ON user_per.iduser = usr_pol_cuenta.iduser AND usr_pol_cuenta.upcfecfin IS NULL "
		l_sql = l_sql & "WHERE user_per.iduser = '" & l_iduser & "'"
		rsOpen l_rs, cn, l_sql, 0
		if not l_rs.eof then
			l_nombre		 = l_rs("usrnombre")
			l_perfnro		 = l_rs("perfnro")
'			l_MRUOrden		 = l_rs("MRUOrden")
'			l_MRUCant		 = l_rs("MRUCant")
			l_pol_nro 		 = l_rs("pol_nro")
			l_ctabloqueada 	 = l_rs("ctabloqueada")
			l_usrpasscambiar = l_rs("usrpasscambiar")
			l_usrdb 		 = l_rs("usrdb")
			l_usrdbpass 	 = l_rs("usrdbpass")
			l_usremail 		 = l_rs("usremail")
		end if
		l_rs.Close
 end select
 
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="document.datos.iduser.focus();">
<form name="datos" target="bblank" action="usuarios_seg_03.asp?Tipo=<%=tipo%>" method="post">
<input type="Hidden" name="ctabloqant" value="<%= l_ctabloqueada %>">
<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
	<tr style="border-color :CadetBlue;">
		<td align="left" colspan=3 class="barra" height="1">Datos del Usuario</td>
		<td class="barra" align="right"><a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a></td>
	</tr>
	<tr>
		<td colspan="4">
			<table cellpadding="0" cellspacing="0">
				<tr>
					<td width="50%"></td>
					<td>
						<table cellpadding="0" cellspacing="0">
							<tr>
								<td align="right" nowrap><b>Id.Usuario:</b></td>
								<td colspan=3><input type="text" name="iduser" size="21" maxlength="20" value="<%= l_iduser %>" <%If tipo = "M" then%>readonly class="deshabinp"<%End If%>></td>
							</tr>
							<tr>
								<td align="right" nowrap><b>Nombre Usuario:</b></td>
								<td colspan=3><input type="text" name="nombre" size="30" maxlength="40" value="<%= l_nombre %>"></td>
							</tr>
							<tr>
								<td align="right" nowrap><b>Perfil:</b></td>
								<td align="left" colspan=3>
									<select name="perfnro" size="1" style="width: 280px;" onchange="javascript:cambiaperfil();">
									<option value="" pol_desc="">&laquo;Seleccione una opci&oacute;n&raquo;</option>
									<%
									l_sql = 		"SELECT perfnro, perfnom, pol_nro "
									l_sql = l_sql & "FROM perf_usr ORDER BY perfnom"
									rsOpen l_rs, cn, l_sql, 0
									do until l_rs.eof
										%><option value="<%= l_rs("perfnro") %>" pol_nro="<%= l_rs("pol_nro") %>"><%= l_rs("perfnom") & " (" &  l_rs("perfnro") & ")" %></option><%
										l_rs.movenext
									loop
									l_rs.close
									%>
								</select> 
								<script>document.datos.perfnro.value = '<%= l_perfnro %>' </script>
								</td>
							</tr>
							<!--tr>
								<td align="right" rowspan="2"><b>Men&uacute; usados:</b>
								</td>
								<td align="left" valign="bottom">
								<%'if l_MRUorden = "1" then %> 
								<input TYPE="radio" NAME="MRUorden" VALUE="1" CHECKED><b>M&aacute;s recientes</b><br>
								<%'else%>
								<input TYPE="radio" NAME="MRUorden" VALUE="1"><b>M&aacute;s recientes</b><br>
								<%'end if%>
								</td>
							    <td align="right" rowspan="2"><b>Cantidad:</b></td>
								<td rowspan="2"><input type="text" name="MRUcant" size="4" maxlength="4" value="<%'= l_MRUCant %>"></td>
							</tr>
							<tr>
								<td align="left" valign="top">
								<%'if l_MRUorden = "2" then %> 
								<input TYPE="radio" NAME="MRUorden" VALUE="2" CHECKED><b>M&aacute;s Usados</b><br>
								<%'else%>
								<input TYPE="radio" NAME="MRUorden" VALUE="2"><b>M&aacute;s Usados</b><br>
								<%'end if%>
								</td>
							</tr-->
						
							<tr>
								<td align="right" nowrap><b>Pol&iacute;tica de Cuenta:</b></td>
								<td colspan=3>
								<select name="pol_nro" size="1" style="width: 280px;" onchange="javascript:cambiapolnro();">
									<option value="">&laquo;Seleccione una opci&oacute;n&raquo;</option>
									<%
									l_sql = "SELECT pol_nro, pol_desc, pass_cambiar FROM pol_cuenta "
									
									if  Session("UserName") <> "sa" then
										  l_sql = l_sql & " WHERE pol_desc <> 'Politica Sistemas' "
									else
										  l_sql = l_sql & " WHERE 1 = 1 "
									end if 
  								    l_sql = l_sql & " ORDER BY pol_desc "									
									rsOpen l_rs, cn, l_sql, 0
									do until l_rs.eof
										%><option value="<%= l_rs("pol_nro") %>" pass_cambiar="<%= l_rs("pass_cambiar") %>"><%= l_rs("pol_desc") & " (" &  l_rs("pol_nro") & ")" %></option><%
										l_rs.movenext
									loop
									l_rs.close
									%>
								</select> 
								<script>document.datos.pol_nro.value = '<%= l_pol_nro %>'</script>
								</td>
							</tr>
							<tr>
								<td>&nbsp;</td>
								<td align="left" colspan=3>
								<input type="checkbox" <%if CInt(l_ctabloqueada) = -1  then%>checked<%end if%> id=ctabloqueada name=ctabloqueada >
								<b>Cuenta Bloqueada</b>
								</td>
							</tr>
							<tr>
								<td>&nbsp;</td>
								<td align="left" colspan=3>
								<input type="checkbox" <%if CInt(l_usrpasscambiar) = -1  then%>checked<%end if%> id=usrpasscambiar name=usrpasscambiar>
								<b>Obligación de cambiar la contrase&ntilde;a al próximo logueo</b>
								</td>
							</tr>
							<!--tr>
								<td colspan=4 height="5"></td>
							</tr>
							
							<tr>
							    <td colspan="4" height="10"><b>Restricci&oacute;n de password</b></td>
							</tr-->
							<!--tr>
								<td>&nbsp;</td>
								<td align="left" colspan=3>
								<%'if l_noexpira  then%>
								<input type="checkbox" checked id=checkbox1 name=noexpira >
								<%'else%>
								<input type="checkbox" id=checkbox1 name=noexpira >
								<%'end if%>
								<b>No expira nunca</b></td>
							</tr>
							<tr>
								<td>&nbsp;</td>
								<td align="left" colspan=3>
								<%'if l_nocambia  then%>
								<input type="checkbox" checked id=checkbox1 name=nocambia >
								<%'else%>
								<input type="checkbox" id=checkbox1 name=nocambia >
								<%'end if%>
								<b>No puede cambiarla</b></td>
							</tr-->
							<% If tipo <> "A" then %>
							<tr>
							    <td></td>
								<td colspan=3><input type="checkbox" name=cambiapass onclick="javascript:habilitarPass();"><b>Cambiar Contrase&ntilde;a</b></td>
							</tr>
							<% End If %>
							<tr>
							    <td align="right"  nowrap><b>Contrase&ntilde;a:</b></td>
								<td colspan=3><input type="password" name="huserpass" size="31" maxlength="30" <% If tipo <> "A" then %>class="deshabinp" readonly <% End If %> value=""></td>
							</tr>
							<tr>
							    <td align="right"  nowrap><b>Confirmaci&oacute;n:</b></td>
								<td colspan=3><input type="password" name="confirmacion" size="31" maxlength="30" <% If tipo <> "A" then %>class="deshabinp" readonly <% End If %> value=""></td>
							</tr>
							
							<!--tr>
							    <td></td>
								<td colspan=3><input type="checkbox" name=crearusrbase onclick="javascript:habilitarPassBD();"><b>Crear usuario de BD con el Id. y Contrase&ntilde;a de este usuario</b></td>
							</tr>
							<tr>
							    <td align="right" ><b>Usuario BD:</b></td>
								<td colspan=3><input type="text" name="usrdb" size="31" maxlength="30" value="<%'= l_usrdb %>"></td>
							</tr>
							<tr>
							    <td align="right" ><b>Contrase&ntilde;a BD:</b></td>
								<td colspan=3><input type="password" name="usrdbpass" size="31" maxlength="30" value="<%'= l_usrdbpass %>"></td>
							</tr>
							<tr>
							    <td align="right" ><b>Confirmaci&oacute;n BD:</b></td>
								<td colspan=3><input type="password" name="usrdbconfirmacion" size="31" maxlength="30" value="<%'= l_usrdbpass %>"></td>
							</tr-->
							
							<tr>
							    <td align="right"  nowrap><b>Email:</b></td>
								<td colspan=3><input type="text" name="usremail" size="60" maxlength="80" value="<%= l_usremail%>"></td>
							</tr>


						</table>
					</td>
					<td width="50%"></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
	    <td align="right" class="th2"  height="1" colspan="4">
			<a class=sidebtnABM style="cursor:hand" onclick="Javascript:Validar_Formulario()">Aceptar</a>
			<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
		</td>
	</tr>
</table>
<%
 Set l_rs = nothing
%>
</form>
</body>
</html>
