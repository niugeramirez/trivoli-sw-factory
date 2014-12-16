<!--#include virtual="/ticket/shared/inc/sec.inc"-->
<!--#include virtual="/ticket/shared/inc/password.inc"-->
<!--#include virtual="/ticket/shared/inc/encrypt.inc"-->
<%
on error goto 0
 
 Dim l_menu
 Dim l_debug
 Dim l_iduser
 Dim l_pass
 Dim l_tiempo
 Dim l_aux
 Dim l_seguir
 Dim l_usrdb
 Dim l_usrdbpass
 Dim l_hlogfecini
 Dim l_hpassfecini
 
 Dim l_pass_expira_dias
 Dim l_pass_camb_dias
 Dim l_pass_int_fallidos
 Dim l_pass_dias_log
 
 Dim l_usrpasscambiar
 Dim l_msgtxt
 
 l_base 	= trim(request.Querystring("base"))
 l_menu 	= trim(request.Querystring("menu"))
 l_debug	= CInt(request.Querystring("debug"))
 
 l_iduser	= Request.Querystring("usr")
 l_pass		= Request.Querystring("pass")
 l_tiempo	= now
 
 Session("UserName") = "sa"
 Session("Password") = ""
 Session("Time") 	 = l_tiempo
 Session("base") 	 = l_base
 
 %>
 <!--#include virtual="/ticket/shared/db/conn_db.inc"-->
 <%
'------------------------------------------------------------------------------------------------- 
' Procedimiento que impreme mensaje de error.
'------------------------------------------------------------------------------------------------- 
Sub MostrarError (texto)
		if l_menu = "html" then
			%><script>alert('<%= texto %>');</script><%
		else
			response.write "&acceso=No Valido&"
			response.write "&cambiapass=0&"
			response.write "&msgtxt=" & texto & "&"
		end if
end sub
 
 
'------------------------------------------------------------------------------------------------- 
' Bloque principal
'------------------------------------------------------------------------------------------------- 
 
 cn.beginTrans
 
 l_pass_expira_dias	 = valorpol_cuenta(l_iduser, "pass_expira_dias")
 l_pass_camb_dias	 = valorpol_cuenta(l_iduser, "pass_camb_dias")
 l_pass_int_fallidos = valorpol_cuenta(l_iduser, "pass_int_fallidos")
 l_pass_dias_log	 = valorpol_cuenta(l_iduser, "pass_dias_log")
 
 l_usrpasscambiar	 = valoruser_per(l_iduser, "usrpasscambiar")
 
 l_seguir = true
 l_cambiarpass = 0
 l_msgtxt = ""
 if not usuariovalido(l_iduser) then
	MostrarError "Usuario no válido."
	l_seguir = false
 else
	if ctabloqueada(l_iduser) then
		MostrarError "Cuenta Bloqueada. Consulte con el administrador."
		l_seguir = false
	else
 		if valorhistpass (l_iduser, "husrpass") <> Decrypt(c_strEncryptionKey, l_pass) then
			l_seguir = false
			' Verifico si se debe bloquear la cuenta por los logueos fallidos.
			l_aux = CInt(logueosfallidos(l_iduser)) + 1
			if l_aux >= CInt(l_pass_int_fallidos) then
				bloquearcuenta l_iduser, -1
				bajapass l_iduser
				MostrarError "Cuenta Bloqueada por intentos fallidos."
			else
				actlogfallidos l_iduser, l_aux
				MostrarError "Contraseña incorrecta."
			end if
		else
			' Verifico si debe cambiar la contraseña por ser el primer logueo.
			if CInt(l_usrpasscambiar) = -1 then
				l_cambiarpass = -1
			else
				' Verifico si se pasaron los dias sin loguearse permitidos por el sistema.
				l_hlogfecini = valorhistlog (l_iduser, "hlogfecini")
				if l_hlogfecini = "" then
					l_hlogfecini = date()
				end if
				l_aux = datediff("d", l_hlogfecini, date())
				if CInt(l_pass_dias_log) <> 0 and CInt(l_pass_dias_log) <= CInt(l_aux) then
					bloquearcuenta l_iduser, -1
					bajapass l_iduser
					MostrarError "Cuenta Bloqueada. Plazo exedido sin loguearse."
					l_seguir = false
				else
					' Verifico si se expidio(vencio) la contraseña
					l_hpassfecini = valorhistpass (l_iduser, "hpassfecini")
					if l_hpassfecini = "" then
						l_hpassfecini = date()
					end if
					l_aux = datediff("d", l_hpassfecini, date())
					if CInt(l_pass_expira_dias) <> 0 and CInt(l_pass_expira_dias - 1) <= CInt(l_aux) then
						if l_pass_expira_dias - 1 = l_aux then
							l_msgtxt = "Su contraseña expira mañana."
						else
							bloquearcuenta l_iduser, -1
							bajapass l_iduser
							MostrarError "Cuenta Bloqueada. Expiro su contraseña."
							l_seguir = false
						end if
					else
						' Verifico si debe cambiar la contraseña.
						l_aux = datediff("d", l_hpassfecini, date())
						if CInt(l_pass_camb_dias) <> 0 and CInt(l_pass_camb_dias) <= CInt(l_aux) then
							l_cambiarpass = -1
						end if
					end if
				end if
			end if
		end if
	end if
 end if
 
 
 if l_seguir then
	l_usrdb 	= valoruser_per (l_iduser, "usrdb")
	l_usrdbpass = valoruser_per (l_iduser, "usrdbpass")
	
	cn.commitTrans
	cn.close
	
	if conectar (cn, l_usrdb, l_usrdbpass, l_base) then
		if CInt(l_cambiarpass) = -1 then
			if l_menu = "html" then
				%>
				<script src="/ticket/shared/js/fn_windows.js"></script>
				<script>parent.location = "../../lanzador/lanzador2.asp?tipo=pass";//abrirVentana('/ticket/sup/cambio_clave_sup_00.asp', '', 350, 200);
//				if (Nuevo_Dialogo(window, '/ticket/sup/cambio_clave_sup_00.asp?iduser=<%'= l_iduser %>', 50, 40)){
				</script>
				<%
			else
				response.write "&acceso=Valido&"
				response.write "&cambiapass=-1&"
				response.write "&msgtxt=" & l_msgtxt & "&"
			end if
		end if
	else
		if err then
			if l_menu = "html" then
				if l_debug = -1 then
					%><script>parent.document.FormVar.desc.value = "<%= Err.Description %>";</script><%
				else
					%><script>parent.document.FormVar.desc.value = "Acceso no válido.";</script><%
				end if
			else
				if l_debug = -1 then
					MostrarError Err.Description
				else
					MostrarError "Acceso no válido"
				end if
			end if
			l_seguir = false
		end if
	end if
 else
 	Session.Abandon
 	cn.commitTrans
 end if
 
 
 if CInt(l_seguir) = -1 then
	' Ingreso en la base de datos el logueo del usuario
	ingresarlogueo l_iduser
	
	Session("loguinUser") = l_iduser
	
	if l_menu = "html" then
		%><script>parent.document.location = "../../lanzador/lanzador3.asp";</script><%
	else
		if CInt(l_cambiarpass) <> -1 then
			response.write "&acceso=Valido&"
			response.write "&cambiapass=0&"
			response.write "&msgtxt=" & l_msgtxt & "&"
		end if
	end if
 end if
 
 Set cn = nothing
 Set l_rs = nothing
 Set l_cm = nothing
%>
