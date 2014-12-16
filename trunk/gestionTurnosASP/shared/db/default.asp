<%
' Opciones de los parametros que devuelve para el caso de flash
' 	ACCESO		|	 CAMBIAPASS		|	  MSGTXT		|	DESCRIPCION
'---------------|-------------------|-------------------|---------------------------
'	Valido		|		-1			|	Indistinto		|	Cambia la contraseña
'	Valido		|		0			|		''			|	Acceso normal al Sistema (Usuario/contraseña validos)
'	Valido		|		0			|	Un mensaje		|	Muestra un mensaje de Advertencia y permite el ingreso al sistema
'	No Valido	|		0			|	Un mensaje		|	Muestra un mensaje de Error y se queda en el loguin (NO permite el ingreso al sistema)
'
'on error goto 0
 Const c_strEncryptionKey = "56238"

 
 Dim l_pass_expira_dias
 Dim l_pass_camb_dias
 Dim l_pass_int_fallidos
 Dim l_pass_dias_log
 Dim l_usrpasscambiar
 Dim l_seguir
 Dim l_cambiarpass
 Dim l_msgtxt
 Dim l_aux
 Dim l_hlogfecini
 Dim l_hpassfecini
 Dim l_MsgAdv
 Dim l_pol_nro
 
 Dim l_iduser
 Dim l_pass
 Dim l_seg_NT
 Dim l_baseBD
 Dim l_menu
 Dim l_debug
 
 l_iduser	= lcase(Request.Form("usr"))
 l_pass	 	= Request.Form("pass")
 l_seg_NT	= CInt(Request.Form("seg_NT"))
 l_baseBD 	= trim(Request.Form("base"))
 l_menu 	= trim(Request.Form("menu"))
 l_debug 	= CInt(Request.Form("debug"))
 


'Response.write "<script>alert('"&  l_iduser &".');</script>"
'Response.write "<script>alert('"&   l_pass &".');</script>"
'Response.write "<script>alert('"&   l_seg_NT &".');</script>"
'Response.write "<script>alert('"&   l_baseBD &".');</script>"
'Response.write "<script>alert('"&   l_menu &".');</script>"
'Response.write "<script>alert('"&   l_debug &".');</script>"
 
 if l_seg_NT = -1 then
 	l_iduser = replace(l_iduser, "#@#", "\")
 end if
 
'------------------------------------------------------------------------------------------------- 
' Procedimiento que imprime mensaje de error.
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
' Procedimiento que imprime mensaje de advertencia.
'------------------------------------------------------------------------------------------------- 
Sub MostrarMsgAdv (texto)
	if l_menu = "html" then
		%><script>alert('<%= texto %>');</script><%
	else
		response.write "&acceso=Valido&"
		response.write "&cambiapass=0&"
		response.write "&msgtxt=" & texto & "&"
	end if
end sub
 
 
'------------------------------------------------------------------------------------------------- 
' Bloque principal
'------------------------------------------------------------------------------------------------- 
 
%>
 <!--#include virtual="/serviciolocal/shared/inc/encrypt.inc"-->
 <!--#include virtual="/serviciolocal/shared/db/conn.inc"-->
<%

 Session("UserName") = "sa" ' l_iduser'
 Session("Password") = "" 'l_pass	'""


' if Cint(l_baseBD) = 2 then
'	' Esto esta cableado para Bahia Blanca, ya que en RHPROX2 no se puede generar el usuario ess
'	Session("UserName") = "sa"
'	Session("Password") = ""
' end if
 Session("base") = l_baseBD
' Session("base_z") = l_baseBD
 Session("Time") = now
 
%>
 <!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<%
  
 l_seguir = true
 l_MsgAdv = false
 l_msgtxt = ""
  
ON ERROR resume next
 if err then
 	' La coneccion a la base de datos da error con el usuario 'ess'
	if l_menu = "html" then
		if l_debug = -1 then
			%><script>parent.document.FormVar.desc.value = "<%= Err.Description %>";</script><%
		else
			%><script>parent.document.FormVar.desc.value = "Acceso no valido";</script><%
		end if
	else
		if l_debug = -1 then
			response.write "&acceso=" & Err.Description & "&"
		else
			response.write "&acceso=Acceso no valido&"
		end if
	end if
	l_seguir = false
 else
 	' La conexion a la base de datos es valida con el usuario 'sa'
	
	if l_seg_NT = 0 then
		%>
		<!--#include virtual="/serviciolocal/shared/inc/password.inc"-->
		<%
		
	 	if not usuariovalido(l_iduser) then
			MostrarError "Usuario no válido."
			l_seguir = false
	 	else
			l_pol_nro = valoruser_pol_cuenta(l_iduser, "pol_nro")
			
		 	l_pass_expira_dias	 = valorpol_cuenta(l_pol_nro, "pass_expira_dias")
		 	l_pass_camb_dias	 = valorpol_cuenta(l_pol_nro, "pass_camb_dias")
		 	l_pass_int_fallidos  = valorpol_cuenta(l_pol_nro, "pass_int_fallidos")
		 	l_pass_dias_log	 	 = valorpol_cuenta(l_pol_nro, "pass_dias_log")
		 	
		 	l_usrpasscambiar	 = valoruser_per(l_iduser, "usrpasscambiar")
		 	
		 	l_cambiarpass = 0

			
			if ctabloqueada(l_iduser) then
				MostrarError "1Cuenta Bloqueada. Consulte con el administrador."
				l_seguir = false
			else
	 			if valorhistpass (l_iduser, "husrpass") <> Decrypt(c_strEncryptionKey, l_pass,true) then
					l_seguir = false
					' Verifico si se debe bloquear la cuenta por los logueos fallidos.
					l_aux = CInt(logueosfallidos(l_iduser)) + 1
					if CInt(l_pass_int_fallidos) <> 0 and l_aux >= CInt(l_pass_int_fallidos) then
						bloquearcuenta l_iduser, -1
						bajapass l_iduser
						MostrarError "2Cuenta Bloqueada por intentos fallidos."
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
						
						'MostrarError (l_aux)
						'MostrarError (l_pass_dias_log)
						
						
						if CInt(l_pass_dias_log) <> 0 and CInt(l_pass_dias_log) <= CInt(l_aux) then
							bloquearcuenta l_iduser, -1
							bajapass l_iduser
							MostrarError "3Cuenta Bloqueada. Consulte con el administrador."
							l_seguir = false
						else
							' Verifico si expiró la contraseña
							l_hpassfecini = valorhistpass (l_iduser, "hpassfecini")
							if l_hpassfecini = "" then
								l_hpassfecini = date()
							end if
							l_aux = datediff("d", l_hpassfecini, date())
							if CInt(l_pass_expira_dias) <> 0 and CInt(l_pass_expira_dias - 1) <= CInt(l_aux) then
								if l_pass_expira_dias - 1 = l_aux then
									MostrarMsgAdv "Su contraseña expira mañana."
									l_MsgAdv = true
								else
									bloquearcuenta l_iduser, -1
									bajapass l_iduser
									MostrarError "4Cuenta Bloqueada. Expiró su contraseña."
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
	else
	 	if not usuariovalido(l_iduser) then
			MostrarError "Usuario no válido."
			l_seguir = false
		else
			Session("UserName") = l_iduser
			Session("Password") = l_pass
		end if
	end if
 end if
 
 if l_seguir then
 	' Restauro los valores de user y pass a los del usuario, en el caso que no utilice NT
	'if l_seg_NT = 0 then
	'	Session("UserName") = l_iduser
	'	Session("Password") = l_pass
	'end if
	
	' Ingreso en la base de datos el logueo del usuario
	'ingresarlogueo l_iduser
	
	Session("loguinUser") = l_iduser
	
	if l_seg_NT = 0 and CInt(l_cambiarpass) = -1 then
		if l_menu = "html" then
			%><script>parent.location = "../../lanzador/lanzador2.asp?tipo=pass";</script><%
		else
			response.write "&acceso=Valido&"
			response.write "&cambiapass=-1&"
			response.write "&msgtxt=" & l_msgtxt & "&"
		end if
	else
	 	' Restauro los valores de user y pass a los del usuario, en el caso que no utilice NT
		if l_seg_NT = 0 then
			Session("UserName") = l_iduser
			Session("Password") = l_pass
		end if
		
		' Ingreso en la base de datos el logueo del usuario
		ingresarlogueo l_iduser
		
		if l_menu = "html" then
			%><script>//parent.document.location = "../../lanzador/lanzador3.asp";</script><%
			%><script>//parent.document.location = "../asp/asistente_00.asp?wiznro=8";</script><%
			%><script>parent.document.location = "../../config/menu.asp?wiznro=8";</script><%
		else
			if not l_MsgAdv then
				response.write "&acceso=Valido&"
				response.write "&cambiapass=0&"
				response.write "&msgtxt=&"
			end if
		end if
	end if
 else
 	Session.Abandon
 end if

%>

