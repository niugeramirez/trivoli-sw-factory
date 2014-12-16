<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/encrypt.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/password.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!-- ------------------------------------------------------------------------------------------
Archivo     : cambio_clave_seg_01.asp
Descripcion : Valida datos.
Autor       : Favre F.
Creacion    : 29/07/2004
Modificar	:
----------------------------------------------------------------------------------------------
-->
<%
on error goto 0

Const c_strEncryptionKey = "56238"

 Dim l_cm
 Dim l_rs
 Dim l_sql
 
 Dim l_iduser
 Dim l_menu
 Dim l_usrpass
 Dim l_usrpassnuevo
 Dim l_usrconfirm
 
 Dim l_usrpassant
 Dim l_usrpassencryp
 Dim l_usrpassnueencryp
 Dim l_seguir
 Dim l_aux
 Dim l_pass_longitud
 Dim l_pass_int_fallidos
 Dim l_pass_historia
 Dim l_pol_nro
 
 'l_iduser  		 = Session("UserName")
 l_iduser  		 = Session("loguinUser")
 l_menu			 = request("menu")
 l_usrpass		 = request("usrpass")
 l_usrpassnuevo	 = request("usrpassnuevo")
 l_usrconfirm 	 = request("usrconfirm")
 
 l_iduser  			= lcase(l_iduser)
 l_usrpassant 	 	= valorhistpass (l_iduser, "husrpass")
 l_usrpassencryp 	= Decrypt(c_strEncryptionKey, l_usrpass,true)
 l_usrpassnueencryp = Decrypt(c_strEncryptionKey, l_usrpassnuevo,true)
 
' response.write l_iduser & "<br>"
' response.write l_menu & "<br>"
' response.write l_usrpass & "<br>"
' response.write l_usrpassnuevo & "<br>"
' response.write l_usrconfirm & "<br>"
 
 l_iduser  		 = lcase(l_iduser)
 l_usrpassencryp = Decrypt(c_strEncryptionKey, l_usrpassnuevo,true)
 l_pol_nro = valoruser_pol_cuenta(l_iduser ,"pol_nro")
 
 
'------------------------------------------------------------------------------------------------- 
' Procedimiento que impreme mensaje de error.
'------------------------------------------------------------------------------------------------- 
Sub ActPassword ()
 Dim l_pass_historia
 Dim l_usrpasscambiar
 	
 	'cn.beginTrans
	
	' Blanqueo la cantidad de intentos fallidos
	actlogfallidos l_iduser, 0
	
	' Doy de baja el viejo password
	bajapass l_iduser

	' Verifico el historico de password, para mantener la cantidad de pass historicos definida en la politica de cuenta
	l_pass_historia = valorpol_cuenta (l_pol_nro, "pass_historia")
	eliminarhistpass l_iduser, l_pass_historia

	on error resume next
	' Ingreso el nuevo password
	ingresarpass l_iduser, l_usrpassencryp

	' Verifico si el cambio del password esta definido en el primer logueo
	l_usrpasscambiar = valoruser_per (l_iduser, "usrpasscambiar")
	if CInt(l_usrpasscambiar) = -1 then
		Set l_cm = Server.CreateObject("ADODB.Command")
		l_cm.activeconnection = cn
		l_sql = 		"UPDATE user_per SET usrpasscambiar = 0 "
		l_sql = l_sql & "WHERE iduser = '" & l_iduser & "'"
		cmExecute l_cm, l_sql, 0
		Set l_cm = nothing
	end if
	
	'cn.commitTrans
 	
	select case l_menu
		case "html":
			%><script>alert('Operación Realizada.');</script><%
			%><script>parent.document.location = "../lanzador/lanzador3.asp";</script><%
		case "flash":
			response.write "&acceso=Valido&"
			response.write "&cambiapass=0&"
			response.write "&msgtxt=Operación Realizada.&"
			response.write "&nroerror=&"
		case "pass":
			%><script>alert('Operación Realizada.');window.parent.close();</script><%
	end select
 end sub
 
 
'------------------------------------------------------------------------------------------------- 
' Procedimiento que impreme mensaje de error.
'------------------------------------------------------------------------------------------------- 
Sub MostrarError (estado, cambiapass, msg, nroerror)
	select case l_menu
		case "html":
			if estado = "Valido" then
				ActPassword
			else
				%><script>alert('<%= msg %>');</script><%
				if CInt(nroerror) = 1 then
					%><script>parent.document.location = "../lanzador/lanzador2.asp";</script><%
				end if
			end if
		case "flash":
			if estado = "Valido" then
				ActPassword
			else
				response.write "&acceso=" & estado & "&"
				response.write "&cambiapass=" & cambiapass & "&"
				response.write "&msgtxt=" & msg & "&"
				response.write "&nroerror=" & nroerror & "&"
			end if
		case "pass":
			if estado = "Valido" then
				ActPassword
			else
				%><script>alert('<%= msg %>');parent.MostrarError('<%= nroerror %>')</script><%
			end if
	end select
end sub
 
 
'------------------------------------------------------------------------------------------------- 
' Bloque Principal.
'------------------------------------------------------------------------------------------------- 
 l_seguir = true
 if not usuariovalido(l_iduser) then
	MostrarError "No Valido", 0, "Usuario no válido.", 1
	l_seguir = false
 end if
 
 
 if l_seguir then
	if ctabloqueada(l_iduser) then
		MostrarError "No valido", 0, "Cuenta Bloqueada. Consulte con el administrador.", 1
		l_seguir = false
	end if
 end if
 
 if l_seguir then
	if  l_usrpassant <> Decrypt(c_strEncryptionKey, l_usrpass,true) then
		l_seguir = false
		' Verifico si se debe bloquear la cuenta por los logueos fallidos.
		l_aux = CInt(logueosfallidos(l_iduser)) + 1
		l_pass_int_fallidos = valorpol_cuenta(l_pol_nro, "pass_int_fallidos")
		if l_aux >= CInt(l_pass_int_fallidos) then
			bloquearcuenta l_iduser, -1
			bajapass l_iduser
			MostrarError "No valido", 0, "Cuenta Bloqueada por intentos fallidos.", 1
		else
			actlogfallidos l_iduser, l_aux
			MostrarError "No valido", -1, "Contraseña incorrecta.", 2
		end if
	end if
 end if
 
 
 if l_seguir then
	l_pass_longitud		= CInt(valorpol_cuenta(l_pol_nro, "pass_longitud"))
	if l_pass_longitud > 0 and len(l_usrpassnuevo) = 0 then
		MostrarError "No valido", -1, "No se permite contraseña en blanco.", 3
		l_seguir = false
	else
		if (l_pass_longitud > 0) and (len(l_usrpassnuevo) < l_pass_longitud) then
			MostrarError "No valido", -1, "La longitud mínima es de " & l_pass_longitud & " caracteres.", 3
			l_seguir = false
		end if
	end if
 end if
 
 
 if l_seguir then
	if l_usrpassnuevo <> l_usrconfirm then
		MostrarError "No valido", -1, "La confirmación no es coincidente.", 4
		l_seguir = false
	end if
 end if
 
 
 if l_seguir then
	l_pass_historia = valorpol_cuenta(l_pol_nro, "pass_historia")
	if passrepetido(l_iduser, l_pass_historia, l_usrpassnueencryp) then
		MostrarError "No valido", -1, "La Contraseña coincide con una histórica.", 3
		l_seguir = false
	end if
 end if
 

 if l_seguir then
	ActPassword
 	' Restauro los valores de user y pass a los del usuario, en el caso que no utilice NT
	Session("UserName") = l_iduser
	'Session("Password") = l_usrpass
	Session("Password") = l_usrpassnuevo
	
	' Ingreso en la base de datos el logueo del usuario
	ingresarlogueo l_iduser
 end if
 
%>
