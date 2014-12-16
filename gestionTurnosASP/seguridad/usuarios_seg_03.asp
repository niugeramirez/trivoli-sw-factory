<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/encrypt.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/password.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<%

'Archivo: usuarios_seg_03.asp
'Descripción: ABM de usuarios 
'Autor: Alvaro Bayon
'Fecha: 21/02/2005

on error goto 0

Const c_strEncryptionKey = "56238"
 Dim l_cm
 Dim l_rs
 Dim l_sql
 Dim l_sql1
 Dim l_sql2
 
 Dim l_tipo
 Dim l_iduser
 Dim l_nombre
 Dim l_perfnro
 Dim l_pol_nro
 Dim l_ctabloqueada
 Dim l_usrpasscambiar
 Dim l_cambiapass
 Dim l_huserpass
 Dim l_crearusrbase
 Dim l_usrdb
 Dim l_usrdbpass
 Dim l_usremail
 Dim l_MRUorden
 Dim l_MRUcant
 
 Dim l_seguir
 Dim l_pass_historia
 
 Dim l_huserpassencryp
 Dim l_tipoerror
 Dim error
 
 Dim l_pass_longitud
 Dim l_pass_int_fallidos
 Dim l_pass_dias_logueo
 
 l_tipo 		= request("tipo")
 
 l_iduser  			= lcase(request.Form("iduser"))
 l_nombre			= request.Form("nombre")
 l_perfnro   		= request.Form("perfnro")
 l_pol_nro			= request.Form("pol_nro")
 l_ctabloqueada 	= request.Form("ctabloqueada")
 l_usrpasscambiar 	= request.Form("usrpasscambiar")
 l_cambiapass		= request.Form("cambiapass")
 l_huserpass  		= request.Form("huserpass")
' l_crearusrbase	= request.Form("crearusrbase")
' l_usrdb			= request.Form("usrdb")
' l_usrdbpass		= request.Form("usrdbpass")
 l_usremail			= request.Form("usremail")
' l_MRUorden			= request.Form("MRUorden")
' l_MRUcant			= request.Form("MRUcant")
 
 
 if len(l_ctabloqueada) > 0 	then l_ctabloqueada = -1	else l_ctabloqueada = 0 	end if
 if len(l_cambiapass) > 0 		then l_cambiapass = -1 		else l_cambiapass = 0 		end if
 if len(l_usrpasscambiar) > 0 	then l_usrpasscambiar = -1 	else l_usrpasscambiar = 0 	end if
' if len(l_crearusrbase) > 0 	then l_crearusrbase = -1 	else l_crearusrbase = 0 	end if
 
 if l_MRUcant = "" 				then l_MRUcant = 0 			end if
 l_MRUorden = 0
 
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 Set l_cm = Server.CreateObject("ADODB.Command")
 
'------------------------------------------------------------------------------------------------------
' Funcion que valida los datos contra la politica de cuenta
'------------------------------------------------------------------------------------------------------
function validar()
	
	l_pass_longitud	= valorpol_cuenta(l_pol_nro, "pass_longitud")
	
	validar = true
	if l_huserpass = "" and l_pass_longitud > 0 then
		l_tipoerror = "2"
		validar = false
	else
		if len(l_huserpass) < l_pass_longitud then
			l_tipoerror = "3"
			validar = false
		end if
	end if
end function
 
 
'------------------------------------------------------------------------------------------------------
' Bloque principal
'------------------------------------------------------------------------------------------------------
' cn.BeginTrans
 
 l_tipoerror = "0"
 ' ALTA
 if l_tipo = "A" then
	l_sql = "SELECT * FROM user_per WHERE iduser = '" & l_iduser & "'"
	rsOpen l_rs, cn, l_sql, 0 
	if not l_rs.eof then
		l_tipoerror = "1"
	else
		if crearUsuario(l_iduser,l_huserpass) then
		
			l_sql1 = 		  "INSERT INTO user_per "
			l_sql1 = l_sql1 & "(iduser, perfnro, usrnombre, usremail, usrdb, usrdbpass, ctabloqueada, usrpasscambiar, MRUorden, MRUcant) "
			l_sql1 = l_sql1 & " VALUES ('" & l_iduser & "'," & l_perfnro & ", '"
			l_sql1 = l_sql1 & l_nombre & "','" & l_usremail & "','" & l_usrdb & "','" & l_usrdbpass & "'," & l_ctabloqueada & "," & l_usrpasscambiar & "," & l_MRUorden & "," & l_MRUcant & ")"
			
			' Valido el nuevo password
			if validar() then
				' Ingreso la configuración para el usuario (user_per)
				l_cm.activeconnection = Cn
				cmExecute l_cm, l_sql1, 0
				
				if l_pol_nro <> "" then
					l_sql2 = 		  "INSERT INTO usr_pol_cuenta (iduser, pol_nro, upcfecini)"
					l_sql2 = l_sql2 & "VALUES ('" & l_iduser & "'," & l_pol_nro & "," & cambiafecha (date(),"MDY","") & ")"
					
					' Ingreso la politica de cuenta para el usuario
					cmExecute l_cm, l_sql2, 0
				end if
				
				' Se realiza el ingreso del nuevo password
				l_huserpassencryp = Decrypt(c_strEncryptionKey, l_huserpass, true)
				ingresarpass l_iduser, l_huserpassencryp
				
				'usuario 1, l_iduser, l_huserpass
				
			end if
		else
			l_seguir = false
			l_tipoerror = "5"
		end if
	end if
	l_rs.close
 ' MODIFICACION
 else
	l_sql = "UPDATE user_per SET "
	l_sql = l_sql & "perfnro = " & l_perfnro & "," 
	l_sql = l_sql & "usrnombre = '" & l_nombre & "',"
	l_sql = l_sql & "usremail = '" & l_usremail & "',"
	l_sql = l_sql & "usrdb = '" & l_usrdb & "',"
	l_sql = l_sql & "usrdbpass = '" & l_usrdbpass & "',"
	l_sql = l_sql & "ctabloqueada = " & l_ctabloqueada & ","
	l_sql = l_sql & "usrpasscambiar = " & l_usrpasscambiar & ","
	l_sql = l_sql & "MRUorden = " & l_MRUorden & ","
	l_sql = l_sql & "MRUcant = " & l_MRUcant
	l_sql = l_sql & " WHERE iduser = '" & l_iduser & "'"
	
	l_seguir = true
	if l_cambiapass then
		' Valido el nuevo password
		if validar() then
			' Valido que no ingrese un password que ya se uso en el historico.
			l_huserpassencryp = Decrypt(c_strEncryptionKey, l_huserpass,true)
			l_pass_historia = valorpol_cuenta(l_pol_nro, "pass_historia")
			
			if passrepetido(l_iduser, l_pass_historia, l_huserpassencryp) then
				l_tipoerror = "4"
				l_seguir = false
			end if
		else
			l_seguir = false
		end if
	end if
	
	
	
	if l_seguir then
		l_cm.activeconnection = Cn
		cmExecute l_cm, l_sql, 0
		
		l_sql2 = "SELECT * FROM usr_pol_cuenta WHERE iduser = '" & l_iduser & "' AND upcfecfin IS NULL"
		rsOpen l_rs, cn, l_sql2, 0
		if not l_rs.eof then
			if CInt(l_rs("pol_nro")) <> CInt(l_pol_nro) then
				l_sql2 = 		  "UPDATE usr_pol_cuenta SET upcfecfin = " & cambiafecha (date(),"MDY","")
				l_sql2 = l_sql2 & " WHERE iduser = '" & l_iduser & "' AND pol_nro = " & l_rs("pol_nro") & " AND upcfecfin IS NULL"
				' Cierro la antigua politica de cuenta del usuario
				l_cm.activeconnection = Cn
				cmExecute l_cm, l_sql2, 0
				
				l_sql2 = 		  "INSERT INTO usr_pol_cuenta (iduser, pol_nro, upcfecini) "
				l_sql2 = l_sql2 & "VALUES ('" & l_iduser & "'," & l_pol_nro & "," & cambiafecha (date(),"MDY","") & ") "
				' Actualiza la politica de cuenta para el usuario
				l_cm.activeconnection = Cn
				cmExecute l_cm, l_sql2, 0
			end if
		else
			l_sql2 = 		  "INSERT INTO usr_pol_cuenta (iduser, pol_nro, upcfecini) "
			l_sql2 = l_sql2 & "VALUES ('" & l_iduser & "'," & l_pol_nro & "," & cambiafecha (date(),"MDY","") & ") "
			' Actualiza la politica de cuenta para el usuario
			l_cm.activeconnection = Cn
			cmExecute l_cm, l_sql2, 0
		end if
		l_rs.close
		' Blanqueo la cantidad de intentos fallidos
		actlogfallidos l_iduser, 0
		
		if l_cambiapass then
			'bajapass l_iduser
			if CambiarPass(l_iduser, l_huserpassencryp) = true then 'l_huserpass
				' Doy de baja el viejo password
				'bajapass l_iduser
				
				' Verifico el historico de password, para mantener la cantidad de pass historicos definida en la politica de cuenta
				'l_pass_historia = valorpol_cuenta (l_pol_nro, "pass_historia")
				'eliminarhistpass l_iduser, l_pass_historia
				' Ingreso el nuevo password
	'			CambiarPass l_iduser, l_huserpass
	'			response.write l_iduser & "-" & l_huserpass 
	'			response.end
	''''''''			ingresarpass l_iduser, l_huserpassencryp
				' Se realiza la actualizacion en la Base de datos correspondiente
	''''''''			usuario 3, l_iduser, l_huserpass
			else
				l_tipoerror = 5
			end if
		end if
	end if
 end if
 
 
 if l_ctabloqueada then
	' Si la cuenta esta bloqueada, doy de baja el password
	bajapass l_iduser
 end if
 
' cn.CommitTrans
 
 if CInt(l_tipoerror) = 0 then
	Response.write "<script>alert('Operación Realizada.');window.opener.opener.ifrm.location.reload();window.opener.close();</script>"
 else
	select case l_tipoerror
		case "1":
			Response.write "<script>alert('Usuario existente.');</script>"
		case "2":
			Response.write "<script>alert('No se permite contraseña en blanco.');</script>"
		case "3":
			Response.write "<script>alert('La longitud mínima es de " & l_pass_longitud & " caracteres.');</script>"
		case "4":
			Response.write "<script>alert('La Contraseña coincide con una histórica.');</script>"
		case "5":
			Response.write "<script>alert('No tiene los permisos para crear o modificar usuarios en la BD.\nConsulte con el administrador del sistema');</script>"
	end select
'	Response.write "<script>history.back();</script>"
 end if
 
 Set l_rs = nothing
 Set l_cm = nothing
 
 Response.write "<script>window.close();</script>"
%>
