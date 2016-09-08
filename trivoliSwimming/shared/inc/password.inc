<!--#include virtual="/trivoliSwimming/shared/inc/fecha.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/adovbs.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/encrypt.inc"-->
<%

on error goto 0

' Creador: Fernando Favre
' Fecha  : 28-07-2004
 
' Funciones
' usuariovalido	(user)			: Funcion que indica si el usuario es válido.
' ctabloqueada (user)			: Funcion que indica si esta bloqueada la cuenta de un usuario.
' bloquearcuenta (user, valor) 	: Procedimiento que bloquea o desbloquea la cuenta de un usuario.
' valorpol_cuenta (polnro, campo)	: Funcion que devuelve el valor del campo de la politica de cuenta.
' valoruser_pol_cuenta(user, campo)	: Funcion que devuelve el valor de la politica de cuenta definida del usuario.
' valorhistlog (user, campo) 	: Funcion que devuelve el valor del campo del historico del logueo del usuario.
' valorhistpass (user, campo) 	: Funcion que devuelve el valor del campo del historico de contraseñas del usuario.
' valoruser_per (user, campo) 	: Funcion que devuelve el valor del campo de la tabla que mantiene a los usuarios (user_per).
' logueosfallidos (user) 		: Funcion que devuelve la cantidad de logueos fallidos del usuario.
' actlogfallidos (user, cant) 	: Procedimiento que actualiza la cantidad de logueos fallidos.
' ingresarlogueo (user) 		: Procedimiento que ingresa el logueo del usuario en la Base de Datos.
' bajapass (user) 				: Procedimiento que da de baja un password en la Base de Datos.
' ingresarpass (user, pass) 	: Procedimiento que ingresa un nuevo password en la Base de Datos.
' passrepetido(user, cant, usrpass) : Funcion que valida si se esta repitiendo el password en base al historico.
' eliminarhistpass (user, cant) : Procedimiento que mantiene cant de password en el historico, eliminando todos los demas.
 
 
'------------------------------------------------------------------------------------------------- 
' Funcion que indica si el usuario es válido.
'------------------------------------------------------------------------------------------------- 
function usuariovalido (user)
 Dim l_rs
 Dim l_sql
 	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = 		"SELECT user_per.ctabloqueada "
	l_sql = l_sql & "FROM user_per "
	l_sql = l_sql & "WHERE user_per.iduser = '" & user & "'"
	rsOpen l_rs, cn, l_sql, 0
	usuariovalido = false
	if not l_rs.eof then
		usuariovalido = true
	end if
	l_rs.close
	Set l_rs = nothing
end function
 
 
'------------------------------------------------------------------------------------------------- 
' Funcion que indica si esta bloqueada la cuenta de un usuario.
'------------------------------------------------------------------------------------------------- 
function ctabloqueada (user)
 Dim l_rs
 Dim l_sql
 	
 	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = 		"SELECT user_per.ctabloqueada "
	l_sql = l_sql & "FROM user_per "
	l_sql = l_sql & "WHERE user_per.iduser = '" & user & "'"
	rsOpen l_rs, cn, l_sql, 0
	ctabloqueada = false
	if CInt(l_rs("ctabloqueada")) = -1 then
		ctabloqueada = true
	end if
	l_rs.close
	Set l_rs = nothing
end function
 
 
'------------------------------------------------------------------------------------------------- 
' Funcion que devuelve el valor del campo de la politica de cuenta.
'------------------------------------------------------------------------------------------------- 
function valorpol_cuenta (polnro, campo)
 Dim l_rs
 Dim l_sql
 	
 	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = 		"SELECT pol_cuenta." & campo & " "
	l_sql = l_sql & "FROM pol_cuenta "
	l_sql = l_sql & "WHERE pol_cuenta.pol_nro = " & polnro
	rsOpen l_rs, cn, l_sql, 0
	valorpol_cuenta = 0
	if not l_rs.eof then
		valorpol_cuenta = l_rs(0)
	end if
	l_rs.close
	Set l_rs = nothing
end function
 
 
'------------------------------------------------------------------------------------------------- 
' Funcion que devuelve el valor de la politica de cuenta definida del usuario.
'------------------------------------------------------------------------------------------------- 
function valoruser_pol_cuenta(user, campo)
 Dim l_rs
 Dim l_sql
 	
 	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = 		"SELECT usr_pol_cuenta." & campo & " "
	l_sql = l_sql & "FROM usr_pol_cuenta "
	l_sql = l_sql & "WHERE usr_pol_cuenta.iduser = '" & user & "'" 
 	l_sql = l_sql & " AND usr_pol_cuenta.upcfecfin IS NULL"
	rsOpen l_rs, cn, l_sql, 0
'	response.write l_sql & "<br>"
	valoruser_pol_cuenta = 0
	if not l_rs.eof then
		valoruser_pol_cuenta = l_rs(0)
	end if
	l_rs.close
	Set l_rs = nothing
end function
 
 
'------------------------------------------------------------------------------------------------- 
' Funcion que devuelve el valor del campo del historico del logueo del usuario.
'------------------------------------------------------------------------------------------------- 
function valorhistlog (user, campo)
 Dim l_rs
 Dim l_sql
 	
 	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = 		"SELECT hist_log_usr." & campo & " "
	l_sql = l_sql & "FROM hist_log_usr "
	l_sql = l_sql & "WHERE hist_log_usr.iduser = '" & user & "'"
	rsOpen l_rs, cn, l_sql, 0
	valorhistlog = ""
	if not l_rs.eof then
		valorhistlog = l_rs(0)
	end if
	l_rs.close
	Set l_rs = nothing
end function
 
 
'------------------------------------------------------------------------------------------------- 
' Funcion que devuelve el valor del campo del historico de contraseñas del usuario.
'------------------------------------------------------------------------------------------------- 
function valorhistpass (user, campo)
 Dim l_rs
 Dim l_sql
 	
 	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = 		"SELECT hist_pass_usr." & campo & " "
	l_sql = l_sql & "FROM hist_pass_usr "
	l_sql = l_sql & "WHERE hist_pass_usr.iduser = '" & user & "' AND hpassfecfin IS NULL"
	rsOpen l_rs, cn, l_sql, 0
	valorhistpass = ""
	if not l_rs.eof then
		valorhistpass = l_rs(0)
	end if
	l_rs.close
	Set l_rs = nothing
end function
 
 
'------------------------------------------------------------------------------------------------- 
' Funcion que devuelve el valor del campo de la tabla que mantiene a los usuarios (user_per).
'------------------------------------------------------------------------------------------------- 
function valoruser_per (user, campo)
 Dim l_rs
 Dim l_sql
 	
 	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = 		"SELECT user_per." & campo & " "
	l_sql = l_sql & "FROM user_per "
	l_sql = l_sql & "WHERE user_per.iduser = '" & user & "'"
	rsOpen l_rs, cn, l_sql, 0
	valoruser_per = ""
	if not l_rs.eof then
		valoruser_per = l_rs(0)
	end if
	l_rs.close
	Set l_rs = nothing
end function
 
 
'------------------------------------------------------------------------------------------------- 
' Funcion que devuelve la cantidad de logueos fallidos del usuario.
'------------------------------------------------------------------------------------------------- 
function logueosfallidos (user)
 Dim l_rs
 Dim l_sql
 	
 	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT hlogintfallidos FROM hist_log_usr WHERE iduser = '" & user & "'"
	rsOpen l_rs, cn, l_sql, 0
	logueosfallidos = 0
	if not l_rs.eof then
		logueosfallidos = l_rs(0)
	end if
	l_rs.close
	Set l_rs = nothing
end function
 
 
'------------------------------------------------------------------------------------------------- 
' Procedimiento que actualiza la cantidad de logueos fallidos.
'------------------------------------------------------------------------------------------------- 
Sub actlogfallidos (user, cant)
 Dim l_cm
 Dim l_sql
 	Set l_cm = Server.CreateObject("ADODB.Command")
	l_cm.activeconnection = cn
	
	l_sql = 		"UPDATE hist_log_usr SET hlogintfallidos = " & CInt(cant) & " "
	l_sql = l_sql & "WHERE iduser = '" & user & "'"
	cmExecute l_cm, l_sql, 0
	
	Set l_cm = nothing
end sub
 
 
'------------------------------------------------------------------------------------------------- 
' Procedimiento que bloquea o desbloquea la cuenta de un usuario.
'------------------------------------------------------------------------------------------------- 
Sub bloquearcuenta (user, valor)
 Dim l_cm
 Dim l_sql
 	Set l_cm = Server.CreateObject("ADODB.Command")
	l_cm.activeconnection = cn
	
	l_sql = 		"UPDATE user_per SET ctabloqueada = " & CInt(valor) & " "
	l_sql = l_sql & "WHERE iduser = '" & user & "'"
	cmExecute l_cm, l_sql, 0
	
	Set l_cm = nothing
end sub
 
 
'------------------------------------------------------------------------------------------------- 
' Procedimiento que ingresa el logueo del usuario en la Base de Datos.
'------------------------------------------------------------------------------------------------- 
Sub ingresarlogueo (user)
 Dim l_cm
 Dim l_rs
 Dim l_sql
 Dim l_fecha
 Dim l_hora
 	
 	Set l_rs = Server.CreateObject("ADODB.RecordSet")
 	Set l_cm = Server.CreateObject("ADODB.Command")
	l_cm.activeconnection = cn
 	
	l_fecha = date()
	l_hora 	= hour(now()) & ":" & minute(now()) & ":" & second(now())
	
	l_sql = "SELECT hlognro FROM hist_log_usr WHERE iduser = '" & user & "'"
	rsOpen l_rs, cn, l_sql, 0
	if l_rs.eof then
		l_sql = 		"INSERT INTO hist_log_usr (iduser, hlogfecini, hloghoraini, hlogintfallidos)"
		l_sql = l_sql & " VALUES ('" & user & "', " & cambiafecha(l_fecha, "", "") & ", '" & l_hora & "', 0)"
	else
		l_sql = 		"UPDATE hist_log_usr SET hlogfecini = " & cambiafecha(l_fecha, "", "") & ", hloghoraini = '" & l_hora & "', hlogintfallidos = 0"
		l_sql = l_sql & " WHERE iduser = '" & user & "'"
	end if
	l_rs.close
	
	cmExecute l_cm, l_sql, 0
	
	Set l_cm = nothing
	Set l_rs = nothing
end sub
 
 
'------------------------------------------------------------------------------------------------- 
' Procedimiento que da de baja un password en la Base de Datos.
'------------------------------------------------------------------------------------------------- 
Sub bajapass (user)
 Dim l_cm
 Dim l_rs
 Dim l_sql
 Dim l_fecha
 Dim l_hora
 	
 	Set l_rs = Server.CreateObject("ADODB.RecordSet")
 	Set l_cm = Server.CreateObject("ADODB.Command")
	l_cm.activeconnection = cn
 	
	l_fecha = date()
	l_hora 	= hour(now()) & ":" & minute(now()) & ":" & second(now())
	
	l_sql = "SELECT hpassnro FROM hist_pass_usr WHERE iduser = '" & user & "' AND hpassfecfin IS NULL"
	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof then
		l_sql = 		"UPDATE hist_pass_usr SET hpassfecfin = " & cambiafecha(l_fecha, "", "") & ", hpasshorafin = '" & l_hora & "'"
		l_sql = l_sql & " WHERE hpassnro = " & l_rs("hpassnro")
		cmExecute l_cm, l_sql, 0
	end if
	l_rs.close
	
	Set l_cm = nothing
	Set l_rs = nothing
end sub
 
 
'------------------------------------------------------------------------------------------------- 
' Procedimiento que ingresa un nuevo password en la Base de Datos.
'------------------------------------------------------------------------------------------------- 
Sub ingresarpass (user, pass)
 Dim l_cm
 Dim l_sql
 Dim l_fecha
 Dim l_hora
 	
 	Set l_cm = Server.CreateObject("ADODB.Command")
	l_cm.activeconnection = cn
 	
	l_fecha = date()
	l_hora 	= hour(now()) & ":" & minute(now()) & ":" & second(now())
	
	'Guardo la clave en la tabla de password
	l_sql = 		"INSERT INTO hist_pass_usr (iduser, husrpass, hpassfecini, hpasshoraini)"
	l_sql = l_sql & " VALUES ('" & user & "','" & pass & "'," & cambiafecha(l_fecha, "", "") & ", '" & l_hora & "')"
	cmExecute l_cm, l_sql, 0
	
	'Cambio la clave del usuario de BD
	if Trim(ucase(user)) = Trim(ucase(session("UserName"))) then
		'Eugenio 22/07/2015. Elimino las ejecuciones de store procedures de la base para manipulacion de usuarios porque en datatec no tienen permisos cn.Execute "exec sp_password '" & session("Password") & "','" & decrypt(c_strEncryptionKey,pass, false) & "'", , 1
		session("password") = decrypt(c_strEncryptionKey,pass, false) 
	ELSE
		'Eugenio 22/07/2015. Elimino las ejecuciones de store procedures de la base para manipulacion de usuarios porque en datatec no tienen permisos cn.Execute "exec sp_password null, '" & decrypt(c_strEncryptionKey,pass, false) & "', '" & user & "'", , 1
	end if
	Set l_cm = nothing
end sub
 
'------------------------------------------------------------------------------------------------- 
' Procedimiento que cambiar un nuevo password en la Base de Datos.
'------------------------------------------------------------------------------------------------- 
function CambiarPass (user, pass)
	Dim l_cm
	Dim l_sql
	Dim l_fecha
	Dim l_hora
 	
 	Set l_cm = Server.CreateObject("ADODB.Command")
	l_cm.activeconnection = cn
 	
	l_fecha = date()
	l_hora 	= hour(now()) & ":" & minute(now()) & ":" & second(now())
	
'	'Guardo la clave en la tabla de password
'	l_sql = 		"INSERT INTO hist_pass_usr (iduser, husrpass, hpassfecini, hpasshoraini)"
'	l_sql = l_sql & " VALUES ('" & user & "','" & pass & "'," & cambiafecha(l_fecha, "", "") & ", '" & l_hora & "')"
'	cmExecute l_cm, l_sql, 0
	'Cambio la clave del usuario de BD
	if Trim(ucase(user)) = Trim(ucase(session("UserName"))) then
		dim l_pass
		if session("Password") = "" then
			l_pass = "null"
		else 
			l_pass = "'" & session("Password") & "'"
		end if
		'response.write ("exec sp_password " & l_pass & ",'" & decrypt(c_strEncryptionKey,pass, false) & "'")
		'response.end
		'Eugenio 22/07/2015. Elimino las ejecuciones de store procedures de la base para manipulacion de usuarios porque en datatec no tienen permisos cn.Execute "exec sp_password " & l_pass & ",'" & decrypt(c_strEncryptionKey,pass, false) & "'", , 1
		session("password") = decrypt(c_strEncryptionKey,pass, false) 
		CambiarPass = true
	ELSE
		on error resume next
		err.clear
		'Eugenio 22/07/2015. Elimino las ejecuciones de store procedures de la base para manipulacion de usuarios porque en datatec no tienen permisos cn.Execute "exec sp_password null, '" & decrypt(c_strEncryptionKey,pass, false) & "', '" & user & "'", , 1
		if err = 0 then
			CambiarPass = true
		else
			err.clear
			CambiarPass = false
		end if
		on error goto 0
	end if
	if CambiarPass = true then
		'Guardo la clave en la tabla de password
		l_sql = " UPDATE hist_pass_usr SET husrpass = '" & pass & "'"
		' 07/09/2016 Raul Agregado para que funcione el cambio de pass
		l_sql = l_sql & " , hpassfecfin = NULL"
		l_sql = l_sql & " , hpasshorafin = ''"
		l_sql = l_sql & " WHERE iduser = '" & Trim(ucase(user)) & "'"
		cmExecute l_cm, l_sql, 0
	end if
	Set l_cm = nothing
end function
 
'------------------------------------------------------------------------------------------------- 
' Procedimiento que mantiene cant de password en el historico, eliminando todos los demas.
'------------------------------------------------------------------------------------------------- 
Sub eliminarhistpass (user, cant)
 Dim l_cm
 Dim l_rs
 Dim l_sql
 Dim i
 	
 	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = 		"SELECT hist_pass_usr.hpassnro "
'	l_sql = l_sql & ",DATEPART(yyyy,hpassfecini) AS ordanio, DATEPART(mm,hpassfecini) AS ordmes, "
'	l_sql = l_sql & "DATEPART(dd,hpassfecini) AS orddia, DATEPART(hh,hpasshoraini) AS ordhora, "
'	l_sql = l_sql & "DATEPART(mi,hpasshoraini) AS ordminutos, DATEPART(ss,hpasshoraini) AS ordseg "
	l_sql = l_sql & "FROM hist_pass_usr "
	l_sql = l_sql & "WHERE hist_pass_usr.iduser = '" & user & "' "
	l_sql = l_sql & "ORDER BY hpassfecini DESC, hpasshoraini DESC"
'	, ordmes DESC, orddia DESC, ordhora DESC, ordminutos DESC, ordseg DESC"
	rsOpenCursor l_rs, cn, l_sql, 0, adOpenKeyset
	
	i = 0
 	Set l_cm = Server.CreateObject("ADODB.Command")
	l_cm.activeconnection = cn
	do until l_rs.eof
		i = i + 1
		
		if i > cant then
			l_sql = "DELETE FROM hist_pass_usr WHERE hpassnro = " & l_rs("hpassnro")
			cmExecute l_cm, l_sql, 0
		end if
		
		l_rs.MoveNext
	loop
	l_rs.close
	
	Set l_cm = nothing
	Set l_rs = nothing
	
end sub
 
 
'------------------------------------------------------------------------------------------------- 
' Funcion que valida si se esta repitiendo el password en base al historico.
'------------------------------------------------------------------------------------------------- 
function passrepetido(user, cant, usrpass)
 Dim l_rs
 Dim l_sql
 Dim i
 	
 	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = 		"SELECT hist_pass_usr.husrpass, hpassfecfin "
'	l_sql = l_sql & ",DATEPART(yyyy,hpassfecini) AS ordanio, DATEPART(mm,hpassfecini) AS ordmes, "
'	l_sql = l_sql & "DATEPART(dd,hpassfecini) AS orddia, DATEPART(hh,hpasshoraini) AS ordhora, "
'	l_sql = l_sql & "DATEPART(mi,hpasshoraini) AS ordminutos, DATEPART(ss,hpasshoraini) AS ordseg "
	l_sql = l_sql & "FROM hist_pass_usr "
	l_sql = l_sql & "WHERE hist_pass_usr.iduser = '" & user & "' "
	l_sql = l_sql & "ORDER BY hpassfecini DESC, hpasshoraini DESC"
'	l_sql = l_sql & "ORDER BY ordanio DESC, ordmes DESC, orddia DESC, ordhora DESC, ordminutos DESC, ordseg DESC"
	rsOpen l_rs, cn, l_sql, 0
	
	passrepetido = false
	i = 1
	do while (i<=cant and not l_rs.eof)
		if l_rs("husrpass") = usrpass then
			i = cant + 1
			passrepetido = true
		end if
		if l_rs("hpassfecfin") <> "" then
			i = i + 1
		end if
		l_rs.movenext
	loop
	l_rs.close
	Set l_rs = nothing
end function
 
 
'------------------------------------------------------------------------------------------------- 
' Sub que crea un usuario en la bd
'------------------------------------------------------------------------------------------------- 
function crearUsuario(user,pass) 'Eugenio 22/07/2015. Elimino las ejecuciones de store procedures de la base para manipulacion de usuarios porque en datatec no tienen permisos
'	Dim cantidad
'	Dim l_cm
'	Dim l_rs
'	Dim l_encontrado
'	
'	'Busco al usuario en la bd
'	cn.execute "exec sp_helplogins '" & user & "'",cantidad
'	if cantidad = 0 then
'		'Si no encuentro al usuario lo agrego
'		cn.Execute "exec sp_addlogin '" & user & "','" & pass & "','turnos'", , 1
'		cn.Execute "exec sp_grantdbaccess '" & user & "','" & user & "'", , 1
'		cn.Execute "exec sp_addrolemember 'db_owner','" & user & "'", , 1
'	else
'	  	Set l_rs = Server.CreateObject("ADODB.RecordSet")
'		Set l_cm = Server.CreateObject("ADODB.Command")
'		l_cm.activeconnection = cn
'
'		'Lista de usuarios de la BD actual
'		l_cm.CommandText = "sp_helpuser"
'		Set l_rs = l_cm.Execute(cantidad, , 4)
'		l_encontrado = false
'
'		do while not l_rs.eof
'			if user = trim(l_rs(0)) then
'				l_encontrado = true
'			end if
'			l_rs.MoveNext
'		loop
'
'		'Si el usuario no figura entre los usuarios de la bd lo agrego
'		if not l_encontrado then
'			on error resume next
'			cn.Execute "exec sp_grantdbaccess '" & user & "','" & user & "'", , 1
'			on error goto 0
'		end if
'	end if
'	
'	if err then
'		crearUsuario = false
'	else
'		crearUsuario = true
'	end if
'	on error goto 0
	crearUsuario = true
end function
%>