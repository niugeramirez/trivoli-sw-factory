<!--#include virtual="/turnos/shared/inc/encrypt.inc"-->
<% 
Sub cmExecute_(ByRef cm, ByVal sql, ByVal nivel)
	Dim Buffer
	nivel = 1 
	on error resume next
	cm.commandText = sql
	cm.Execute
	if err then
		Buffer = "<script>"
		if debug then 
			Buffer = Buffer & "alert(" & chr(34) & "Debug: TRUE\n"
			Buffer = Buffer & "Archivo: " & Request.ServerVariables("SCRIPT_NAME") & "\n"
			Buffer = Buffer & "Numero Error: " & err.number & "\n"
			Buffer = Buffer & "Descripcion: " & replace(err.description,chr(34),"´") & "\n"
			Buffer = Buffer & "SQL: " & replace(sql,chr(34),"´") & chr(34) & ");"
			'Buffer = Buffer & "prompt('SQL String:',"&chr(34)&replace(sql,chr(34),"´")&chr(34)&");"
		else
			Buffer = Buffer & "alert('" & replace(err.description,chr(34),"´") & "');"
		end if
		if nivel = 1 then
			Buffer = Buffer & "window.close();"
		else
			if nivel = 2 then
				Buffer = Buffer & "window.opener.close();window.close();"
			end if
		end if
		Buffer = Buffer & "</script>"
		response.write Buffer
		response.end
	end if
End Sub

function crear_usuario(id, clave)
	Dim cm3
	Dim l_username
	Dim l_password
	l_username = id
	l_password = ucase(clave)
	set cm3 = Server.CreateObject("ADODB.Command")
	cm3.ActiveConnection = cn
	
	'Crea el login al servidor de la bd si es que no existe
	cm3.CommandText = "if not exists (select * from master.dbo.syslogins where loginname = N'" & l_username & "') " &_
							"BEGIN " &_
							"declare @logindb nvarchar(128) " &_
							"select @logindb = db_name(dbid) from master..sysprocesses where spid=@@SPID " &_
							"exec sp_addlogin N'" & l_username & "', '" & l_password & "', @logindb, @@language " &_
							"END"
	cmExecute_ cm3, cm3.CommandText, 0
	
	'Crea el mapeo entre el login y el usuario de la bd si es que no existe
	cm3.CommandText = "if not exists (select * from dbo.sysusers where name = N'" & l_username & "' and uid < 16382) " &_
							"BEGIN " &_
							"   exec sp_grantdbaccess N'" & l_username & "', N'" & l_username & "' " &_
							"   exec sp_addrolemember N'rhprox2', N'" & l_username & "' " &_
							"END"
							'"   exec sp_addrolemember N'db_owner', N'" & l_username & "' " &_
	
	cmExecute_ cm3, cm3.CommandText, 0
	
	Set cm3 = Nothing
end function

function cambiar_password(id, clave)
	Dim cm3
	Dim l_username
	Dim l_password
	l_username = id
	l_password = ucase(clave)
	set cm3 = Server.CreateObject("ADODB.Command")
	cm3.ActiveConnection = cn
	if ucase(Session("username")) = Ucase(l_username) then
		cm3.CommandText = "EXEC sp_password '" & Session("password") & "', '" & l_password & "'"
	else
		cm3.CommandText = "EXEC sp_password NULL, '" & l_password & "', '" & l_username & "'"
	end if
	cmExecute_ cm3, cm3.CommandText, 0
	on error resume next
	if (not err) and (ucase(Session("username")) = Ucase(l_username)) then Session("password") = l_password
	on error goto 0
	Set cm3 = Nothing
end function

function borrar_usuario(id)
	Dim cm3
	Dim l_username
	l_username = id
	set cm3 = Server.CreateObject("ADODB.Command")
	cm3.ActiveConnection = cn
	cm3.CommandText = "if exists (select * from master.dbo.syslogins where loginname = N'" & l_username & "') " &_
							" BEGIN " &_
							" EXEC sp_dropuser '" & l_username & "' " &_
							" EXEC sp_droplogin '" & l_username & "' " &_
							" END"
	cmExecute_ cm3, cm3.CommandText, 0
	Set cm3 = Nothing
end function
%>

