<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/password.inc"-->
<% 

'Archivo: usuarios_seg_04.asp
'Descripción: ABM de usuarios 
'Autor: Alvaro Bayon
'Fecha: 21/02/2005

on error goto 0

Dim l_cm
Dim l_sql
Dim l_iduser
Dim l_clave

Dim l_rs
Dim l_rs1
Dim l_rs2
Dim l_rs3
Dim l_rs4
Dim l_rs5
Dim l_rs6
Dim l_rs7
Dim l_rs8
Dim l_rs9
Dim l_rs10
Dim l_rs11
Dim l_rs12
Dim l_rs13
Dim l_rs14
Dim l_rs15
Dim l_rs16
Dim l_rs17
Dim l_rs18
Dim l_rs19
Dim l_rs20
Dim l_rs21
	
	l_iduser = request("iduser")
	
	if UCase(l_iduser) <> UCase(Session("username")) then
		
		usuario 2, l_iduser, l_clave
		
		cn.BeginTrans
			
			' Recupero la clave del usuario antes de darlo de baja
			l_clave = valorhistpass (l_iduser, "husrpass")
			
			' Doy de baja la contraseña
			bajapass l_iduser
			
			' Elimino las contraseñas del historico
			eliminarhistpass l_iduser, 0
			
			' Elimino el logueo histórico del usuario
			set l_cm = Server.CreateObject("ADODB.Command")
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			
			l_sql = "DELETE FROM hist_log_usr WHERE iduser = '" & l_iduser & "'"
			cmExecute l_cm, l_sql, 0
			
			' Elimino la política de cuenta del usuario
			l_sql = "DELETE FROM usr_pol_cuenta WHERE iduser = '" & l_iduser & "'"
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0
			
			' Elimino la cuenta del usuario
			l_sql = "DELETE FROM user_per WHERE iduser = '" & l_iduser & "'"
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0
		
		cn.CommitTrans
		
		Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location.reload();window.close();</script>"
	else
		Response.write "<script>alert('No puede eliminar su usuario.');window.close();</script>"
	end if
	
	Set l_cm = Nothing
	cn.Close
	Set cn = Nothing

%>
