 <% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/inc/encrypt.inc"-->
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/password.inc"-->
<!--------------------------------------------------------------------------------------------
Archivo     : cambio_clave_seg_02.asp
Descripcion : Valida datos.
Autor       : Favre F.
Creacion    : 30/07/2004
Modificar	:
----------------------------------------------------------------------------------------------
-->
<%
'on error goto 0
 Dim l_cm
 Dim l_sql
 
 Dim l_tipo
 Dim l_iduser
 Dim l_usrpassnuevo
 Dim l_pass_historia
 Dim l_usrpasscambiar
 
 Dim l_usrpassencryp
 dIM l_pol_nro
' Dim l_pass_int_fallidos
 
 l_tipo			= request.QueryString("tipo")
 l_iduser 		= request.Form("iduser")
 l_usrpassnuevo	= request.Form("usrpassnuevo")
 
 l_iduser  		 = lcase(l_iduser)
 l_usrpassencryp = Decrypt(c_strEncryptionKey, l_usrpassnuevo)
 
 
 if l_tipo = "M" then 
 	cn.beginTrans
	
	' Blanqueo la cantidad de intentos fallidos
	actlogfallidos l_iduser, 0
	
	' Doy de baja el viejo password
	bajapass l_iduser
	
	' Verifico el historico de password, para mantener la cantidad de pass historicos definida en la politica de cuenta
	l_pol_nro = valoruser_pol_cuenta(l_iduser ,"pol_nro")
	l_pass_historia = valorpol_cuenta (l_pol_nro, "pass_historia")
	eliminarhistpass l_iduser, l_pass_historia
	
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
	
	cn.commitTrans
 end if
 
 Response.write "<script>alert('Operación Realizada.');window.close();</script>"
%>
