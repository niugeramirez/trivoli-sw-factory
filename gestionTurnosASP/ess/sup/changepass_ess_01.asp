<%Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/encrypt.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : changepass_ess_01.asp
Descripcion    : Pagina encargada de grabar el nuevo password del usuario
Creador        : GdeCos
Fecha Creacion : 1/4/2005
-----------------------------------------------------------------------------
-->
<% 
on error goto 0

' Variables
Dim l_empleg
Dim l_pass
Dim l_passold
Dim l_passnew
Dim l_es_MSS

dim l_rs
dim l_sql
Dim l_cm

l_empleg = l_ess_empleg
l_passold = Request.Form("passold")
l_passnew = Request.Form("passnew")

l_es_MSS = (CStr(Session("empleg")) <> CStr(l_ess_empleg))

set l_cm = Server.CreateObject("ADODB.Command")
Set l_rs = Server.CreateObject("ADODB.RecordSet")	

    if l_es_MSS then
		' Obtengo el password encriptado de la BD
		l_sql = "SELECT emppass FROM empleado "
		l_sql = l_sql & " WHERE empleado.empleg = " & l_empleg 
		l_rs.MaxRecords = 1
		rsOpen l_rs, cn, l_sql, 0 
		
		if not l_rs.eof then
			l_pass  = Decrypt(c_seed1,l_rs("emppass"))
		else
			' Error no se pudo obtener el password
			Response.write "<script>parent.window.alert('Error. No se puede recuperar la contraseña del Usuario.');</script>"
			Response.end
		end if
		
	    l_rs.Close
	else
    	l_pass    = ""
	    l_passold = l_pass
	end if
	
	set l_rs = nothing

	if (l_pass <> l_passold) then
		' Password Incorrecto
		Response.write "<script>"
		Response.write "	parent.window.alert('Contraseña Actual Incorrecta.');"
		Response.write "	parent.window.location = 'changepass_ess_00.asp';"
		Response.write "</script>"
		Response.end
	else
		' Encriptarlo y Grabarlo en la BD
		l_passnew = Encrypt(c_seed1,l_passnew)
		
		l_sql = "UPDATE empleado SET emppass = '" & l_passnew
		l_sql = l_sql & "' WHERE empleado.empleg = " & l_empleg
	
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0

		Set l_cm = Nothing

	end if
	
	' Termino. Aviso que se cambio el password satisfactoriamente
	Response.write "<script>"
	Response.write "	parent.window.alert('Se ha cambiado su contraseña.\nPara seguir utilizando el sistema, ingrese con su nueva contraseña.');"
	Response.write "	parent.parent.parent.window.location = '../closesession.asp';"
	Response.write "</script>"
	Response.end

%>

