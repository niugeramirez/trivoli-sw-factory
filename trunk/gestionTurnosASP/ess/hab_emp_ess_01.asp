<%Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/encrypt.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : hab_emp_ess_01.asp
Descripcion    : Pagina utilizada para habilitar al empleado a entrar a Autogestion
Creador        : GdeCos
Fecha Creacion : 14/4/2005
modificado	   : 21/09/2006 - Martin Ferraro - Se agrego actualizacion del password
-----------------------------------------------------------------------------
--><% 
on error goto 0

' Variables
Dim l_ternro
Dim l_estado
Dim l_perfnro
Dim l_passnew

dim l_rs
dim l_sql
Dim l_cm

l_estado  = Request.Form("estado")
l_ternro  = Request.Form("ternro")
l_perfnro = Request.Form("perfnro")
l_passnew = Request.Form("newpass")

if Cint(l_estado) > 0 then l_estado ="-1" end if

set l_cm = Server.CreateObject("ADODB.Command")
		
		l_sql = "UPDATE empleado SET empessactivo = " & l_estado & ",perfnro = " & l_perfnro
		'Modificar tambien el pass
		if l_passnew <> "" then
			' Encriptarlo y Grabarlo en la BD
			l_passnew = Encrypt(c_seed1,l_passnew)
			l_sql = l_sql & " ,emppass = '" & l_passnew & "'"
		end if
		l_sql = l_sql & " WHERE ternro = " & l_ternro
		
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0

Set l_cm = Nothing

cn.close

set cn = nothing
%>

<script>
	alert('Operacion Realizada.');
    parent.window.close();
    //window.close();
</script>


