<%  Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : conf_x_empr_03.asp
Descripcion    : Modulo que se encarga de las Altas/Modificaciones de conf de
                 empresas.
Creador        : Scarpa D.
Fecha Creacion : 21/08/2003
-----------------------------------------------------------------------------
-->
<% 
' Declaracion de Variables locales  -------------------------------------
Dim l_tipo
Dim l_cm
Dim l_sql
dim l_rs

Dim l_confnro

Dim	l_confdesc
Dim	l_confint 
Dim	l_confactivo 

Dim	l_confdescant
Dim	l_confintant 
Dim	l_confactivoant 

' traer valores del form de alta/modificacion -------------------------------------

l_tipo = request("tipo")

l_confnro	  = Request.form("confnro")

l_confdesc    = Request.form("confdesc")
l_confint     = Request.form("confint")
l_confactivo  = Request.form("confactivo")

l_confdescant   = Request.form("confdescant")
l_confintant    = Request.form("confintant")
l_confactivoant = Request.form("confactivoant")

if l_confactivo = "on" then
	l_confactivo = -1
else
	l_confactivo = 0
end if

set l_cm = Server.CreateObject("ADODB.Command")
if l_tipo = "A" then 

	  l_sql = "INSERT INTO confper "
	  l_sql = l_sql & "(confdesc , confint, confactivo, empnro) "
	  l_sql = l_sql & " VALUES ('" 
	  l_sql = l_sql & l_confdesc   & "'," 
	  l_sql = l_sql & l_confint  & "," 
	  l_sql = l_sql & l_confactivo  & "," 
	  l_sql = l_sql & "1)"
	  l_cm.activeconnection = Cn
	  l_cm.CommandText = l_sql
	  cmExecute l_cm, l_sql, 0	

else
		if (l_confdescant <> l_confdesc) or _  
		   (l_confintant   <> l_confint  ) or _
		   (l_confactivoant   <> l_confactivo  ) then
		   ' algo modifico asi que me fijo si existe otrto reg. con lo nuevo

			l_sql = "UPDATE confper SET "
			l_sql = l_sql & " confdesc		= '" & l_confdesc   & "', " 
			l_sql = l_sql & " confint		= "  & l_confint    & ", " 
			l_sql = l_sql & " confactivo    = "  & l_confactivo 
			l_sql = l_sql & " WHERE confnro = "  & l_confnro
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			'response.write l_sql
			cmExecute l_cm, l_sql, 0	
		end if ' modifico algo	
end if

Response.write "<script>alert('Operación Realizada.');window.opener.opener.ifrm.location.reload();window.opener.close();window.close();</script>"
%>
