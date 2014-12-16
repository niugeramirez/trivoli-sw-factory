<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->

<!--
-----------------------------------------------------------------------------
Archivo        : closesession_ess.asp
Descripcion    : Pagina designada para limpiar las variables de la conexion con la BD
					y llamar nuevamente a la pagina de login
Creador        : GdeCos
Fecha Creacion : 1/4/2005
-----------------------------------------------------------------------------
-->

<% 
	Session("UserName") = ""
	Session("Password") = ""
	Session("Time") = ""
	Session("empleg") = ""
	Session.Abandon
	cn.close	
	
	Response.write "<script> 	window.location = 'index.asp';</script>"
%>