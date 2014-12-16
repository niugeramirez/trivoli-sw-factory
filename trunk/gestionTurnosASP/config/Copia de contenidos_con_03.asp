<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: berths_con_03.asp
'Descripción: ABM de Berths
'Autor : Raul Chinestra
'Fecha: 23/11/2007

Dim l_tipo
Dim l_cm
Dim l_sql

Dim l_connro
Dim l_mernro
Dim l_buqnro
Dim l_expnro
Dim l_sitnro
Dim l_conton
Dim l_desnro

l_tipo 		= request.querystring("tipo")
l_connro 	= request.Form("connro")
l_mernro	= request.Form("mernro")
l_buqnro 	= request.Form("buqnro")
l_expnro 	= request.Form("expnro")
l_sitnro 	= request.Form("sitnro")
l_conton 	= request.Form("conton")
l_desnro 	= request.Form("desnro")

set l_cm = Server.CreateObject("ADODB.Command")
if l_tipo = "A" then 
	l_sql = "INSERT INTO buq_contenido"
	l_sql = l_sql & " (buqnro, mernro, expnro, sitnro, conton, desnro)"
	l_sql = l_sql & " VALUES (" & l_buqnro & "," & l_mernro & "," & l_expnro & "," & l_sitnro & "," & l_conton & "," & l_desnro & " )"
else
	l_sql = "UPDATE buq_contenido "
	l_sql = l_sql & " SET buqnro = " & l_buqnro
	l_sql = l_sql & "    ,mernro = " & l_mernro 
	l_sql = l_sql & "    ,expnro = " & l_expnro 
	l_sql = l_sql & "    ,sitnro = " & l_sitnro
	l_sql = l_sql & "    ,conton = " & l_conton
	l_sql = l_sql & "    ,desnro = " & l_desnro
	l_sql = l_sql & " WHERE connro = " & l_connro
end if
'response.write l_sql & "<br>"
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
%>

