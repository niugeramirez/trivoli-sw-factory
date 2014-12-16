<% Option Explicit %>
<!--#include virtual="/ticket/shared/inc/sec.inc"-->
<!--#include virtual="/ticket/shared/inc/const.inc"-->
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<%
'Archivo: vendedores inhabilitados_con_03.asp
'Descripción: Modificacion de Vendedores inhabilitados
'Autor : Lisandro Moro
'Fecha: 10/02/2005

Dim l_tipo
Dim l_rs
Dim l_cm
Dim l_sql

Dim l_vencornro
Dim l_venhab
Dim l_vencorfull
Dim l_a

'on error goto 0 
l_vencorfull = request.QueryString("cabnro")
l_tipo	= request.QueryString("tipo")
l_vencornro = split(l_vencorfull,",")

if l_tipo = "H" then
	l_venhab = -1
else
	l_venhab = 0
end if

set l_cm = Server.CreateObject("ADODB.Command")
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
for l_a = 0 to UBound(l_vencornro)
	l_sql = " UPDATE tkt_vencor "
	l_sql = l_sql & " SET venhab = " & l_venhab 
	l_sql = l_sql & " WHERE vencornro = " & l_vencornro(l_a)
	cmExecute l_cm, l_sql, 0
next

Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.parent.ifrm.location.reload();</script>"
%>
