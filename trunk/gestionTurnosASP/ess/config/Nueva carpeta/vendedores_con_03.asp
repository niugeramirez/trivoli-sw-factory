<% Option Explicit %>
<!--#include virtual="/ticket/shared/inc/sec.inc"-->
<!--#include virtual="/ticket/shared/inc/const.inc"-->
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<!--#include virtual="/Ticket/shared/inc/fecha.inc"-->
<% 
'Archivo: vendedores_con_03.asp
'Descripción: Rutina que almacena la Camara asociada al Vendedor
'Autor : Raul Chinestra	
'Fecha: 04/01/2006

on error goto 0

Dim l_tipo
Dim l_cm
Dim l_sql

Dim l_vencornro
Dim l_camnro

l_vencornro 	= request.Form("vencornro")
l_camnro		= request.Form("camnro")

set l_cm = Server.CreateObject("ADODB.Command")

l_sql = "UPDATE tkt_vencor "
l_sql = l_sql & " SET camnro = " & l_camnro
l_sql = l_sql & " WHERE vencornro = " & l_vencornro

'response.write l_sql & "<br>"
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');;window.parent.close();</script>"
%>

