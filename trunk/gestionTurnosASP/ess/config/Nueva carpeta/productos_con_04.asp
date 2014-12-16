<% Option Explicit %>
<!--#include virtual="/ticket/shared/inc/sec.inc"-->
<!--#include virtual="/ticket/shared/inc/const.inc"-->
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<!--#include virtual="/Ticket/shared/inc/fecha.inc"-->
<% 
'Archivo: productos_con_04.asp
'Descripción: Actualizar el campo Acumulador Auxiliar
'Autor : Raul Chinestra	
'Fecha: 11/07/2006

on error goto 0

Dim l_tipo
Dim l_cm
Dim l_sql

Dim l_pronro
Dim l_proacuaux

l_pronro 	= request.Form("pronro")
l_proacuaux	= request.Form("proacuaux")

set l_cm = Server.CreateObject("ADODB.Command")

l_sql = "UPDATE tkt_producto "
l_sql = l_sql & " SET proacuaux = " & l_proacuaux
l_sql = l_sql & " WHERE pronro = " & l_pronro

'response.write l_sql & "<br>"
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.parent.close();</script>"
%>

