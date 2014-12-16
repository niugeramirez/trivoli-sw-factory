<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo: columnas_reporte_04.asp
Descripcion: Modulo que se encarga de realizar la baja de una columna
             de un reporte con confrep.
Modificacion:
    29/07/2003 - Scarpa D. - Agregado de la columna confsuma   
    25/08/2003 - Scarpa D. - Agregado de la columna confval2		
-----------------------------------------------------------------------------
-->
<% 
Dim l_cm
Dim l_sql

Dim l_repnro
Dim l_confnrocol
Dim l_conftipo
Dim l_confetiq
Dim l_confval
Dim l_confval2
Dim l_confaccion

l_repnro		= Request.QueryString("repnro")
l_confnrocol 	= Request.QueryString("confnrocol")
l_conftipo		= Request.QueryString("conftipo")
l_confetiq		= Request.QueryString("confetiq")
l_confval		= Request.QueryString("confval")
l_confval2		= Request.QueryString("confval2")
l_confaccion	= Request.QueryString("confaccion")

set l_cm = Server.CreateObject("ADODB.Command")
l_sql = "DELETE FROM confrep WHERE repnro = " & l_repnro
l_sql = l_sql & " AND confnrocol = " & l_confnrocol
l_sql = l_sql & " AND confetiq = '" & l_confetiq & "'"
l_sql = l_sql & " AND conftipo = '" & l_conftipo & "'"
l_sql = l_sql & " AND confval = " & l_confval
l_sql = l_sql & " AND confval2 = '" & l_confval2 & "'"
l_sql = l_sql & " AND confaccion = '" & l_confaccion & "'"
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0	

Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location = 'columnas_reporte_01.asp?repnro=" & l_repnro & "';window.close();</script>"
%>
