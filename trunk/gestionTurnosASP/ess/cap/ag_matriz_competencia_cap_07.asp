<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--
Archivo: ag_matriz_competencia_cap_07.asp
Descripción: 
Autor : Lisandor Moro
Fecha: 29/03/2004
-->

<script src="/serviciolocal/shared/js/fn_fechas.js"></script>

<% 
'on error goto 0 
Dim l_tipo
Dim l_rs
Dim l_cm
Dim l_sql
Dim l_evafacnro
Dim l_fecha
Dim l_porcentaje
Dim l_evanro

 set l_cm = Server.CreateObject("ADODB.Command")

l_evafacnro 	= request.Form("evafacnro")
l_fecha 	    = request.Form("fecha")
l_porcentaje	= request.Form("porcentaje")
l_evanro        = request.Form("evanro")

l_sql = "UPDATE cap_capacita "
l_sql = l_sql & "SET fecha = " & cambiafecha(l_fecha,"YMD", true)
l_sql = l_sql & ",porcen = " & l_porcentaje
l_sql = l_sql & ",entnro = " & l_evanro
l_sql = l_sql & " WHERE entnro = " & l_evafacnro

l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0

Set l_cm = Nothing
Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location.reload();window.close();</script>"
%>
