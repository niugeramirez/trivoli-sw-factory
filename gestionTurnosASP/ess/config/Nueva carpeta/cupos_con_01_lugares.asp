<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/adovbs.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/asistente.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--
Archivo: cupos_con_01.asp 
Descripción: Asignacion de Lugares que se desea bajar los Cupos
Autor : Raul CHinestra	
Fecha: 10/01/2006
-->
<% 
'on error goto 0
'ADO
 Dim l_cm
 Dim l_sql
 Dim l_rs
 
' variables
 Dim l_grabar 
 dim i
 dim l_lista

' parametros de entrada
 l_grabar   = Request.QueryString("grabar")
 
' -------------------------------------------------------------------
' BODY --------------------------------------------------------------
' -------------------------------------------------------------------
set l_cm = Server.CreateObject("ADODB.Command")

l_sql = " UPDATE tkt_config SET concup = '" & l_grabar & "'"
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0

 Set cn = Nothing
 Set l_cm = Nothing
 
 response.write "<script>alert('Operación Realizada.');window.opener.close();window.close();</script>"
%>
