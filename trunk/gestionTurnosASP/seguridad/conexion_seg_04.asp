<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<% 
'Archivo: conexion_seg_00.asp
'Descripción: 
'Autor: Lisandro Moro
'Fecha: 15/03/2005
'Modificado:
on error goto 0

Dim l_cm
Dim l_sql
Dim l_cnnro

l_cnnro = request.querystring("cabnro")

set l_cm = Server.CreateObject("ADODB.Command")
l_sql = "DELETE FROM conexion " 
l_sql = l_sql & "WHERE cnnro = " & l_cnnro
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0

cn.Close
Set l_cm = Nothing
Set cn = Nothing
Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location = 'conexion_seg_01.asp';window.close();</script>"
%>
