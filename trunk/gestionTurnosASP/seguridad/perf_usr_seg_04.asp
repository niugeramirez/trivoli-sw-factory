<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<% 

'Archivo: perf_usr_seg_04.asp
'Descripción: Abm de Perfiles de usuario
'Autor : Alvaro Bayon
'Fecha: 21/02/2005

Dim l_cm
Dim l_sql
Dim l_perfnro

l_perfnro = request.querystring("cabnro")

set l_cm = Server.CreateObject("ADODB.Command")
l_sql = "DELETE FROM perf_usr " 
l_sql = l_sql & "WHERE perfnro = " & l_perfnro
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0

l_cm.close
cn.Close
Set cn = Nothing
Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location = 'perf_usr_seg_01.asp';window.close();</script>"
%>
