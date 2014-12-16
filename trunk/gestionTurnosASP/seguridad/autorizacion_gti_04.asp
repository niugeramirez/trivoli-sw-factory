<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<% 
Dim l_cm
Dim l_sql
Dim l_cystipnro

l_cystipnro = Request("cystipnro")

set l_cm = Server.CreateObject("ADODB.Command")
l_sql = "DELETE FROM cystipo WHERE cystipnro = " & l_cystipnro 
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
l_cm.close
cn.Close
Set cn = Nothing
Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location = 'autorizacion_gti_01.asp';window.close();</script>"
%>
