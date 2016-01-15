<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/const.inc"-->
<% 
Dim l_cm
Dim l_sql
Dim l_mulnro

l_mulnro = request.QueryString("mulnro")

set l_cm = Server.CreateObject("ADODB.Command")
l_sql = "DELETE FROM multiempresa " 
l_sql = l_sql & "WHERE mulnro = " & l_mulnro
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0

cn.Close
Set cn = Nothing
Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location = 'multiempresa_gti_01.asp';window.close();</script>"
%>
