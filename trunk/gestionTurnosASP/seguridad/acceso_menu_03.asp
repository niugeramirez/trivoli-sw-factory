<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
Dim l_cm
Dim l_sql
Dim l_menuaccess
Dim l_menuimg
Dim l_menuorder
Dim l_menuraiz

l_menuaccess = request("menuaccess")
l_menuimg = request("menuimg")
l_menuorder = request("menuorder")
l_menuraiz = request("menuraiz")

set l_cm = Server.CreateObject("ADODB.Command")
l_sql = "update menumstr"
l_sql = l_sql & " set menuaccess = '" & l_menuaccess & "', menuimg = '" & l_menuimg & "'"
l_sql = l_sql & " where menuorder = " & l_menuorder & " AND menuraiz = " & l_menuraiz
response.write l_sql

l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
l_cm.Execute l_sql
'cmExecute l_cm, l_sql, 0
Set cn = Nothing
Set l_cm = Nothing
Response.write "<script>alert('Operación Realizada.');window.close();</script>"
%>
