<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<% 
Dim l_cm
Dim l_rs
Dim l_sql
Dim l_menuorder
Dim l_menuraiz
Dim l_menudesc
Dim l_parent

l_menuorder = Request("menuorder")
l_menuraiz  = Request("menuraiz")
l_parent  = Request("parent")

set l_cm = Server.CreateObject("ADODB.Command")
l_sql = "DELETE FROM menumstr WHERE menuorder = " & l_menuorder & " and menuraiz = " & l_menuraiz
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT * FROM menuraiz where menunro = " & l_menuraiz
rsOpen l_rs, cn, l_sql, 0
l_menudesc =  l_rs("menudesc")
l_rs.close
l_rs.Maxrecords = 1
l_sql = "SELECT * FROM menumstr where parent = '" & l_parent & "'"
rsOpen l_rs, cn, l_sql, 0
if l_rs.eof then
'  Response.write "<script>alert('" & (l_menudesc) & "');</script>"
  l_menuorder = left(l_parent, (len(l_parent) * 1 - len(l_menudesc)))
  l_sql = "update menumstr set tipo = 'I' where menuorder = " & l_menuorder & " AND menuraiz = " & l_menuraiz
  l_cm.activeconnection = Cn
  l_cm.CommandText = l_sql
  cmExecute l_cm, l_sql, 0
end if
l_rs.Close

l_cm.close
cn.Close
Set cn = Nothing
Response.write "<script>alert('Operación Realizada.');window.opener.menu2.location='blanc.html';window.opener.ifrm.location = 'armado_menu_01.asp?menuraiz=" & l_menuraiz &"';window.close();</script>"
%>
