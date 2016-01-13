<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 
Dim l_cm
Dim l_rs
Dim l_sql
Dim l_menuaccess
Dim l_menuimg
Dim l_menuorder
Dim l_neworder
Dim l_menuraiz
Dim l_menuname
Dim l_action
Dim l_menudesc
Dim l_tipo
Dim l_parent
Dim j

l_menuaccess = request("menuaccess")
l_menuimg = request("menuimg")
l_menuname = request("menuname")
l_action = request("action")
l_menuorder = request("menuorder")
l_menuraiz = request("menuraiz")
l_tipo = request("tipo")

l_action = Replace(l_action , "'", "''")

Set l_rs = Server.CreateObject("ADODB.RecordSet")
'l_rs.Maxrecords = 1
l_sql = "SELECT * FROM menuraiz where menunro = " & l_menuraiz
rsOpen l_rs, cn, l_sql, 0
l_menudesc =  l_rs("menudesc")
l_parent = l_menuorder & l_menudesc
l_rs.Close

set l_cm = Server.CreateObject("ADODB.Command")

' cn.begintrans  /*NO FUNCIONA CON SQL SERVER

if l_tipo = "HIJO" then
  l_sql = "update menumstr set tipo = 'S' where menuorder = " & l_menuorder & " AND menuraiz = " & l_menuraiz
  l_cm.activeconnection = Cn
  l_cm.CommandText = l_sql
  cmExecute l_cm, l_sql, 0
else
  l_parent = request("parent")
end if

l_menuorder = l_menuorder * 1 + 1

l_sql = "SELECT * FROM menumstr where menuorder >= " & l_menuorder & " AND menuraiz = " & l_menuraiz & " ORDER BY menuorder desc"
rsOpen l_rs, cn, l_sql, 0
do until l_rs.eof 
  l_neworder = l_rs("menuorder") * 1 + 1
  if l_rs("tipo") = "S" then
    l_sql = "update menumstr set parent = '" & l_neworder & l_menudesc & "' where parent = '" & l_rs("menuorder") & l_menudesc & "' AND menuraiz = " & l_menuraiz 
    l_cm.activeconnection = Cn
    l_cm.CommandText = l_sql
    cmExecute l_cm, l_sql, 0
  end if
  l_sql = "update menumstr set menuorder = menuorder + 1 where menuorder = " & l_rs("menuorder") & " AND menuraiz = " & l_menuraiz 
  l_cm.activeconnection = Cn
  l_cm.CommandText = l_sql
  cmExecute l_cm, l_sql, 0

  l_rs.movenext
loop
l_rs.Close

l_sql = "insert into menumstr"
l_sql = l_sql & " (menuaccess, menuimg, menuname, action, menuorder, menuraiz, parent, tipo) "
l_sql = l_sql & " values ('" & l_menuaccess & "','" & l_menuimg & "', '" & l_menuname & "', '" & l_action & "', " & l_menuorder & ", " & l_menuraiz & ", '" & l_parent & "', 'I');"

l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
Dim Buffer
on error resume next
l_cm.Execute
if err then
	Buffer = "<script>"
	if debug then 
		Buffer = Buffer & "alert(" & chr(34) & "Debug: TRUE\n"
		Buffer = Buffer & "Archivo: " & Request.ServerVariables("SCRIPT_NAME") & "\n"
		Buffer = Buffer & "Numero Error: " & err.number & "\n"
		Buffer = Buffer & "Descripcion: " & err.description & "\n"
		Buffer = Buffer & "SQL: " & l_sql & chr(34) & ");"
		Buffer = Buffer & "prompt('SQL String:',"&chr(34)&l_sql&chr(34)&");"
	else
		Buffer = Buffer & "alert('" & err.description & "');"
	end if
	Buffer = Buffer & "</script>"
	response.write Buffer
	response.end
end if

' cn.committrans
Set cn = Nothing
Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');opener.menu2.location='blanc.html';opener.ifrm.location='armado_menu_01.asp?menuraiz=" & l_menuraiz & "';window.close();</script>"
%>
