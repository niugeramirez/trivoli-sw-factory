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
Dim l_menuname
Dim l_parent
Dim l_action
Dim l_nivel
Dim j

l_menuaccess = request("menuaccess")
l_menuimg = request("menuimg")
l_menuname = request("menuname")
l_action = request("action")
l_parent = request("parent")
l_menuorder = request("menuorder")
l_menuraiz = request("menuraiz")
l_nivel = request("nivel")

set l_cm = Server.CreateObject("ADODB.Command")

l_action = Replace(l_action , "'", "''")

l_sql = "update menumstr"
l_sql = l_sql & " set menuaccess = '" & l_menuaccess & "', menuimg = '" & l_menuimg & "', menuname = '" & l_menuname & "', action = '" & l_action & "'"
'l_sql = l_sql & " set menuaccess = '" & l_menuaccess & "', menuimg = '" & l_menuimg & "', menuname = '" & l_menuname & "', action = '" & l_action & "'"
l_sql = l_sql & " where menuorder = " & l_menuorder & " AND menuraiz = " & l_menuraiz

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

Set cn = Nothing
Set l_cm = Nothing

for j = 1 to l_nivel * 1
	l_menuname = "- " & l_menuname
next
Response.write "<script>alert('Operación Realizada.');opener.ifrm.jsSelRow.cells(0).innerText='" & l_menuname & "';window.close();</script>"
%>
