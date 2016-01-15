<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<%
Set rs = Server.CreateObject("ADODB.RecordSet")
Set rs2 = Server.CreateObject("ADODB.RecordSet")
Set rs3 = Server.CreateObject("ADODB.RecordSet")
Set rs4 = Server.CreateObject("ADODB.RecordSet")

Dim l_rs
Dim l_sql
Dim l_username
dim tr
dim l_menuraiz
dim l_menudesc

l_menuraiz = request.QueryString("menuraiz")

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT usrnombre from user_per where upper(iduser) = '" & uCase(Session("Username")) & "'"
l_rs.Maxrecords = 1
rsOpen l_rs, cn, l_sql, 0
l_username = l_rs(0)
l_rs.Close
l_rs = nothing

%>
<html>
<head>
<title><%= Session("Titulo")%>Ticket - Usuario: <%= l_username %></title>
<link href="/trivoliSwimming/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<script>
var jsSelRow = null;

function Deseleccionar(fila)
{
 fila.className = "MouseOutRow";
}
function Seleccionar(fila,cabnro)
{
 if (jsSelRow != null)
 {
  Deseleccionar(jsSelRow);
 };

 document.datos.cabnro.value = cabnro;
 parent.menu2.location = "acceso_menu_02.asp?menuraiz=" + document.datos.menuraiz.value + "&menuorder=" + cabnro;
 fila.className = "SelectedRow";
 jsSelRow		= fila;
}
</script></head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0">

<% 

dim linea, sql, rs, rs2, rs3, rs4, submenu, orden, orden2, padre, destino, perfil, username, menuraiz, f, id

username = UCase(l_Username)

if not username = "SUPER" then
	sql = "SELECT * FROM user_per where upper(user_per.iduser) = '" & username & "'"
	rs3.Maxrecords = 1
	rsOpen rs3, cn, sql, 0
	perfil = rs3("perfnro")
	rs3.Close
end if

sql = "SELECT * FROM menuraiz where menunro = " & l_menuraiz
rsOpen rs4, cn, sql, 0
l_menudesc = rs4("menudesc")
rs4.Close


sql = "SELECT MenuName, MenuOrder, MenuRaiz, Parent, tipo, action, menuaccess, menuimg FROM menumstr where menuraiz = " & l_menuraiz
'if not username = "SUPER" then
'	sql = sql & " AND (menumstr.menuaccess LIKE '%%" & perfil & "%%' OR "
'	sql = sql & " menumstr.menuaccess = '*') "
'end if
sql = sql & " ORDER BY parent, menuorder"
rsOpen rs, cn, sql, 0
menumstr = rs.GetRows
rs.Close

sub imprimir(nodo,nivel)
	for i = 0 to Ubound(menumstr,2)
		if uCase(trim(menumstr(3,i))) = trim(nodo & UCase(l_menudesc)) then
			if trim(menumstr(0,i)) <> "rule" then
    			tr = "<tr onclick='Javascript:Seleccionar(this,"&trim(menumstr(1,i))&")'>"
				response.write tr & "<td>"
				for j = 1 to nivel
					response.write "-&nbsp;"
				next
				response.write menumstr(0,i) & "</td></tr>"
			end if
			if trim(menumstr(4,i)) = "S" then
				imprimir trim(menumstr(1,i)),nivel + 1
			end if
			if trim(menumstr(3,i + 1)) <> trim(menumstr(3,i)) then 
				exit for
			end if
		end if
	next
end sub
%>
<table>
<%
sql = "SELECT MenuName, MenuOrder, MenuRaiz, Parent, tipo, action, menuaccess, menuimg FROM menumstr where menuraiz = " & l_menuraiz
sql = sql & " AND upper(parent) = '" & UCase(l_menudesc) & "'"
'if not username = "SUPER" then
'	sql = sql & " AND (menumstr.menuaccess LIKE '%%" & perfil & "%%' OR "
'	sql = sql & " menumstr.menuaccess = '*') "
'end if
sql = sql & " ORDER BY parent, menuorder"
rsOpen rs, cn, sql, 0
do until rs.eof
    tr = "<tr onclick='Javascript:Seleccionar(this,"&rs(1)&")'>"
	response.write tr & "<td>" & rs(0) & "</td></tr>"
	imprimir rs(1),1
	rs.moveNext
loop
rs.close
%>	
</table>
<form name="datos" method="post">
<input type="Hidden" name="cabnro" value="0">
<input type="Hidden" name="menuraiz" value="<%= l_menuraiz %>">
</form>
</body>
</html>
