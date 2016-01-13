<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
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
dim l_indice
dim l_orden

l_menuraiz = request.QueryString("menuraiz")
l_orden = request.QueryString("orden")

if l_orden = "" then
  l_orden = 1
end if

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
<link href="/turnos/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<script>
var jsSelRow = null;

function Deseleccionar(fila)
{
 fila.className = "MouseOutRow";
}
function Seleccionar(fila,cabnro,nivel)
{
 if (jsSelRow != null)
 {
  Deseleccionar(jsSelRow);
 };

 document.datos.cabnro.value = cabnro;
 parent.menu2.location = "armado_menu_02.asp?menuraiz=" + document.datos.menuraiz.value + "&menuorder=" + cabnro + "&nivel=" + nivel;
 fila.className = "SelectedRow";
 jsSelRow		= fila;
}
</script>
</head>

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

sql = sql & " ORDER BY parent, menuorder"
rsOpen rs, cn, sql, 0
menumstr = rs.GetRows
rs.Close

sub imprimir(nodo,nivel)
	for i = 0 to Ubound(menumstr,2)
		if uCase(trim(menumstr(3,i))) = trim(nodo & UCase(l_menudesc)) then
			if trim(menumstr(0,i)) <> "rule" then
    			tr = "<tr onclick='Javascript:Seleccionar(this,"&trim(menumstr(1,i))& "," & nivel &")'>"
				response.write tr & "<td>"
				for j = 1 to nivel
					response.write "-&nbsp;"
				next
				response.write menumstr(0,i) & "</td><td style='display:none' width='1'>" & l_indice & "</td>"
				response.write "<td style='display:none' width='1'>" & menumstr(1,i) & "</td>"
				response.write "<td style='display:none' width='1'>" & menumstr(2,i) & "</td>"
				response.write "<td style='display:none' width='1'>" & menumstr(3,i) & "</td>"
				response.write "<td style='display:none' width='1'>" & menumstr(4,i) & "</td>"
				response.write "<td style='display:none' width='1'>" & nivel & "</td>"
            	response.write "</tr>"
  	            l_indice = l_indice + 1
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
<form name="datos" method="post">
<input type="Hidden" name="cabnro" value="0">
<input type="Hidden" name="menuraiz" value="<%= l_menuraiz %>">
</form>
<table id="tabla">
<%
sql = "SELECT MenuName, MenuOrder, MenuRaiz, Parent, tipo, action, menuaccess, menuimg FROM menumstr where menuraiz = " & l_menuraiz
sql = sql & " AND upper(parent) = '" & UCase(l_menudesc) & "'"
sql = sql & " ORDER BY parent, menuorder"
rsOpen rs, cn, sql, 0
l_indice = 1
do until rs.eof
    tr = "<tr onclick='Javascript:Seleccionar(this,"&rs(1)&")'>"
	response.write tr & "<td>" & rs(0) & "</td><td style='display:none' width='1'>" & l_indice & "</td>"
	response.write "<td style='display:none' width='1'>" & rs(1) & "</td>"
	response.write "<td style='display:none' width='1'>" & rs(2) & "</td>"
	response.write "<td style='display:none' width='1'>" & rs(3) & "</td>"
	response.write "<td style='display:none' width='1'>" & rs(4) & "</td>"
	response.write "<td style='display:none' width='1'>0</td>"
	response.write "</tr>"

	l_indice = l_indice + 1
	imprimir rs(1),1
	rs.moveNext
loop
rs.close
%>	
</table>
</body>
</html>
