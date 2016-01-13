<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->

<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<head>
<link href="/turnos/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Ayuda - Ticket</title>
<script language="JavaScript">
var jsSelRow = null;

function SelEmp(tex)
{
    parent.returnValue = tex;
	parent.close();
   	return true;
}
	
function Deseleccionar(fila)
{
 fila.className = "MouseOutRow";
}

function Seleccionar(fila)
{
 if (jsSelRow != null)
 {
  Deseleccionar(jsSelRow);
 };
 fila.className = "SelectedRow";
 jsSelRow		= fila;
}
</script>
	
</script>	
</head>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<%
Dim l_titulo
Dim l_descrip
Dim l_tabla
Dim l_des
Dim l_has
Dim l_campocod
Dim l_campodest
Dim l_titcols
Dim l_where
Dim l_cantidad
Dim l_actual
Dim l_i
Dim l_primero
Dim l_ultimo
Dim l_rs
Dim l_sql
Dim l_texto
Dim l_texto2

l_titulo    = request("titulo")
l_descrip   = request("descrip")
l_tabla     = request("tabla")
l_des       = request("des")
l_has       = request("has")
l_campocod  = request("campocod")
l_campodest = request("campodest")
l_titcols   = request("titcols")
l_where     = request("where")

if l_des = "" then
  l_des = 0
end if
  
if l_has = "" then
  l_has = 0
end if

if l_descrip = "" then
  l_descrip = " "
end if
%>
<table border=0>
<form action="ayuda_codigochar.asp" method="get" name="f1">
<input type="Hidden" name="titulo" value="<%= l_titulo %>">
<input type="Hidden" name="tabla" value="<%= l_tabla %>">
<input type="Hidden" name="campocod" value="<%= l_campocod %>">
<input type="Hidden" name="campodest" value="<%= l_campodest %>">
<input type="Hidden" name="titcols" value="<%= l_titcols %>">
<input type="Hidden" name="where" value="<%= l_where %>">
<input type="Hidden" name="Acepta" value="Si">

<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")
if l_des = 0 then
	l_sql = "SELECT min(" & l_campocod & ") FROM " & l_tabla
	rsOpen l_rs, cn, l_sql, 0 
	l_primero = l_rs(0)
	l_rs.close
else
	l_primero = l_des
end if

if l_has = 0 then
	l_sql = "SELECT max(" & l_campocod & ") FROM " & l_tabla
	rsOpen l_rs, cn, l_sql, 0 
	l_ultimo = l_rs(0)
	l_rs.close
else
	l_ultimo = l_has
end if
%>
<tr>
	<td colspan="10" class="th2"><%= l_titulo %></td>
</tr>
<tr>
	<td>Desde: </td>
	<td>
		<input type="text" name="des" value="<%= l_primero %>" size="5" style="height: 20; background: White;" onReturn="f1.submit">
	    Hasta: <input name="has" size=5 value="<%= l_ultimo %>" style="height: 20; background: White;" onReturn="f1.submit">
	</td>
</tr>
<tr>
	<td>Descripción: </td>
	<td><input type="text" name="descrip" size="10" value="<%= l_descrip %>" style="height: 20; background: White;" onReturn="f1.submit">
	    <input type=submit value="" style="height: 20;"></td>
</tr>
</form>
</table>
<table centered width=100%>
    <tr>
	<%
	l_cantidad = 0
	do while len(l_titcols) > 0
	  if inStr(l_titcols,";") <> 0 then
	    l_actual  = left(l_titcols, inStr(l_titcols,";") - 1)
	    l_titcols = mid (l_titcols, inStr(l_titcols,";") + 1)
	  else
	    l_Actual = l_titcols
		l_titcols = ""
	  end if
	  l_cantidad = l_cantidad + 1
    %>  
      <th><%= l_actual %></th>
    <%
	loop
    %>  
    </tr>
    <%  
    IF request("Acepta") <> "" THEN
	  if l_where = "" then
	    l_where = "(1=1)"
	  end if
      l_sql = "SELECT " & l_campocod & ", " & l_campodest
	  l_sql = l_sql & " FROM " & l_tabla & " WHERE " & l_where 
	  l_sql = l_sql & " AND ('" & l_des & "' = '' OR " & l_campocod & " >= '" & l_des & "')"
	  l_sql = l_sql & " AND ('" & l_has & "' = '' OR " & l_campocod & " <= '" & l_has & "')"
	  if l_descrip <> " " then l_sql = l_sql & " AND (terape LIKE """ & l_descrip & "%" & """)" end if
	  l_sql = l_sql & " ORDER BY " &  l_campocod
      rsOpen l_rs, cn, l_sql, 0
	  do until l_rs.eof 
	    l_texto  = l_rs(0) & "__"
		l_texto2 = ""
	    for l_i = 0 to l_rs.fields.count - 1
		  if l_i <> 0 then
		    l_texto = l_texto & ", " & l_rs(l_i)
		  end if
          l_texto2 = l_texto2 & "<td align='center'>" & l_rs(l_i) & " </td>"
	    next
    %>
    <tr ondblclick="javascript:SelEmp('<%= l_texto %>')" onclick="javascript:Seleccionar(this)"> 
        <%= l_texto2 %>
    </tr> 
    <%
 	    l_rs.movenext
	  loop
	  l_rs.close
    end if
    %>
</table>
</body>
<% cn.close %>
</html>
