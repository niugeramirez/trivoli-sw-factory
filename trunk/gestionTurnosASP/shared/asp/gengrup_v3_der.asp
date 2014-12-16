<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 
Dim rs
Dim sql
Dim l_filtro
Dim l_orden
Dim l_lista
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_canttotal
Dim l_cantFiltro


l_filtro = request("filtro")
l_orden  = request("orden")
l_lista  = request("lista")

function esta(ternro)
	if Instr(1,","&l_lista&",",","&ternro&",")= 0 then
		esta= false
	else
		esta= true
	end if

end function


if l_orden = "" then
  l_orden = " ORDER BY empleg"  'orden por defecto legajo
end if

if l_lista = "" then
  l_lista = "0"
end if

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<link href="/turnos/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<html>
<head>
	<title><%= Session("Titulo")%>Untitled</title>
<script languaje="javascript">
function Cargar1(){
<%

Set rs = Server.CreateObject("ADODB.RecordSet")

'sql = "SELECT count(*) FROM v_empleado WHERE ternro IN (" & l_lista & ")"
'rsOpen rs, cn, sql, 0 

'response.write "parent.totalder.value = " & rs(0) & ";"  & vbCrLf

'rs.Close

sql = "SELECT ternro,empleg, terape, ternom "
sql = sql & "FROM v_empleado "
'sql = sql & "WHERE ternro IN (" & l_lista & ")"

if l_filtro <> "" then
  sql = sql & " WHERE " & l_filtro & " "
end if
sql = sql & l_orden
rsOpen rs, cn, sql, 0 
l_cantfiltro = -1
do until rs.eof
	if esta(rs("ternro")) then
	    response.write "newOp = new Option();" & vbCrLf
	    response.write "newOp.value  = '" & rs("ternro") & "';" & vbCrLf
	    response.write "newOp.text   = '" & rs("empleg") & " - " & rs("terape") & ", " & rs("ternom") & "';"  & vbCrLf
	    l_cantfiltro = l_cantfiltro + 1
	    response.write "document.registro.selfil.options[" & l_cantfiltro & "] = newOp;" & vbCrLf
	end if	
	
	rs.MoveNext
loop
rs.Close


response.write " if (parent.totalder.value=='') parent.totalder.value = " & l_cantfiltro + 1 & ";"  & vbCrLf


set rs = Nothing
cn.Close
set cn = Nothing
%>  
}

</script>	
</head>

<body topmargin="0" leftmargin="0" rightmargin="0" scroll=no>
<form name="registro">
<input type="Hidden" name="lista" value="<%= l_lista %>">
<select class="gengrup" size=20  width="100%" name="selfil" ondblclick="parent.Uno(selfil,parent.nselfil.registro.nselfil, parent.totalder, parent.totalizq);"></select>
</form>
<script>
Cargar1();
parent.filtroder.value = document.registro.selfil.length;
</script>
</body>
</html>
