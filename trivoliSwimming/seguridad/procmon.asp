<!--#include virtual="/turnos/shared/db/conn_db.inc"-->

<%
'Archivo: procmon.asp
'Descripción: Se encarga de mostrar los procesos del sistema
'Autor : Lisandro Moro
'Fecha : 10/03/2005
'Modificado:

'on error resume next
on error goto 0

Dim l_filtro
Dim l_refrescar
Dim rs1
Dim sql

l_filtro    = request.querystring("filtro")
l_refrescar = request.querystring("refrescar")

if l_filtro = "" then
   l_filtro = " iduser ='" & Session("username")  & "' and bprcestado <> 'Procesado'"
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="/turnos/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Monitor de Procesos - </title>

<script language="JavaScript">
var jsSelRow = null;

function Deseleccionar(fila){
	fila.className = "MouseOutRow";
}

function Seleccionar(fila,cabnro){
	if (jsSelRow != null){
		Deseleccionar(jsSelRow);
	};
	document.datos.cabnro.value = cabnro;
	fila.className = "SelectedRow";
	jsSelRow		= fila;
}

function refresh(){
	document.control.location = 'procmon01.asp?filtro=' + escape("<%= l_filtro%>");
	setTimeout( "refresh()", 60*1000 );	
} 
</script>

</head>
<body leftmargin="0" rightmargin="0" topmargin="0">
<table border="0" cellpadding="0" cellspacing="1">
<tr>
	<td colspan="11" class="th2">Procesos</td>
</tr>
    <tr>
        <th>Nro </th>
        <th>Fecha </th>
        <th>Hora </th>
        <th>Proceso </th>
        <th>Usuario </th>
        <th>Estado </th>
    </tr>
<%
 if l_refrescar = "SI" then

    Set rs1 = Server.CreateObject("ADODB.RecordSet")
    sql = "SELECT bpronro, bprchora, bprcfecha,btprcdesabr,bprcestado,bprcprogreso,iduser "
	sql = sql & "FROM batch_proceso, batch_tipproc "
	sql = sql & "where batch_tipproc.btprcnro = batch_proceso.btprcnro "
	if l_filtro<>"" then
		sql = sql & " and " & l_filtro
	end if
	sql = sql & " ORDER BY bpronro desc "
    'rsOpen rs1, cn, sql, 0
    rs1.Open sql, cn
  
    if not err then

	    do until rs1.eof
	
	%>
	    <tr onclick="Javascript:Seleccionar(this,<%= rs1("bpronro")%>)">
	        <td><%= rs1("bpronro") %> </td>
	        <td><%= rs1("bprcfecha") %> </td>
	        <td><%= rs1("bprchora") %> </td>
	        <td nowrap><%= rs1("btprcdesabr") %> </td>
	        <td><%= rs1("iduser") %> </td>
	        <td nowrap><%= rs1("bprcestado") %> <% If rs1("bprcestado") = "Procesando" Then %> &nbsp;<%= rs1("bprcprogreso") %>% <% End If %></td>
	    </tr>
	<%
	      rs1.movenext
	    loop
		rs1.close
		
	end if
	
  end if
  
'if l_refrescar = "SI" then
'    Set rs1 = Server.CreateObject("ADODB.RecordSet")
'    sql = "SELECT bpronro, bprchora, bprcfecha,btprcdesabr,bprcestado,bprcprogreso,iduser "
'	sql = sql & "FROM batch_proceso, batch_tipproc "
'	sql = sql & "where batch_tipproc.btprcnro = batch_proceso.btprcnro "
'	if l_filtro<>"" then'
'		sql = sql & " and " & l_filtro
'	end if
'	sql = sql & " ORDER BY bpronro desc "
'    'rsOpen rs1, cn, sql, 0
'    rs1.Open sql, cn
'    if not err then
'		if rs1.eof then
'		%><!--<tr><td colspan="6"><b>No se encontraron Procesos</b></td></tr>--><%
'		else
'		    do until rs1.eof	%>
		    <!--<tr onclick="Javascript:Seleccionar(this,<%'= rs1("bpronro")%>)">
		        <td align="center"><%'= rs1("bpronro") %> </td>
		        <td align="center"><%'= rs1("bprcfecha") %> </td>
		        <td align="center"><%'= rs1("bprchora") %> </td>
		        <td align="center" nowrap><%'= rs1("btprcdesabr") %> </td>
		        <td align="center"><%'= rs1("iduser") %> </td>
		        <td align="center" nowrap><%'= rs1("bprcestado") %> <% 'If rs1("bprcestado") = "Procesando" Then %> &nbsp;<%'= rs1("bprcprogreso") %>% <% 'End If %></td>
		    </tr>-->
		<%	'	rs1.movenext
'		    loop
'		end if
'		rs1.close
'	else%>
		<!--<tr><td colspan="6"><b>Se detectaron errores</b></td></tr>-->
	<%'end if
'else%>
		<!--<tr><td colspan="6"><b>Aplicacion Inactiva</b></td></tr>-->
	<%
'end if
%>
</table>

<form name="datos" method="post">
<input type="Hidden" name="cabnro" value="0" >
</form>
<iframe name="control" height="0" width="0" src=""></iframe>

<% if l_refrescar = "SI" then %>
<script>
   setTimeout( "refresh()", 60*1000 );
</script>
<% else %>
<script>
  document.control.location = 'procmon01.asp?filtro=' + escape("<%= l_filtro%>");
</script>
<% end if%>

</body>
<% cn.close %>
</html>
