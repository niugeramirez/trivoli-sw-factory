<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : criterios_05.asp
Descripcion    : Complemento de criterios
Creador        : Scarpa D.
Fecha Creacion : 28/11/2003
Modificacion   :
-----------------------------------------------------------------------------
-->
<% 
on error goto 0

Dim l_rs
Dim l_sql

Dim l_selnro
Dim l_clase
Dim l_seleccion
Dim l_selsql
Dim l_selasp

l_selnro    = request("selnro")
l_clase     = request("clase")
l_seleccion = request("seleccion")
l_selsql    = request("sql")
l_selasp    = request("asp")

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<%if CInt(l_clase) = 3 then %>
<link href="/serviciolocal/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<%else%>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<%end if%>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Liquidaci&oacute;n de haberes - B&uacute;squedas - RHPro &reg;</title>
</head>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">

<form name="datos">
<%select case CInt(l_clase)%>

<%
'--------------------------------------------------------------------------------------------------------------
'Es una consulta SQL
case 1
%>

<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
   <tr>
     <td width="100%" align="left" valign="top">
	    <b>Consulta SQL:</b><br>
	    <textarea name="sql" rows="6" cols="50" ><%=trim(l_selsql)%></textarea>
	 </td>   
   </tr>
</table>

<%
'--------------------------------------------------------------------------------------------------------------
'Es un programa
case 2
%>

<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
   <tr>
	 <td width="100%" align="center" valign="middle">
	    <b>Programa:</b><br>
	    <input type="text" name="asp" size="50" maxlength="100" value="<%= l_selasp%>">
	 </td>   
   </tr>   
</table>

<%
'--------------------------------------------------------------------------------------------------------------
'Es una lista de empleados
case 3
%>

<table>
    <tr>
        <th>Empleado</th>	
        <th>Apellido y Nombre</th>
    </tr>
<%

if l_seleccion <> "" then

   Set l_rs = Server.CreateObject("ADODB.RecordSet")

   Dim l_arr
   Dim l_arr2
   Dim l_i

   l_arr = Split(l_seleccion,",")		
   
   if instr(1,l_arr(1),"@") then

	   l_seleccion = "0"

	   for l_i = 1 to UBound(l_arr)
	       l_arr2 = split(l_arr(l_i),"@")
	       l_seleccion = l_seleccion & "," & l_arr2(0)
	   next

   end if

   l_sql = "SELECT * "
   l_sql = l_sql & " FROM v_empleado "
   l_sql = l_sql & " WHERE ternro IN ( " & l_seleccion & " ) "
   l_sql = l_sql & " ORDER BY empleg "

   rsOpen l_rs, cn, l_sql, 0 
   do until l_rs.eof
   %>
    <tr>
        <td><%= l_rs("empleg")%></td>	
		<td><%= l_rs("terape") & ", " & l_rs("ternom") %></td>
    </tr> 
   <%
	  l_rs.MoveNext
    loop
    l_rs.Close
    set l_rs = Nothing

end if
%>
</table>

<%
end select
'--------------------------------------------------------------------------------------------------------------

cn.Close
set cn = Nothing
%>

</form>

</script>

</body>
</html>