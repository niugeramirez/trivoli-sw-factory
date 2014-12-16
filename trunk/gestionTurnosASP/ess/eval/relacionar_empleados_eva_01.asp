<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<% Server.ScriptTimeout = 720 %>
<%
'---------------------------------------------------------------------------------
'Archivo	: relacionar_empleados_eva_01.asp
'Descripción: browse de empelados del evento
'Autor		: CCRossi
'Fecha		: 18-05-2004
'Modificado	: 13-07-2004 CCRossi. Cambiar el titulo de columna "Evaluadores Sin Asignar" 
'			  por "Evaluadores", asi que cambiar logica para que ponga SI si tiene 
'			  y NO si no tiene asigados
' Modificado: 1 de Junio CCRossi. Separar en listas pequeñas de 500 empleados.

'---------------------------------------------------------------------------------------
on error goto 0

'Variables base de datos
 Dim l_rs
 Dim l_rs1 
 Dim l_sql

'Variables filtro y orden
 dim l_filtro
 dim l_orden

 dim l_listempleados ' viene del form del 00

'locales
 dim l_listacompleta
 dim l_listempleados0
 dim l_listempleados1
 dim l_listempleados2
 dim l_listempleados3
 dim l_listempleados4
 dim l_listempleados5
 dim l_cantidad
 dim l_datos
 dim l_inicio
 
 l_datos	=request("datos")
 l_inicio	=request("inicio")
 
 Dim l_evaevenro

 Dim l_nombre
 Dim l_esta ' para verificar si el empleado ya esta relacionado o no
 Dim l_sinasignar ' para verificar si el empleado tiene todos los evaluadores asignados
 
'Tomar parametros
 l_filtro	 = request("filtro")
 l_orden	 = request("orden")
 l_evaevenro = request("evaevenro")
 l_listempleados = request("listempleados")

 
 if l_inicio="SI" then%>
<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">

<meta http-equiv="Content-Type" http-equiv="refresh" content="text/html; charset=iso-8859-1">
<title><%if ccodelco=-1 then%>Supervisados<%else%>Empleados<%end if%> del Evento -  Gesti&oacute;n de Desempeño - RHPro &reg;</title>
</head>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table id="tabla" name="tabla">
    <tr>
        <th nowrap><%if ccodelco=-1 then%>N&uacute;mero<%else%>Empleado<%end if%></th>
        <th nowrap>Apellido y Nombre</th>
        <th nowrap>Relacionado</th>
        <th nowrap><%if ccodelco=-1 then%>Roles Asignados<%else%>Evaluadores Asignados<%end if%></th>
    </tr>
	<tr>
		<th  colspan="4"><b>SELECCIONE UN FILTRO</td>
	</tr>
</table>
<form name="datos" method="post">
<input type="Hidden" name="cabnro" value="0" >
<input type="Hidden" name="orden" value="<%= l_orden %>">
<input type="hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>

 
 <%else
 l_listempleados0 = "0"
 l_listempleados1 = "0"
 l_listempleados2 = "0"
 l_listempleados3 = "0"
 l_listempleados4 = "0"
 l_listempleados5 = "0"

'Body 
 
 if l_orden = "" then
	l_orden = " ORDER BY empleg"
 end if

'recorrer lista para sacar los empleg y dejar solo ternro
  dim arr1
  dim arr2
  dim i
  dim l_lista 
  Dim l_ternro

  l_lista = "0"
   
  arr1 = Split(l_listempleados,",")
  i=0
  do while i<=Ubound(arr1) 
  
	  arr2 = split(arr1(i),"@")
	  l_lista = l_lista & "," & arr2(0)
	
	i = i+1
  loop	
   
  l_listempleados= l_lista  
  'Response.Write l_listempleados
  'Response.End
  

'armar listas de a 500..................................

arr1 = Split(l_listempleados,",") 

l_cantidad=0

 'response.write Ubound(arr1) &"<br>"

if trim(l_listempleados)<>"" and trim(l_listempleados)<>"0" then
  do while l_cantidad<=Ubound(arr1) 
  
    l_ternro = arr1(l_cantidad)
	
	if l_cantidad < 501 then
 		l_listempleados0 = l_listempleados0 & "," & l_ternro
	else
	   if l_cantidad > 500 and l_cantidad < 1001 then
		l_listempleados1 = l_listempleados1 & "," & l_ternro
	   else
		if l_cantidad > 1000 and l_cantidad < 1501 then
			l_listempleados2 = l_listempleados2 & "," & l_ternro
	   	else
		    if l_cantidad > 1500 and l_cantidad < 2001 then
			l_listempleados3 = l_listempleados3 & "," & l_ternro
		    else
			if l_cantidad > 2000 and l_cantidad < 2501 then
			   l_listempleados4 = l_listempleados4 & "," & l_ternro
	   		else
			    if l_cantidad > 2500 and l_cantidad < 3001 then
				l_listempleados5 = l_listempleados5 & "," & l_ternro
	   		    end if
			end if
		    end if '1500
		end if '1000
	   end if '500
	end if ' menos de 500
	
	l_cantidad = l_cantidad + 1

  loop	
end if

'Response.write l_listempleados0 &"<br>"
'Response.write l_listempleados1 &"<br>"
'Response.write l_listempleados2 &"<br>"
'Response.write l_listempleados3 &"<br>"
'Response.write l_listempleados4 &"<br>"
'Response.write l_listempleados5 &"<br>"
'response.end

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" http-equiv="refresh" content="text/html; charset=iso-8859-1">
<title><%if ccodelco=-1 then%>Supervisados<%else%>Empleados<%end if%> del Evento -  Gesti&oacute;n de Desempeño - RHPro &reg;</title>
</head>
<script>
var jsSelRow = null;

function Deseleccionar(fila)
{
 fila.className = "MouseOutRow";
}
function Seleccionar(fila,codigo){
 if (jsSelRow != null)
    Deseleccionar(jsSelRow);


 document.datos.cabnro.value = codigo;
 fila.className = "SelectedRow";
 fila.focus();
 jsSelRow		= fila;
 
}

</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table id="tabla" name="tabla">
    <tr>
        <th nowrap><%if ccodelco=-1 then%>N&uacute;mero<%else%>Empleado<%end if%></th>
        <th nowrap>Apellido y Nombre</th>
        <th nowrap>Relacionado</th>
        <th nowrap><%if ccodelco=-1 then%>Roles Asignados<%else%>Evaluadores Asignados<%end if%></th>
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT DISTINCT ternro, empleg, terape,terape2, ternom, ternom2 FROM empleado WHERE "
if trim(l_listempleados0)="" or trim(l_listempleados0)="0" then
l_sql = l_sql & " (EXISTS (SELECT * FROM evacab WHERE evacab.empleado = empleado.ternro AND evaevenro=  " & l_evaevenro & "))"
else
l_sql = l_sql & " (ternro IN (" & l_listempleados0 & ") OR ternro in (" & l_listempleados1 & ") OR ternro IN (" & l_listempleados2 & ") OR ternro IN (" & l_listempleados3 & ") or ternro IN (" & l_listempleados4 & ") or ternro IN (" & l_listempleados5 & "))"
end if
if trim(l_filtro)<>"" then
l_sql = l_sql & " AND ( " & l_filtro & " )" 
end if
l_sql = l_sql & l_orden	
'response.write l_sql & "<br>"
'response.end
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	<%if l_datos="NO" then%>
	<td  colspan="4">Buscando Información...</td>
	<%else
		if trim(l_filtro)<>"" then%>
			<td  colspan="4">No hay <%if ccodelco=-1 then%>Supervisados<%else%>Empleados<%end if%> en el evento para el FILTRO.</td>
		<%else%>
			<td  colspan="4">No hay <%if ccodelco=-1 then%>Supervisados<%else%>Empleados<%end if%> en el evento.</td>
		<%end if
	end if%>
</tr>
<%else
	if trim(l_filtro)<>"" then%>
	<tr>
		<th  align=center colspan="4"><B><font size="1.5">RECUERDE QUE HAY UN FILTRO APLICADO.</font></b></td>
	</tr>
	<%
	end if
	do until l_rs.eof
		l_nombre = l_rs("terape")
		if l_rs("terape2") <>"" then
		l_nombre = l_nombre & " " &l_rs("terape2") 
		end if
		if l_rs("ternom") <>"" or l_rs("ternom2") <>"" then
		l_nombre = l_nombre & ", " 
		end if
		l_nombre = l_nombre & l_rs("ternom") 
		l_nombre = l_nombre & " " &l_rs("ternom2") 
		
		Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT * FROM empleado WHERE (EXISTS (SELECT * FROM evacab WHERE evacab.empleado = empleado.ternro AND evaevenro=  " & l_evaevenro & ")) AND empleado.ternro = " & l_rs("ternro")
		rsOpen l_rs1, cn, l_sql, 0 
		l_esta=0
		if not l_rs1.EOF then
			l_esta=-1
		end if
		l_rs1.close	
		
		l_sinasignar = 0
		if l_esta=-1 then
			l_sql = "SELECT * FROM evadetevldor INNER JOIN evacab ON evadetevldor.evacabnro=evacab.evacabnro AND evacab.empleado = "& l_rs("ternro") & " AND evaevenro=  " & l_evaevenro & " WHERE evadetevldor.evaluador IS NULL"
			rsOpen l_rs1, cn, l_sql, 0 
			if not l_rs1.EOF then
				l_sinasignar =-1
			end if
			l_rs1.close	
		else
			l_sinasignar =-1
		end if
		set l_rs1=nothing
		%>
	<tr onclick="Javascript:Seleccionar(this,<%=l_rs("ternro")%>)">
		<td nowrap><%=l_rs("empleg")%> </td>
		<td nowrap><%=l_nombre%> </td>
		<td nowrap align=center><%if l_esta=-1 then%>SI<%else%>NO<%end if%></td>
		<td nowrap align=center><%if l_sinasignar=-1 then%>NO<%else%>SI<%end if%></td>
	</tr>
	<%l_rs.MoveNext
	loop
end if ' del if l_rs.eof
l_rs.Close
set l_rs = nothing

cn.Close	
set cn = nothing
%>
</table>

<form name="datos" method="post">
<input type="Hidden" name="cabnro" value="0" >
<input type="Hidden" name="orden" value="<%= l_orden %>">
<input type="hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
<%end if%>