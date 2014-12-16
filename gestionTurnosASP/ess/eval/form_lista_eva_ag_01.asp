<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<%

'=====================================================================================
'Archivo  : form_lista_eva_ag_01.asp
'Objetivo : lista de formularios de evaluacion
'Fecha	  : 06-05-2004
'Autor	  : CCRossi 
'Fecha	  : 22-10-2004-Cambiar Evento por Periodo
'Fecha	  : 22-10-2004-Arreglar duplicacion para registros del autoevaluador logeado
'Fecha	  : 04-11-2004-Si es superevaluador para ABN, que muestre las evaluaciones aunque no esté aun habilitado.
'Fecha	  : 17-03-2005-Cambiar dobleclick segun Cliente
'			 11/10/2005 - Leticia Amadio - autogestion
' 		    11-11-2005 - LA. - Mostrar solo las evaluaciones en la que el logueado se autoevalua
'		    18-08-2006 - LA. - Sacar la vista v_empleado
'			18-08-2006 - LA. - se saco la restriccion del 11-11-2005 - se mustra evaluaciones dde lo evaluan
'			25-10-2006 - LA. - mostrar cartel de que no hay evaluaciones cdo sea necesario.
'====================================================================================================

'on error goto 0

'variables 
 Dim l_rs
 Dim l_rs1
 Dim l_sql
 dim l_filtro2
 dim l_nombre

dim l_listainicial

 dim l_color
 dim l_yapaso 
 dim l_entro
  
'parametros
 dim l_filtro
 dim l_orden
 dim l_logeadoternro
 dim l_logeadoempleg
 
l_filtro = request("filtro")
l_orden  = request("orden")

l_empleg = request.QueryString("empleg")
l_logeadoternro = l_ess_ternro
l_logeadoempleg = l_ess_empleg

'l_logeadoternro  = request("logeadoternro") ' viene el ternro del empleg de autogestion o de ....
'l_logeadoempleg = Session("empleg")



if len(l_filtro) <> 0 then
	if left(l_filtro,1) <> "'" then
		l_filtro2 = "'" & l_filtro & "'"
	else
		l_filtro2 =  mid(l_filtro,2,len(request("filtro")) - 1)
	end if	
end if	

if l_orden = "" then
	l_orden = " ORDER BY evaevedesabr, empleado.terape"
end if

' ________________________________________
' ________________________________________
sub armarnombre(nombre)

	l_nombre = l_rs("terape")
	if trim(l_rs("terape2"))<>"" then
		l_nombre = l_nombre & " " & trim(l_rs("terape2"))
	end if	
	if trim(l_rs("ternom"))<>"" or trim(l_rs("ternom2"))<>"" then
		l_nombre = l_nombre & ","
	end if	
	if trim(l_rs("ternom"))<>"" then
		l_nombre = l_nombre & " " & trim(l_rs("ternom"))
	end if	
	if trim(l_rs("ternom2"))<>"" then
		l_nombre = l_nombre & " " & trim(l_rs("ternom2"))
	end if	
	
end sub
%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" http-equiv="refresh" content="text/html; charset=iso-8859-1">
<title>Proceso de Gesti&oacute;n de Desempe&ntilde;o - Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
</head>
<style>
<%if ccodelco=-1 then%>
.autoelegida
{
	COLOR: Black;
	FONT-FAMILY: Verdana;
	FONT-SIZE: 08pt;
	BACKGROUND-COLOR: "#DFBDA2";
	padding : 2;
	padding-left : 5;
}
.autoNOelegida
{
	COLOR: Black;
	FONT-FAMILY: Verdana;
	FONT-SIZE: 08pt;
	BACKGROUND-COLOR: "#FFFCD7";
	padding : 2;
	padding-left : 5;
}
<%else%>
.autoelegida
{
	COLOR: Black;
	FONT-FAMILY: Verdana;
	FONT-SIZE: 08pt;
	BACKGROUND-COLOR: "#B0E0E6";
	padding : 2;
	padding-left : 5;
}
.autoNOelegida
{
	COLOR: Black;
	FONT-FAMILY: Verdana;
	FONT-SIZE: 08pt;
	BACKGROUND-COLOR: "#fffaf2";
	padding : 2;
	padding-left : 5;
}
<%end if%>
</style>

<script>
var jsSelRow = null;
var color = null;

function Deseleccionar(fila)
{

 if (color!==1)
	fila.className = "MouseOutRow";
 else
	fila.className = "autoNOelegida";
	
}

function Seleccionar(fila,cabnro,evaevenro,evatevnro,empleg)
{
	

 if (jsSelRow != null)
 {
  Deseleccionar(jsSelRow);
 };

 document.datos.cabnro.value = cabnro;
 document.datos.evaevenro.value = evaevenro;
 document.datos.empleg.value = empleg;
 if (evatevnro!==1)
 {
	fila.className = "SelectedRow";
	jsSelRow       = fila;
 	<%if ccodelco=-1 then%>
 	parent.habbtn();
 	<%end if%>	
	color=2
 }	
 else
 {
 	fila.className = "autoelegida";
	jsSelRow       = fila;
 	<%if ccodelco=-1 then%>
 	parent.deshabbtn();
 	<%end if%>
 	color=1
 }	
 

}

</script>


<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th><%if ccodelco=-1 then%>Per&iacute;odo <%else%>Evento<%end if%></th>
        <th>Formulario</th>
        <th><%if ccodelco=-1 then%>Supervisado<%else%>Empleado a Evaluar<%end if%></th>
        <%if cejemplo=-1 then%>
			<th>Tiene Objetivos</th>
		<%else%>
			<%if ccodelco=-1 then %>
			<th>Usa Compromisos Predefinidos</th>
			<% end if%>
        <%end if%>
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
Set l_rs1 = Server.CreateObject("ADODB.RecordSet")

' Obs: Se saco la restriccion de que solo el empleado pueda entrar a autoevaluarse -  evatevnro= cautoevaluador  
l_sql = "SELECT DISTINCT evacab.evacabnro,tieneobj, evaevento.evaevenro, evaevento.evaevedesabr, evatipoeva.evatipdesabr, empleado.ternro,empleado.empleg,empleado.terape,empleado.terape2,empleado.ternom,empleado.ternom2 "
l_sql = l_sql & " FROM evacab INNER JOIN evaevento ON evaevento.evaevenro = evacab.evaevenro "
l_sql = l_sql & " INNER JOIN evatipoeva   ON evatipoeva.evatipnro = evaevento.evatipnro "
l_sql = l_sql & " INNER JOIN empleado   ON empleado.ternro = evacab.empleado "
l_sql = l_sql & " WHERE EXISTS ( "
l_sql = l_sql & " 	SELECT * FROM evadetevldor "
l_sql = l_sql & "   WHERE evadetevldor.evacabnro=evacab.evacabnro AND  evacab.cabaprobada <> -1  " 
l_sql = l_sql & "	       AND  ( evadetevldor.evaluador="& l_logeadoternro & "	OR evacab.empleado="& l_logeadoternro & ") )" ' evaluaciones dde el empleado evalua o donde es evaluado
if l_filtro <> "" then
 l_sql = l_sql & " AND " & l_filtro 
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
'response.write l_sql

if l_rs.eof then%>
<tr>
	 <td colspan="4">No hay Evaluaciones Habilitadas para el Evaluador o Filtro Ingresado.</td>
</tr>
<%else
	
	l_entro=0
	l_listainicial="0"
	do until l_rs.eof
	
		l_listainicial= l_listainicial & "," & 	l_rs("ternro")
		' se fija si esta habilitado para Evaluar o su Evaluacion fue Terminada SINO es un empleado que fue evaluado (no evalua)
		l_sql =" SELECT DISTINCT evatevnro FROM evadetevldor WHERE evadetevldor.evacabnro="& l_rs("evacabnro") & " AND evadetevldor.evaluador = " & l_logeadoternro & "  AND  (evadetevldor.habilitado = -1 OR evadetevldor.evldorcargada = -1 "
		IF cejemplo=-1 then ' es ABN
			l_sql = l_sql & "   	   OR (evadetevldor.evatevnro  <> " & cautoevaluador
			l_sql = l_sql & "   	   AND evadetevldor.evatevnro  <> " & cevaluador &")"
		end if
		l_sql = l_sql & "   	)"
		rsOpen l_rs1, cn, l_sql, 0 
		'response.write l_sql & "<br>"
		
		if not l_rs1.eof then
			l_yapaso=0
			do while not l_rs1.eof 
				l_entro = -1
				armarnombre l_nombre
				
				if l_yapaso=0 then
					if l_rs1("evatevnro")<>cautoevaluador or l_rs1("evatevnro")<> cevaluador then
						l_yapaso=-1
					end if%>
					<%if cdeloitte=-1 then%>
						<tr <%if l_rs1("evatevnro")=1 then%>class="autoNOelegida" <%end if%> onclick="Javascript:Seleccionar(this,<%=l_rs("evacabnro")%>,<%=l_rs("evaevenro")%>,<%=l_rs1("evatevnro")%>,<%=l_rs("empleg")%>)" ondblclick="Javascript:parent.abrirVentanaVerif('form_carga_eva_DEL_ag_00.asp?evacabnro=<%=l_rs("evacabnro")%>&evaevenro=<%=l_rs("evaevenro")%>&empleg=<%=l_rs("empleg")%>&logeadoempleg=<%=l_logeadoempleg%>' ,'',800,600)">
				 	<%else
						if ccodelco=-1 then%>
							<tr <%if l_rs1("evatevnro")=1 then%>class="autoNOelegida" <%end if%> onclick="Javascript:Seleccionar(this,<%=l_rs("evacabnro")%>,<%=l_rs("evaevenro")%>,<%=l_rs1("evatevnro")%>,<%=l_rs("empleg")%>)" ondblclick="Javascript:parent.abrirVentanaVerif('form_carga_eva_COD_ag_00.asp?evacabnro=<%=l_rs("evacabnro")%>&evaevenro=<%=l_rs("evaevenro")%>&empleg=<%=l_rs("empleg")%>&logeadoempleg=<%=l_logeadoempleg%>&pantalla='+parent.document.datos.pantalla.value,'',800,600)">
						<%else%>
							<tr <%if l_rs1("evatevnro")=1 then%>class="autoNOelegida" <%end if%> onclick="Javascript:Seleccionar(this,<%=l_rs("evacabnro")%>,<%=l_rs("evaevenro")%>,<%=l_rs1("evatevnro")%>,<%=l_rs("empleg")%>)" ondblclick="Javascript:parent.abrirVentanaVerif('form_carga_eva_ag_00.asp?evacabnro=<%=l_rs("evacabnro")%>&evaevenro=<%=l_rs("evaevenro")%>&empleg=<%=l_rs("empleg")%>&logeadoempleg=<%=l_logeadoempleg%>','',800,600)">
						<%end if
					end if%>
				
					<td nowrap><%=l_rs("evaevedesabr")%></td>
					<td nowrap><%=l_rs("evatipdesabr")%></td>
					<td nowrap><%=l_nombre%></td>
					<%if cejemplo=-1 or ccodelco=-1 then%>
					<td align=center><%if cint(l_rs("tieneobj"))=-1 then%>SI<%else%>NO<%end if%></td>
					<%end if%>
					</tr>
				<%end if
				l_rs1.MoveNext
				loop
				
		else '  el empleado no evalua en la evaluacion (el empleado es el evaluado), obs:evatevnro le pongo 1
			
			if l_rs("ternro") = cdbl(l_logeadoternro) then 
				l_entro = -1
				armarnombre l_nombre
			%>
				<tr class="autoNOelegida" onclick="Javascript:Seleccionar(this,<%=l_rs("evacabnro")%>,<%=l_rs("evaevenro")%>,1,<%=l_rs("empleg")%>)" ondblclick="Javascript:parent.abrirVentanaVerif('form_carga_eva_ag_00.asp?evacabnro=<%=l_rs("evacabnro")%>&evaevenro=<%=l_rs("evaevenro")%>&empleg=<%=l_rs("empleg")%>&logeadoempleg=<%=l_logeadoempleg%>','',800,600)">
					<td nowrap><%=l_rs("evaevedesabr")%></td>
					<td nowrap><%=l_rs("evatipdesabr")%></td>
					<td nowrap><%=l_nombre%></td>
					<%if cejemplo=-1 or ccodelco=-1 then%>
					<td align=center><%if cint(l_rs("tieneobj"))=-1 then%>SI<%else%>NO<%end if%></td>
					<%end if%>
					</tr>
			<% end if
		end if
		
		l_rs1.Close	
		
	l_rs.MoveNext
	loop
	
	
	if l_entro=0 then%>
	<tr>
	 <td colspan="3">No hay Evaluaciones Habilitadas para el Evaluador o Filtro Ingresado.</td>
	</tr>
	<%end if
	
end if ' del if l_rs.eof

l_rs.Close
set l_rs = Nothing
set l_rs1=nothing
cn.Close
set cn = Nothing

%>
</table>

<form name="datos" method="post">
<input type="Hidden" name="cabnro" value="0" >
<input type="Hidden" name="listainicial" value="<%=l_listainicial%>" >
<input type="Hidden" name="evaevenro" value="0" >
<input type="Hidden" name="empleg" value="0" >
<input type="Hidden" name="orden"  value="<%= l_orden %>">
<input type="hidden" name="filtro" value="<%= l_filtro2 %>">
</form>

</body>
</html>
