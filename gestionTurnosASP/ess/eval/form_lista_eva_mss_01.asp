<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<%
on error goto 0
'=====================================================================================
'Archivo  : form_lista_eva_mss_01.asp
'Objetivo : lista de formularios de evaluacion MSS
'Fecha	  : 11-11-2005
'Autor	  :  Leticia A.
' 		  : 11-11-05 - Leticia A. - Mostrar solo las evaluaciones en la que el logueado es evaluador y empleado actual es evaluado
'			04-10-2006 - LA- mostar evaluaciones de cqr tipo de evaluador, no solo cevaluador
'====================================================================================================

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
 dim l_actualempleg
 dim l_actualternro
 
l_filtro = request("filtro")
l_orden  = request("orden")

'l_empleg = request.QueryString("empleg")

l_logeadoempleg =  Session("empleg")    ' l_empleglogeado  = Session("empleg") 
l_actualempleg   = request("empleg")   ' estaria en l_ess_empleg


'__________________________________________________________________
'   
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT ternro  from empleado WHERE empleg = " 	& l_logeadoempleg
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	'l_ternrologeado= l_rs("ternro")
	l_logeadoternro = l_rs("ternro")
end if
l_rs.close

l_actualternro= l_ess_ternro

'l_sql = "SELECT ternro  from empleado WHERE empleg = " 	& l_actualempleg
'rsOpen l_rs, cn, l_sql, 0 
'if not l_rs.eof then 
	'l_ternroactual= l_rs("ternro")
'end if
'l_rs.close
set	 l_rs=nothing


'_________________________________________________

if len(l_filtro) <> 0 then
	if left(l_filtro,1) <> "'" then
		l_filtro2 = "'" & l_filtro & "'"
	else
		l_filtro2 =  mid(l_filtro,2,len(request("filtro")) - 1)
	end if	
end if	

if l_orden = "" then
	l_orden = " ORDER BY evaevedesabr, v_empleado.terape"
end if

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
l_sql = "SELECT DISTINCT evacab.evacabnro,tieneobj, evaevento.evaevenro, evaevento.evaevedesabr, evatipoeva.evatipdesabr, v_empleado.ternro,v_empleado.empleg,v_empleado.terape,v_empleado.terape2,v_empleado.ternom,v_empleado.ternom2 "
l_sql = l_sql & " FROM evacab INNER JOIN evaevento ON evaevento.evaevenro = evacab.evaevenro "
l_sql = l_sql & " INNER JOIN evatipoeva   ON evatipoeva.evatipnro = evaevento.evatipnro "
l_sql = l_sql & " INNER JOIN v_empleado   ON v_empleado.ternro = evacab.empleado "
l_sql = l_sql & " WHERE EXISTS ( "
l_sql = l_sql & " 	SELECT * FROM evadetevldor "
l_sql = l_sql & "	WHERE evadetevldor.evacabnro = evacab.evacabnro AND  evacab.cabaprobada <> -1 "
l_sql = l_sql & " 		  AND evadetevldor.evaluador=" & l_logeadoternro &"  AND evatevnro<> "& cautoevaluador  '=cevaluador 
l_sql = l_sql & "		  AND  evacab.empleado="& l_actualternro & " ) " 

'l_sql = l_sql & " WHERE EXISTS ( SELECT * FROM evadetevldor WHERE  evadetevldor.evacabnro = evacab.evacabnro AND  evadetevldor.evaluador = " & l_logeadoternro & "	AND  evacab.cabaprobada <> -1  ) " 

if l_filtro <> "" then
 l_sql = l_sql & " AND " & l_filtro 
end if
l_sql = l_sql & " " & l_orden
'response.write l_sql
rsOpen l_rs, cn, l_sql, 0 

if l_rs.eof then%>
<tr>
	 <td colspan="4">No hay Evaluaciones Habilitadas para el Evaluador o Filtro Ingresado.</td>
</tr>
<%else
	
	l_entro=0
	l_listainicial="0"
	do until l_rs.eof
	
	l_listainicial= l_listainicial & "," & 	l_rs("ternro")
	
	Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
	l_sql = " SELECT DISTINCT evatevnro FROM evadetevldor WHERE evadetevldor.evacabnro =  " & l_rs("evacabnro") & " AND evadetevldor.evaluador = " & l_logeadoternro & "  AND  (evadetevldor.habilitado = -1 OR evadetevldor.evldorcargada = -1 "
	IF cejemplo=-1 then ' es ABN ---  NOSE xq cejemplo tb esta como estandar!!!!!!!!!!!!!!!
		l_sql = l_sql & "   OR (evadetevldor.evatevnro  <> " & cautoevaluador
		l_sql = l_sql & "   AND evadetevldor.evatevnro  <> " & cevaluador &")"
	end if
	l_sql = l_sql & "  	)"
	rsOpen l_rs1, cn, l_sql, 0 
	'response.write l_sql & "<br>"
	l_yapaso=0
	do while not l_rs1.eof 
		l_entro = -1
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
		l_rs1.Close
		set l_rs1=nothing
		
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
