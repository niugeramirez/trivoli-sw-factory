<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/adovbs.inc"-->
<!--
Archivo: 		ag_evaluar_eventos_cap_03.asp
Descripción: 	Muestra las preguntas del formularios
Autor : 		Raul Chinestra (listo)
Fecha: 			25/06/2007
-->

<% 
on error goto 0

Dim l_ternro
Dim l_tesnro
Dim l_rs
Dim l_rs2
Dim l_sql
Dim l_sql2
Dim l_predesabr

Dim l_fornro

Dim l_cantidad_preguntas
Dim l_pregunta
Dim l_tesfin
Dim l_estado
Dim l_preexp

l_ternro	= request.QueryString("ternro")
l_tesnro	= request.QueryString("tesnro")
l_fornro	= request.QueryString("fornro")
l_pregunta	= request.QueryString("pregunta")
l_tesfin	= request.QueryString("tesfin")

if (l_pregunta = "") or (cint(l_pregunta) < 1) then
	l_pregunta = 1
end if

l_cantidad_preguntas = 1

if Cint(l_tesfin) = -1 then
	l_estado = " readonly disabled "
else
	l_estado = ""
end if

function selecta
	If l_rs("resval") = Cint(l_rs2("opcnro")) then 
		Response.write "checked"
	End If
end function

Set l_rs = Server.CreateObject("ADODB.RecordSet")
Set l_rs2 = Server.CreateObject("ADODB.RecordSet")

l_sql = "SELECT count(pos_pregunta.prenro) cantidad"
l_sql = l_sql & " FROM pos_pregunta "
l_sql = l_sql & " LEFT JOIN pos_respuesta on pos_respuesta.prenro = pos_pregunta.prenro "
l_sql = l_sql & " and pos_respuesta.tesnro = " & l_tesnro
l_sql = l_sql & " WHERE pos_pregunta.fornro = " & l_fornro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_cantidad_preguntas = Cint(l_rs("cantidad"))
	'l_prenro = 1
else
	l_cantidad_preguntas = 0
	'l_prenro = 0
end if
l_rs.close

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%= c_Estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Evaluaciones - Capacitación - RHPro &reg;</title>
</head>
<script>
function Left(str, n){
	if (n <= 0)
	    return "";
	else if (n > String(str).length)
	    return str;
	else
	    return String(str).substring(0,n);
}

function alertar(texto){
	if (texto.value.length>500) {      
		alert('La cantidad máxima de caracteres es 500');
		texto.value = Left(texto.value,500);
	}
}

//parent.document.all.totalempl.value='5<%'= l_cantidad_registros %>';
parent.document.all.pregunta.value='<%= l_pregunta %>';
parent.document.all.totpregunta.value='<%= l_cantidad_preguntas %>';
//parent.document.all.porpagina.value='1<%'= l_porpagina %>';
	
</script>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos" action="ag_evaluar_eventos_cap_04.asp?fornro=<%= l_fornro %>&tesnro=<%= l_tesnro%>" method="post">
<table border="3" width="100%" >
<%


l_sql = "SELECT  predesabr, pretipo, pos_pregunta.prenro, resval, resdes "
l_sql = l_sql & " FROM pos_formulario "
l_sql = l_sql & " INNER JOIN pos_pregunta on pos_pregunta.fornro = pos_formulario.fornro "
l_sql = l_sql & " LEFT JOIN pos_respuesta on pos_respuesta.prenro = pos_pregunta.prenro "
l_sql = l_sql & " and pos_respuesta.tesnro = " & l_tesnro
l_sql = l_sql & " WHERE pos_formulario.fornro = " & l_fornro
l_sql = l_sql & " ORDER by pretipo DESC" 
'rsOpen l_rs, cn, l_sql, 0 
rsOpencursor l_rs, cn, l_sql, 0, adopenkeyset 
l_rs.absoluteposition = l_pregunta
if l_rs.eof then%>
<tr>
	 <td colspan="3">No existen preguntas para este test.</td>
</tr>
<%else
	'do until l_rs.eof	%>	
		<tr>
			<td align="left" colspan="3">
				<b>&nbsp;&nbsp;<%= l_rs("predesabr") %></b>
				<input type="Hidden" name="pretipo<%= l_rs("prenro") %>" value="<%= l_rs("pretipo") %>">
				<input type="Hidden" name="pretipo" value="<%= l_rs("pretipo") %>">
				<input type="Hidden" name="prenro" value="<%= l_rs("prenro") %>">
			</td>
		</tr>
			<% If l_rs("pretipo") = -1 then %>
		<tr>
			<td align="left" colspan="3">
				<textarea rows="5" cols="70" style="width:100%" name="<%= l_rs("prenro") %>" onKeyUp="javascript:alertar(this);" <%= l_estado %> ><%= l_rs("resdes") %></textarea>
			</td>	
		</tr>
			<% Else  %>
		<tr>
			<td align="left" colspan="3">
 				<table cellspacing="0" cellpadding="0" border="3">
							<!-- Ingreso la Opción SIN CONTESTAR -->
							<tr>
							<td align="center" width="10%">
							<input checked type="Radio" name="<%= l_rs("prenro") %>" value="0" <%= l_estado %> >
							</td>
							<td colspan="3">&nbsp; <i>Sin Contestar</i>
							</td>
							</tr>											

				<%	l_sql2 = "SELECT  opcdesabr, opcnro, opcok "
					l_sql2 = l_sql2 & " FROM pos_opcion "
					l_sql2 = l_sql2 & " WHERE prenro = " & l_rs("prenro")
					l_sql2 = l_sql2 & " ORDER by opcnro " 
					rsOpen l_rs2, cn, l_sql2, 0 
					do until l_rs2.eof %>
							<tr>
							<td align="center">
							<input <% call selecta %> type="Radio" name="<%= l_rs("prenro") %>" value="<%= l_rs2("opcnro") %>" <%= l_estado %>>
							</td>
							<% If Cint(l_tesfin) = -1 Then %>
								<% If l_rs2("opcok") = -1 Then %>							
									<td colspan="3">&nbsp;<b><%= l_rs2("opcdesabr") %></b>
								<% else %>									
									<td colspan="3">&nbsp; <%= l_rs2("opcdesabr") %>
								<% end if %>									
								
							<% else %>
							<td colspan="3">&nbsp; <%= l_rs2("opcdesabr") %>
							<% end if %>							
							</td>
							</tr>							
					<%		l_rs2.MoveNext
					loop
				    l_rs2.close 
				%>
				</table>
			</td>
		</tr>
		<tr>
			<td colspan="3">&nbsp;</td>
		</tr>
			<% End If %>
		<% If Cint(l_tesfin) = -1 Then %>
		<tr>
			<td colspan="3"><b>&nbsp;&nbsp;&nbsp;Explicación / Referencias:</b></td>
		</tr>
		<tr>
			<td>
				<%	
					l_sql2 = "SELECT preexp "
					l_sql2 = l_sql2 & " FROM pos_pregunta "
					l_sql2 = l_sql2 & " WHERE prenro = " & l_rs("prenro")
					
					rsOpen l_rs2, cn, l_sql2, 0 
					if not l_rs2.eof then 
						l_preexp = l_rs2("preexp")
					else
						l_preexp = ""
					end if
					l_rs2.close 
					%>
				<textarea rows="5" cols="70" style="width:100%" name="justificacion"  readonly="readonly" ><%= l_preexp %></textarea>
			</td>
		</tr>
		<% End If %>
	<%
end if
l_rs.Close
set l_rs = Nothing
cn.Close
set cn = Nothing
%>
</table>
</form>
</body>
</html>
