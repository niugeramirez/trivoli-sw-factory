<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<% 
'================================================================================
'Archivo		: evento_evaluacion_06.asp
'Descripción	: Validacion de descripcion unica y de periodos
'Autor			: CCRossi
'Fecha			: 02-12-2004
'Modificado		: 03-12-2004 arreglar validaciones de periodo prox vacio 
'================================================================================

Dim l_tipo
Dim l_rs
Dim l_sql

Dim l_evaevenro
Dim l_evaevedesabr
Dim texto
Dim l_evaperact
Dim l_evaperprox

dim l_actdesde
dim l_acthasta
dim l_proxdesde
dim l_proxhasta

texto = ""
l_tipo 			= request.QueryString("tipo")
l_evaevenro		= request.QueryString("evaevenro")
l_evaevedesabr	= request.QueryString("evaevedesabr")
l_evaperact		= request.QueryString("evaperact")
l_evaperprox	= request.QueryString("evaperprox")

'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")

'Verifico que no este repetida la descripción
l_sql = "SELECT evaevenro FROM evaevento WHERE evaevedesabr='" & trim(l_evaevedesabr) & "'"
if l_tipo = "M" then
	l_sql = l_sql & " AND evaevenro <> " & l_evaevenro
end if
Response.Write l_sql & "<br>"
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
	l_rs.close
    texto =  "Existe otro Evento con esa Descripción Abreviada."

else
if ccodelco<>-1 then
	l_rs.close
	l_sql = "SELECT evaperdesde,  evaperhasta FROM  evaperiodo WHERE evapernro ="& l_evaperact
	rsOpen l_rs, cn, l_sql, 0 
	if not l_rs.eof then
		l_actdesde = l_rs("evaperdesde")
		l_acthasta = l_rs("evaperhasta")
	end if
	
	if trim(l_evaperprox)<>"" and l_evaperprox<>"0" and not isnull(l_evaperprox)  then
		l_rs.close
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT evaperdesde,evaperhasta FROM  evaperiodo WHERE evapernro ="& l_evaperprox
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_proxdesde = l_rs("evaperdesde")
			l_proxhasta = l_rs("evaperhasta")
		end if
		l_rs.close
	
		if CDate(l_actdesde) >= CDate(l_proxdesde) or CDate(l_acthasta) >= CDate(l_proxhasta) then
			 texto =  "El período próximo es anterior o igual al actual."
		end if
	end if	
end if
end if 

%>

<script>
<% 
 if texto <> "" then
%>
   parent.invalido('<% =texto %>')
<% else%>
   parent.valido();
<% end if%>
</script>

<%
Set l_rs = Nothing
%>

