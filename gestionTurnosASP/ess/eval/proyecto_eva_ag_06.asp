<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<% 
'================================================================================
'Archivo		: proyecto_eva_ag_06.asp
'Descripción	: Validacion de nombre unico
'Autor			: CCRossi
'Fecha			: 30-08-2004
'Modificado		: 10-03-2005 - verificar estructuras de socio, gerente, y que exita el cliente - engag
'				: 14-03-2005 - ver si se puede cambiar el revisor o no.
'				: 08-08-2005 - cambio en la consulta de para ver rol:gerente
'================================================================================
on error goto 0

Dim l_tipo
Dim l_rs
Dim l_sql

Dim l_evaproynro
Dim l_evaproynom
	' VIENEN LOS EMPLEG - ver bien.....
Dim l_proysocio  
Dim l_proygerente 
Dim l_proyrevisor 

Dim l_evaclinro
Dim l_evaengnro

Dim texto
dim l_esta
dim l_revanterior

texto = ""
l_tipo 		 = request.QueryString("tipo")
l_evaproynro = request.QueryString("evaproynro")
l_evaproynom = request.QueryString("evaproynom")
	' VIENEN LOS EMPLEG
l_proysocio  = request.QueryString("proysocio")
l_proygerente = request.QueryString("proygerente")
l_proyrevisor = request.QueryString("proyrevisor")
l_evaclinro = request.QueryString("evaclinro")
l_evaengnro = request.QueryString("evaengnro")

'response.write  " alta <br>"

'Response.Write l_tipo & "<br>"
'Response.Write l_evaproynro & "<br>"
'Response.Write l_evaproynom & "<br>"
'Response.Write l_proysocio & "<br>"
'Response.Write l_proyrevisor & "<br>"


'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")

' verificar que exista el cliente y el engagement...........!!!!!!!!
l_sql = "SELECT evaclinro "
l_sql = l_sql & " FROM evacliente "
l_sql = l_sql & " WHERE evaclinro=" & l_evaclinro
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then
	texto = texto & " El Cliente no existe."
end if
l_rs.Close

l_sql = "SELECT evaengnro "
l_sql = l_sql & " FROM evaengage "
l_sql = l_sql & " WHERE evaengnro=" & l_evaengnro
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then
	if texto <>  ""  then texto = texto & " - " end if
	texto = texto & " El Engagement no existe."
end if
l_rs.Close

' verificar que el socio y gerente tengan rol socio y gerentee...
l_sql = "SELECT estrdabr, his_estructura.estrnro "
l_sql = l_sql & " FROM empleado "
l_sql = l_sql & " INNER JOIN his_estructura ON his_estructura.ternro=empleado.ternro"
l_sql = l_sql & "   AND his_estructura.tenro= " & ctenroRol
l_sql = l_sql & "   AND ( his_estructura.htethasta IS NULL OR  his_estructura.htethasta <"& cambiafecha(Date(),"","") &" )"
l_sql = l_sql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro"
l_sql = l_sql & " WHERE empleg =" & l_proysocio
rsOpen l_rs, cn, l_sql, 0 
'response.write l_sql &"<br>"
if not l_rs.eof then
	l_esta= "" 
	do while (not l_rs.eof) 
		if ( UCase(trim(l_rs("estrdabr"))) = "SOCIO") then 
			l_esta="SI"
		end if
	l_rs.MoveNext
	Loop
	
	if l_esta <> "SI" then
		if texto <>  "" then texto = texto & " - \n" end if
		texto = texto & " El Socio elegido no tiene asociado el Rol: Socio"
	end if
else
	if texto <>  "" then texto = texto & " - \n" end if
	texto = texto & " El Socio elegido no tiene asociado el Rol: Socio"		
end if
l_rs.Close


l_sql = "SELECT estructura.estrdabr, his_estructura.estrnro "
l_sql = l_sql & " FROM empleado "
l_sql = l_sql & " INNER JOIN his_estructura ON his_estructura.ternro=empleado.ternro"
l_sql = l_sql & "   AND his_estructura.tenro=" & cTenroRol
l_sql = l_sql & "   AND ( his_estructura.htethasta IS NULL OR  his_estructura.htethasta <"& cambiafecha(Date(),"","") &" )"
l_sql = l_sql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro"
l_sql = l_sql & " WHERE empleg =" & l_proygerente
rsOpen l_rs, cn, l_sql, 0 

'response.write l_sql &"<br>"
if not l_rs.eof then
	l_esta= "" 
	do while (not l_rs.eof) 
		if ( UCase(trim(l_rs("estrdabr"))) = "GERENTE") then
			l_esta="SI"
		end if
	l_rs.MoveNext
	Loop
	
	if l_esta <> "SI" then
		if texto <>  "" then texto = texto & " - \n" end if
		texto =  texto & " El Gerente elegido no tiene asociado el Rol: Gerente.  "
	end if
	
else
	if texto <>  "" then texto = texto & " - \n" end if
	texto = texto & " El Gerente elegido no tiene asociado el Rol: Gerente."
end if

l_rs.Close


'__________________________________________________________________
' Verificar si se cambio el revisor
'l_revanterior=""

'if trim(l_proyrevisor)<>"" and not isnull(l_proyrevisor) then
	'l_sql = "SELECT ternro "
	'l_sql = l_sql & " FROM empleado "
	'l_sql = l_sql  & " WHERE empleg =" & l_proyrevisor
	'rsOpen l_rs, cn, l_sql, 0 
	'if not l_rs.eof then
		'l_proyrevisor= l_rs("ternro")
	'end if
	'l_rs.close
'else
	'l_proyrevisor="null"	
'end if

'l_sql = "SELECT proyrevisor "
'l_sql = l_sql & " FROM evaproyecto "
'l_sql = l_sql  & " WHERE evaproynro = " & l_evaproynro
'rsOpen l_rs, cn, l_sql, 0 
'if not l_rs.eof then
	'l_revanterior = l_rs("proyrevisor")
'end if
'l_rs.close


' si cambia el revisor --> ver si ya se genero la evaluacion (se crearon evadetevldor)
'						--> si se crearon evadetevdr  no se puede actualizar el revisor.
'if l_proyrevisor <> l_revanterior then
	'l_sql = "SELECT evadetevldor.evacabnro "
	'l_sql = l_sql & " FROM evaproyecto "
	'l_sql = l_sql & " INNER JOIN evaevento ON evaevento.evaproynro = evaproyecto.evaproynro "
	'l_sql = l_sql & " INNER JOIN evacab ON evacab.evaproynro = evaevento.evaproynro "
	'l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro = evacab.evacabnro "
	'l_sql = l_sql & " WHERE evaproyecto.evaproynro ="& l_evaproynro
	'rsOpen l_rs, cn, l_sql, 0 
	'if l_rs.eof then ' no hay evaluacin
		' l_proyrevisor = ""
	'else
		'if texto <>  "" then texto = texto & " - \n" end if
		'texto = texto & " El Revisor no se puede cambiar dado que ya se generó la evaluación."
	'end if
	'l_rs.close 
'end if

set l_rs = nothing
' campo descripcion? obligatorio? se puede repetir???
'Verifico que no este repetida la descripción
'l_sql = "SELECT evaproynro "
'l_sql = l_sql & " FROM evaproyecto "
'l_sql = l_sql & " WHERE evaproynom ='" & l_evaproynom & "'"
'if l_tipo = "M" then
	'l_sql = l_sql & " AND evaproynro <> " & l_evaproynro
'end if
'rsOpen l_rs, cn, l_sql, 0
'Response.Write l_sql & "<br>"
'if not l_rs.eof then
	'l_rs.close
    'texto =  "Existe otro Proyecto con este Nombre."
    'set l_rs = nothing
'end if 

%>

<script>
<% if texto <> "" then %>
   parent.invalido('<%=texto %>')
<% else %>
   parent.valido();
<% end if%>
</script>

<%
Set l_rs = Nothing


cn.Close
Set cn = Nothing
%>

