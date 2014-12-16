<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<%
'================================================================================
'Archivo		: proyecto_eva_ag_07.asp
'Descripción	: Incorpora un empleado a un proyecto
'Autor			: Leticia Amadio
'Fecha			: 11-03-2005	
'================================================================================

on error goto 0
Dim l_tipo
Dim l_rs
Dim l_cm
Dim l_sql

Dim l_evaproynro 
dim l_proygerente
Dim l_ternro 
dim l_incorporar

Set l_rs = Server.CreateObject("ADODB.RecordSet")

'l_perfil	= request.Form("perfil")

l_evaproynro	= request.QueryString("proyecto")
l_ternro		= request.QueryString("ternro")

l_sql = " SELECT proygerente FROM evaproyecto WHERE evaproyecto.evaproynro="& l_evaproynro 
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then 
	l_proygerente= l_rs("proygerente")	
end if
l_rs.Close

l_sql = " SELECT ternro FROM evaproyemp "
l_sql = l_sql & " WHERE evaproynro="& l_evaproynro & " AND ternro="& l_ternro ' del logueado
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then 
	l_incorporar = -1
else
	l_incorporar = 0
end if 
l_rs.Close
Set l_rs = Nothing

if (l_incorporar <> 0) then
	set l_cm = Server.CreateObject("ADODB.Command")
	l_sql = "INSERT INTO evaproyemp "
	l_sql = l_sql & " (evaproynro,ternro) "
	l_sql = l_sql & " VALUES ("& l_evaproynro &","
	l_sql = l_sql & l_ternro & ")"
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql 
	cmExecute l_cm, l_sql, 0 
	Set l_cm = Nothing 
	
	'response.write l_proygerente & "<br>"
	'response.write l_ternro
	'response.write "<script>alert(g:'" &l_proygerente & "');</script>"
		'response.write "<script>alert(:'"&l_ternro &"')</script>"
	
	' mandar mail..
	if cUsaMail= -1 and trim(l_ternro)<>"" and trim(l_proygerente)<>"" then %>
		<script>
		abrirVentanaH('mailavisoincorpproy_eva_00.asp?ternro=<%=l_ternro%>&proygerente=<%=l_proygerente%>&evaproynro=<%=l_evaproynro%>',"",5,5);
		</script>
<%	end if 
	Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location.reload();window.close();</script>"
else 
	Response.write "<script>alert('El empleado ya está asociado al proyecto.');window.close();</script>"
end if



'Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location.reload();window.close();</script>"
%>
