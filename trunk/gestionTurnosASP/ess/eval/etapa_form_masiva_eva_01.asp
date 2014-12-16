<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/adovbs.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/asistente.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<script>
var xc = screen.availWidth;
var yc = screen.availHeight;
window.moveTo(xc,yc);	
window.resizeTo(150,150);
</script>
<%
'========================================================================================
'Archivo	: etapa_form_masiva_eva_01.asp
'Descripción: grabar cambio de etapa
'Autor		: CCRossi
'Fecha		: 24-05-2004
'========================================================================================


'ADO
 Dim l_cm
 Dim l_sql
 Dim l_rs
 
'parametros
 Dim l_evaevenro
 Dim l_evaetanro
 Dim l_evatipnro
 Dim l_todos
 l_evaevenro  = Request.Form("evaevenro")
 l_evaetanro  = Request.Form("evaetanro")
 l_evatipnro  = Request.Form("evatipnro")
 l_todos	  = Request.Form("todos")
'uso local
 dim i
 Dim l_grabar
 Dim l_aux
  
 'cn.BeginTrans
 
  l_evatipnro = "," & l_evatipnro
  l_grabar = Split(l_evatipnro,",")
  
  'Response.Write("<script>alert('"&l_todos&"')</script>")
  'Response.end
  
  for i = 1 to Ubound(l_grabar) 
	 
	l_aux = l_grabar(i)
	
	'busco todos los EVENTOS que tiene el formulario ------------------------------
    Set l_rs = Server.CreateObject("ADODB.RecordSet")		
	l_sql = "SELECT evaevenro "
	l_sql = l_sql & " FROM evaevento"
	l_sql  = l_sql  & " WHERE evatipnro = " & l_aux
	if trim(l_todos)<>"on" then
		l_sql = l_sql & " AND evaevenro = " & l_evaevenro
	end if
	rsOpen l_rs, cn, l_sql, 0 
	do while not l_rs.eof 
		'Actualizao TODAS las Cabeceras del EVENTO  ------------------------------
		l_sql = " UPDATE evacab  "
		l_sql = l_sql & " SET evaetanro = " & l_evaetanro
		l_sql = l_sql & " WHERE evaevenro = " & l_rs("evaevenro")
		 'response.write l_sql
		set l_cm = Server.CreateObject("ADODB.Command")  
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
		l_rs.MoveNext
	loop
		
	l_rs.Close
	set l_rs=nothing
    
 next	 

  
 ' cn.CommitTrans 
 
 Set cn = Nothing
 Set l_cm = Nothing
response.write "<script>alert('Operación Realizada.');window.close();</script>"
%> 