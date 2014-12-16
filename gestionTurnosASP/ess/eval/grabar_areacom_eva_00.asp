<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<% 
'Autor			: CCRossi 
'Fecha			: 05-01-2005 - 
'Modificacion	: 16-04- 2005 - LA. 
'=====================================================================================

on error goto 0
' variables
' parametros de entrada ----------------------------------------
  Dim l_evldrnro
  Dim l_evaareacom
  Dim l_evatitnro
  dim l_evaareacomnro  
  
  
  dim l_campo
  
' variables de base de datos ------------------------------------
  Dim l_cm
  Dim l_sql
  Dim l_rs
  
  Dim l_tipo  
  Dim l_grabado
  
' parametros de entrada
  l_evaareacom = left(trim(request("evaareacom")),2500)
  l_evldrnro = request("evldrnro")
  l_tipo = request.querystring("tipo")
  l_grabado = request.querystring("grabado")
  l_evatitnro = request("evatitnro")
  l_evaareacomnro  = request("evaareacomnro")
  
     ' l_campo = request.querystring("campo")

 ' response.write l_evaareacom & "<br>"
 ' response.write l_evldrnro & "<br>"
 ' response.write l_grabado & "<br>"
 ' response.write l_evatitnro & "<br>"
 ' response.write l_evaareacomnro & "<br>"

'BODY ----------------------------------------------------------
if l_tipo ="A" then
	l_sql= "insert into evaareacom (evaareacom,evatitnro,evldrnro) "
	l_sql = l_sql & " values ('" & trim(l_evaareacom) & "'," & l_evatitnro & "," & l_evldrnro &")"
else
	if l_tipo = "M" then
		l_sql = "UPDATE evaareacom SET "
		l_sql = l_sql & " evaareacom = '" & trim(l_evaareacom) & "' "
		l_sql = l_sql & " WHERE  evaareacomnro="& l_evaareacomnro
		'l_sql = l_sql & " AND  evatitnro = "  & l_evatitnro
		'l_sql = l_sql & " AND   evldrnro  = "  & l_evldrnro 
	else
		l_sql = "DELETE from evaareacom "
		l_sql = l_sql & " WHERE evldrnro  = "  & l_evldrnro 
		l_sql = l_sql & " AND  evatitnro  = "  & l_evatitnro 
		l_sql = l_sql & " AND  evaareacomnro ="  & l_evaareacomnro
	end if	
end if	

set l_cm = Server.CreateObject("ADODB.Command")  
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
	
 ' response.write " <script> alert('"&l_campo&"')</script>"
 
response.write "<script> parent."&l_grabado&".value = '" &l_tipo &"'; </script>"

'response.write "<script> parent.location.reload(); </script>"


'response.write "<script> parent.location.onload()={ document.datos."&l_campo&".focus();}; </script>"
%>
