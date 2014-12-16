<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<% 
'=====================================================================================
'Archivo    : grabar_objcom_eva_00.asp
'Objetivo   : ABM de comentarios por objetivos
'Autor		: Leticia Amadio 
'Fecha		: 10-01-2005 
'Modificacion: 16-04- 2005 - LA.
'=====================================================================================


' variables
' parametros de entrada ----------------------------------------
  Dim l_evldrnro
  Dim l_evaobjcom
  Dim l_evaobjnro
  dim l_evaobjcomnro  
  
' variables de base de datos ------------------------------------
  Dim l_cm
  Dim l_sql
  Dim l_rs
  
  Dim l_tipo  
  Dim l_grabado
  
' parametros de entrada
  l_evaobjcom = left(trim(request("evaobjcom")),1500) ' ++++++++ 
  l_evldrnro = request("evldrnro")
  l_tipo = request.querystring("tipo")
  l_grabado = request.querystring("grabado")
  l_evaobjnro = request("evaobjnro")
  l_evaobjcomnro  = request("evaobjcomnro")

'  response.write l_evaobjcom & "<br>"
 '  response.write l_evldrnro & "<br>"
  '  response.write l_grabado & "<br>"
	' response.write l_evaobjnro & "<br>"
	 ' response.write l_evaobjcomnro & "<br>"
  
  
'BODY ----------------------------------------------------------
if l_tipo ="A" then
	l_sql= "insert into evaobjcom (evaobjcom,evaobjnro,evldrnro) "
	l_sql = l_sql & " values ('" & trim(l_evaobjcom) & "'," & l_evaobjnro & "," & l_evldrnro &")"
else
	if l_tipo = "M" then
		l_sql = "UPDATE evaobjcom SET "
		l_sql = l_sql & " evaobjcom = '" & trim(l_evaobjcom) & "'"     

		l_sql = l_sql & " WHERE evldrnro  = "  & l_evldrnro
		l_sql = l_sql & " AND  evaobjnro = "  & l_evaobjnro
		l_sql = l_sql & " AND  evaobjcomnro ="  & l_evaobjcomnro
	else
		l_sql = "DELETE from evaobjcom "
		l_sql = l_sql & " where evldrnro  = "  & l_evldrnro
		l_sql = l_sql & " AND  evaobjnro  = "  & l_evaobjnro
		l_sql = l_sql & " AND  evaobjcomnro ="  & l_evaobjcomnro
	end if	
end if	
	set l_cm = Server.CreateObject("ADODB.Command")  
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0

response.write "<script> parent."&l_grabado&".value = '" &l_tipo &"'; </script>"
'response.write "<script> parent.location.reload(); </script>"
%>
