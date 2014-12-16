<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<% 
'=====================================================================================
'Archivo  : grabar_calificobjRDP_eva_00.asp 
'Objetivo : grabar calificacion de competencias de evaluacion RDP 
'Fecha	  : 17-02-2005 
'Autor	  : Leticia Amadio
'=====================================================================================

' variables 
' parametros de entrada ----------------------------------------
  Dim l_evldrnro 
  Dim l_evatrnro 
  Dim l_tipo     

' parametros de entrada
  l_evldrnro	 = request.querystring("evldrnro")
  l_evatrnro	 = request.querystring("evatrnro")
  l_tipo		 = request.querystring("tipo")
  
' variables de base de datos ------------------------------------
  Dim l_cm
  Dim l_sql
  Dim l_rs
  Dim l_rs1
  
  
'BODY ----------------------------------------------------------

' uso grabar areas y competemcias !!!!

response.write "<script> parent.location.reload(); </script>"
	
%>
