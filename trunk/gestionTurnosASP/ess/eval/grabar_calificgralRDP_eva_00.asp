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
'Objetivo : grabar calificacion gral de objetivos de evaluacion RDP
'Fecha	  : 12-02-2005 
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
Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT * FROM evagralrdp "
l_sql = l_sql & " WHERE evldrnro  = " & l_evldrnro
rsOpen l_rs1, cn, l_sql, 0
if  l_rs1.eof then 
	l_sql= "INSERT INTO evagralrdp (evldrnro,evatrnro) "
	l_sql = l_sql & " VALUES (" & l_evldrnro & "," & l_evatrnro  &")"
else 
	l_sql = "UPDATE evagralrdp SET "
	l_sql = l_sql & " evatrnro   = " & l_evatrnro
	l_sql = l_sql & " WHERE evldrnro = "  & l_evldrnro
end if 
l_rs1.Close 
set l_rs1=nothing 
' Response.Write l_sql 

set l_cm = Server.CreateObject("ADODB.Command")  
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0

response.write "<script> parent.location.reload(); </script>"
%>
