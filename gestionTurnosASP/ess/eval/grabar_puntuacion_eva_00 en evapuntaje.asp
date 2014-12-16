<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo	: graba_puntuacion_eva_00.asp
'Descripción: grabar puntuacion automatica de objetivos smart
'Autor		: CCRossi
'Fecha		: 22-06-2004
'Modificacion: 
'-------------------------------------------------------------------------------
 Dim l_cm
 Dim l_sql
 Dim l_rs
 Dim l_evacabnro 
 Dim l_evatevnro 
 
'parametros de entrada
 dim l_evldrnro 
 dim l_puntaje
   
l_evldrnro	 = Request.QueryString("evldrnro")
l_puntaje = Request.QueryString("puntaje")

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evacabnro,evatevnro "
l_sql = l_sql & " FROM  evadetevldor "
l_sql = l_sql & " WHERE evldrnro = " & l_evldrnro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.EOF then
	l_evacabnro= l_rs("evacabnro")
	l_evatevnro= l_rs("evatevnro")
end if
l_rs.close
set l_rs=nothing

' ------------------------------------------------------------------------------------
'											BODY 
' ------------------------------------------------------------------------------------

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evacabnro,evatevnro "
l_sql = l_sql & " FROM  evapuntaje"
l_sql = l_sql & " WHERE evacabnro = " & l_evacabnro
l_sql = l_sql & " AND   evatevnro = " & l_evatevnro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.EOF then
	l_rs.close
	set l_rs=nothing
	
	set l_cm = Server.CreateObject("ADODB.Command")
	l_sql = "UPDATE evapuntaje SET "
	l_sql = l_sql & " puntaje = " & l_puntaje
	l_sql = l_sql & " WHERE evacabnro = " & l_evacabnro
	l_sql = l_sql & " AND   evatevnro = " & l_evatevnro
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
else
	l_rs.close
	set l_rs=nothing
	
	set l_cm = Server.CreateObject("ADODB.Command")
	l_sql = "INSERT INTO evapuntaje (evacabnro, evatevnro, puntaje) "
	l_sql = l_sql &  "VALUES ("& l_evacabnro &","& l_evatevnro &","& l_puntaje &")"
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
end if		
Set cn = Nothing
Set l_cm = Nothing
%>
<script>window.close();</script>


