<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo	: graba_puntuacion_eva_00.asp
'Descripción: grabar puntuacion automatica de objetivos smart
'Autor		: CCRossi
'Fecha		: 28-07-2004
'Modificacion: 
'Modificacion: 02-11-2004 CCRossi -  si la puntuacioon viene nula o vacia no grabar puntuacion.
'Modificacion: 04-11-2004 CCRossi -  si es ABN grabar solo si es EVALUADOR
'Modificacion: 09-11-2004 CCRossi -  poner NULL en puntajemanual en lugar de cero.
'-------------------------------------------------------------------------------
 Dim l_cm
 Dim l_sql
 Dim l_rs
 Dim l_evacabnro 

 dim l_puntajemanual 
 dim l_evatevnro
'parametros de entrada
 dim l_evldrnro 
 dim l_puntaje
   
l_evldrnro = Request.QueryString("evldrnro")
l_puntaje  = Request.QueryString("puntaje")

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evacabnro,evatevnro, evatevnro "
l_sql = l_sql & " FROM  evadetevldor "
l_sql = l_sql & " WHERE evldrnro = " & l_evldrnro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.EOF then
	l_evacabnro = l_rs("evacabnro")
	l_evatevnro = l_rs("evatevnro")
end if
l_rs.close
set l_rs=nothing

' ------------------------------------------------------------------------------------
'											BODY 
' ------------------------------------------------------------------------------------

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT puntajemanual "
l_sql = l_sql & " FROM  evacab "
l_sql = l_sql & " WHERE evacabnro = " & l_evacabnro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.EOF then
	l_puntajemanual= l_rs("puntajemanual")
end if
l_rs.close
set l_rs=nothing

if (cejemplo=-1 and (l_evatevnro<>cautoevaluador)) or (ccodelco=-1)then
	if trim(l_puntaje)<>"" and not IsNull(l_puntaje) then
		set l_cm = Server.CreateObject("ADODB.Command")
		l_sql = "UPDATE evacab SET "
		l_sql = l_sql & " puntaje		  = " & l_puntaje & ","
		if trim(l_puntajemanual)="" or isnull(l_puntajemanual) then
		l_sql = l_sql & " puntajemanual	  = " & l_puntaje & ","
		end if
		l_sql = l_sql & " puntajeevldrnro = " & l_evldrnro
		l_sql = l_sql & " WHERE evacabnro = " & l_evacabnro
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
	end if
end if

cn.close
Set cn = Nothing

%>
<script>window.close();</script>


