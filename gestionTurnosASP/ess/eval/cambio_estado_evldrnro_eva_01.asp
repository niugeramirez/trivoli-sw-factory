<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<%
'Archivo	: cambio_estado_evldrnro_eva_01.asp
'Descripción: cambio estados
'Autor		: CCRossi
'Fecha		: 11-06-2004
'Modificacion: 12-11-04 CCRossi habilitar el siguiente evaluador si termina una seccion
'-------------------------------------------------------------------------------
 Dim l_cm
 Dim l_sql
 Dim l_rs
 
'variables locales
 dim l_hora 
 dim l_arrhr
 dim l_ingresoant		
 dim l_habilitadoant		
 dim l_evldorcargadaant	
 
 dim l_evaseccnro
 
'parametros de entrada
 dim l_evldrnro 
 dim l_ingreso
 dim l_habilitado
 dim l_evldorcargada
   
l_evldrnro	= Request.QueryString("evldrnro")
l_ingreso	= Request.QueryString("ingreso")
l_habilitado	= Request.QueryString("habilitado")
l_evldorcargada	= Request.QueryString("evldorcargada")

Set l_rs = Server.CreateObject("ADODB.RecordSet")		
l_sql = "SELECT evaseccnro "
l_sql = l_sql & " FROM evadetevldor "
l_sql = l_sql & " WHERE evldrnro = " & l_evldrnro
rsOpen l_rs, cn, l_sql, 0 
if not  l_rs.eof then
	l_evaseccnro		= l_rs("evaseccnro")
end if
l_rs.close	
set l_rs=nothing

function strto2(cad)
	if trim(cad) <>"" then
		if len(cad)<2 then
			strto2= "0" & cad
		else
			strto2= cad
		end if 
	else
		strto2= "00"
	end if	
end function

' ------------------------------------------------------------------------------------
'											BODY 
' ------------------------------------------------------------------------------------

Set l_rs = Server.CreateObject("ADODB.RecordSet")		
l_sql = "SELECT ingreso,habilitado,evldorcargada,fechaing,fechahab,fechacar,horaing,horahab,horacar "
l_sql = l_sql & " FROM evadetevldor "
l_sql = l_sql & " WHERE evldrnro = " & l_evldrnro
rsOpen l_rs, cn, l_sql, 0 
if not  l_rs.eof then
	l_ingresoant		= l_rs("ingreso")
	l_habilitadoant		= l_rs("habilitado")
	l_evldorcargadaant	= l_rs("evldorcargada")
end if
l_rs.close	
set l_rs=nothing

if l_ingreso <> l_ingresoant or l_habilitado <> l_habilitadoant or l_evldorcargada <> l_evldorcargadaant  then
	
	l_hora = mid(time,1,8)
	l_arrhr= Split(l_hora,":")
	l_hora = strto2(l_arrhr(0))& l_arrhr(1)
	
	
if cint(l_ingreso) <> cint(l_ingresoant) or cint(l_habilitado) <> cint(l_habilitadoant) or cint(l_evldorcargada) <> cint(l_evldorcargadaant) then
	
	set l_cm = Server.CreateObject("ADODB.Command")
	l_sql = "UPDATE  evadetevldor SET "
	if cint(l_ingreso) <> cint(l_ingresoant) then
		l_sql = l_sql & " ingreso= "  & l_ingreso & ","
		if l_ingreso=0 then
		l_sql = l_sql & " fechaing  = NULL,"
		l_sql = l_sql & " horaing   = ''"
		else
		l_sql = l_sql & " fechaing  =   " & cambiafecha(Date(),"","") & ","
		l_sql = l_sql & " horaing   =	'" & l_hora & "'"
		end if	
		if cint(l_habilitado) <> cint(l_habilitadoant) or cint(l_evldorcargada) <> cint(l_evldorcargadaant) then
			l_sql = l_sql & ","
		end if
	end if	
	if cint(l_habilitado) <> cint(l_habilitadoant) then
		l_sql = l_sql & " habilitado= "  & l_habilitado & ","
		if l_habilitado=0 then
		l_sql = l_sql & " fechahab  = NULL,"
		l_sql = l_sql & " horahab   = ''"
		else
		l_sql = l_sql & " fechahab  =   " & cambiafecha(Date(),"","") & ","
		l_sql = l_sql & " horahab   =	'" & l_hora & "'"
		end if
		if cint(l_evldorcargada) <> cint(l_evldorcargadaant) then
			l_sql = l_sql & ","
		end if
	end if
	if cint(l_evldorcargada) <> cint(l_evldorcargadaant) then
		l_sql = l_sql & " evldorcargada= "  & l_evldorcargada & ","
		if l_evldorcargada=0 then
		l_sql = l_sql & " fechacar  = NULL,"
		l_sql = l_sql & " horacar   = ''"
		else
		l_sql = l_sql & " fechacar  =   " & cambiafecha(Date(),"","") & ","
		l_sql = l_sql & " horacar   =	'" & l_hora & "'"
		end if
	end if
	l_sql = l_sql & " WHERE evldrnro="& l_evldrnro
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
end if


	if cint(l_evldorcargada)=-1 then%>	
	<script>
//	alert('habilitar');
	abrirVentanaH('habilitar_evaluador_eva_ag_00.asp?evldrnro=<%=l_evldrnro%>&evldorcargada=-1&evaseccnro=<%=l_evaseccnro%>','',500,500);
	//var r = showModalDialog('habilitar_evaluador_eva_ag_00.asp?evldrnro=<%=l_evldrnro%>&evldorcargada=-1&evaseccnro=<%=l_evaseccnro%>', '','dialogWidth:20;dialogHeight:20'); 
	</script>
	<%end if

end if			


cn.close
Set cn = Nothing

response.write "<script>window.returnValue='0';window.close();</script>"
%>