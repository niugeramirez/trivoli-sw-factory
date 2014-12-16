<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/adovbs.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/asistente.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--
Archivo: ag_matriz_competencia_cap_05.asp
Descripción: 
Autor : Raul Chinestra

-->
<% 
'on error goto 0
'ADO
 Dim l_cm
 Dim l_sql
 Dim l_rs
 
' variables
 Dim l_grabar1
 Dim l_grabar2

 Dim l_ternro
 Dim l_entnro 
 Dim l_porcnro

 l_ternro   = Request.QueryString("ternro")
 l_grabar1   = Request.QueryString("grabar1")
 l_grabar2   = Request.QueryString("grabar2")

'uso local
 dim i
 Dim j
 j = 0
 dim l_lista
 
 'Al operar sobre varias tablas debo iniciar una transacción
 cn.BeginTrans
 
 'Creo los objetos ADO
 set l_rs = Server.CreateObject("ADODB.RecordSet")
 set l_cm = Server.CreateObject("ADODB.Command")

 l_lista= Split(l_grabar1,",")		
 i = 1
 do while i <= UBound(l_lista)-1
	l_entnro   = l_lista(i)
	l_porcnro =  l_lista(i+1)
	l_sql = "INSERT INTO cap_capacita "
	l_sql = l_sql & "(origen1, idnro1, origen2, entnro, porcen, fecha) "
	l_sql = l_sql & " VALUES (5," & l_ternro & ", 3 ," & l_entnro & "," & l_porcnro & "," & cambiafecha(date(),"YMD",true) & ")"
	'response.write l_sql
	'response.end
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	i = i + 2
    j = 1
 loop	

 l_lista= Split(l_grabar2,",")		
 i = 1
 do while i <= UBound(l_lista)-1
	l_entnro   = l_lista(i)
	l_porcnro =  l_lista(i+1)
	l_sql = "INSERT INTO cap_capacita "
	l_sql = l_sql & "(origen1, idnro1, origen2, entnro, porcen, fecha) "
	l_sql = l_sql & " VALUES (" & 5 & "," & l_ternro & ", 3 ," & l_entnro & "," & l_porcnro & "," & cambiafecha(date(),"YMD",true) &")"
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	i = i + 2
    j = 1
 loop	

 cn.CommitTrans
 
 Set cn = Nothing
 Set l_cm = Nothing
%>

 <script>
 if (opener.parent.RefrescarPasos){
    opener.parent.RefrescarPasos();
 }
 </script>
 
<%
 response.write "<script>alert('Operación Realizada.');window.opener.opener.ifrm.location.reload();window.opener.close();window.close();</script>"
%> 
