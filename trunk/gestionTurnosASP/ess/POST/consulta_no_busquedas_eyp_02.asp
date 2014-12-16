<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!-----------------------------------------------------------------------------------------------
Archivo		: consulta_busquedas_eyp_03.asp
Descripción	: Inserta al tercero en la busqueda seleccionada.
Autor 		: Liosandro moro
Fecha		: 27/05/2004
-------------------------------------------------------------------------------------------------
-->
<% 
 Dim l_rs
 Dim l_cm
 Dim l_sql
 Dim l_orden
 
 Dim l_reqbusnro
 Dim l_ternro
 
 l_reqbusnro = request.QueryString("reqbusnro")
 l_ternro = request.QueryString("ternro")

set l_cm = Server.CreateObject("ADODB.Command")

l_sql = "INSERT INTO pos_terreqbus "
l_sql = l_sql & "(ternro, reqbusnro, conf) "
l_sql = l_sql & "VALUES (" & l_ternro & "," & l_reqbusnro & ", 0)"

'	l_sql = "UPDATE pos_circuito "
'	l_sql = l_sql & "SET cirdesabr = '" & l_cirdesabr & "'"
'	l_sql = l_sql & ",cirdesext = '" & l_cirdesext & "'"
'	l_sql = l_sql & " WHERE cirnro = " & l_cirnro

'response.write l_sql & "<br>"
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location.reload();window.opener.document.all.textsql.value='';window.close();</script>"
%>
