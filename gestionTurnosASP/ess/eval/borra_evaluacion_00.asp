<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<% 
'Modificacion: 10-05-2004 - CCROssi sacar alerts.
'Modificacion: 20-05-2004 - CCROssi agrgagar parametro de viene de relacionar empleados.
%>
<body>
<head>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script>
function Nuevo_Dialogo(w_in, pagina, ancho, alto)
{
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}
window.resizeTo(5,5);
</script>
</head>
<% 
'de base de datos
 Dim l_cm
 Dim l_sql
 Dim l_rs

'de uso local
 dim l_tipsecprogdel 

'parametro entrada
 Dim l_evaevenro
 Dim l_empleado
 dim l_llamadora
 
l_llamadora = Request.QueryString("llamadora") 
l_evaevenro  = Request.QueryString("evaevenro")
l_empleado   = Request.QueryString("empleado")

' ================================================================================
' BODY
' ================================================================================
on error goto 0
cn.BeginTrans

Set l_rs = Server.CreateObject("ADODB.RecordSet") 
l_sql = "SELECT evasecc.orden, evasecc.evaseccnro, evatiposecc.tipsecprogdel  "
l_sql = l_sql & " FROM evadet "
l_sql = l_sql & " INNER JOIN evasecc     ON evadet.evaseccnro=evasecc.evaseccnro "
l_sql = l_sql & " INNER JOIN evatiposecc ON evasecc.tipsecnro=evatiposecc.tipsecnro "
l_sql = l_sql & " INNER JOIN evacab      ON evacab.evacabnro=evadet.evacabnro "
l_sql = l_sql & " WHERE evacab.evaevenro = " & l_evaevenro
l_sql = l_sql & " AND   evacab.empleado  = " & l_empleado
l_sql = l_sql & " GROUP BY evasecc.orden, evasecc.evaseccnro, evatiposecc.tipsecprogdel , evacab.evacabnro "
l_sql = l_sql & " ORDER BY orden "
rsOpen l_rs, cn, l_sql, 0 

do until l_rs.eof
    l_tipsecprogdel = l_rs("tipsecprogdel")
    if len(trim(l_tipsecprogdel)) <> 0  then
	  ' correr el programa correspondiente%>
	  <script>
		Nuevo_Dialogo(window, '<%=l_tipsecprogdel%>?evaevenro=<%=l_evaevenro%>&empleado=<%=l_empleado%>', 5,5);
	  </script>
	<% 
	end if
	l_rs.MoveNext
loop
l_rs.Close


if err.number > 0 then
	cn.RollbackTrans
	Response.write "<script>alert('Error.La evaluación no se eliminó.');window.close();</script>"	
else
	cn.CommitTrans%>
	<script>
		Nuevo_Dialogo(window, 'borra_evaluacion_detallecabecera_00.asp?evaevenro=<%=l_evaevenro%>&empleado=<%=l_empleado%>', 5,5);
	</script>
<%end if	

if l_llamadora<>"relacionar" then
	Response.write "<script>alert('Operación Realizada.');</script>"	
	Response.write "<script>opener.Volver_primero();</script>"	
end if
Response.write "<script>window.close();</script>"	   

%>
</body>