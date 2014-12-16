<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : noved_horarias_gti_04.asp
Descripcion    : Modulo que se encarga de eliminar los datos de las nov horarias
Modificacion   :
    06/10/2003 - Scarpa D. - Punto de procesamiento
	07/10/2005- Leticia A. - 
-----------------------------------------------------------------------------
-->
<html>
<head>
</head>
<body>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<% 
Dim l_cm
Dim l_sql
Dim l_emp_licnro
Dim l_datos
Dim l_ternro
Dim l_gnovnro

Dim l_fechadesde
Dim l_fechahasta

Dim l_desde
Dim l_hasta

l_fechadesde = request.querystring("fechadesde")
l_fechahasta  = request.querystring("fechahasta")

l_gnovnro = request.querystring("cabnro")
l_ternro = request.querystring("ternro")

Dim l_tipAutorizacion  'Es el tipo del circuito de firmas
Dim l_HayAutorizacion  'Es para ver si las autorizaciones estan activas
Dim l_PuedeVer         'Es para ver si las autorizaciones estan activas
Dim l_rs

Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_tipAutorizacion = 7  'Es del tipo novedades

l_sql = "select * from cystipo "
l_sql = l_sql & "where (cystipo.cystipact = -1) and cystipo.cystipnro = " & l_tipAutorizacion 

rsOpen l_rs, cn, l_sql, 0 

l_HayAutorizacion = not l_rs.eof

l_rs.close

'Busco cual es el rango de fechas de la novedad
l_sql = " SELECT * FROM gti_novedad " 
l_sql = l_sql & "WHERE gnovnro = " & l_gnovnro

rsOpen l_rs, cn, l_sql, 0 

if not l_rs.eof then
  l_desde = l_rs("gnovdesde")
  l_hasta = l_rs("gnovhasta")
end if

'Controlo la autorizacion
if l_HayAutorizacion then

  l_sql = "select cysfirautoriza, cysfirsecuencia, cysfirdestino from cysfirmas "
  l_sql = l_sql & "where cysfirmas.cystipnro = " & l_tipAutorizacion & " and cysfirmas.cysfircodext = '" & l_gnovnro & "' " 
  l_sql = l_sql & "order by cysfirsecuencia desc"

  rsOpen l_rs, cn, l_sql, 0 

  l_PuedeVer = False

  if not l_rs.eof then
    if (l_rs("cysfirautoriza") = session("UserName")) or (l_rs("cysfirdestino") = session("UserName")) then 
	   'Es una modificación del ultimo o es el nuevo que autoriza 
       l_PuedeVer = True 
    end if
  end if
  l_rs.close
  If not l_PuedeVer then
    response.write "<script>alert('No esta autorizado a ver o modificar este registro.');window.close()</script>"
	response.end
  End if
End if

cn.beginTrans

set l_cm = Server.CreateObject("ADODB.Command")
l_sql = "DELETE FROM gti_novedad " 
l_sql = l_sql & "WHERE gnovnro = " & l_gnovnro
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0

l_sql = "DELETE FROM gti_justificacion WHERE jussigla = 'NOV' and juscodext = " & l_gnovnro
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
l_cm.close

l_sql = "DELETE FROM cysfirmas where cystipnro = 7 and cysfircodext = '" & l_gnovnro & "' "
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
l_cm.close

cn.CommitTrans

cn.Close
Set cn = Nothing

%>

<script>
  window.opener.ifrm.location.reload();
  alert('Operación Realizada.');  
  window.close();
</script>

</body>
</html>

