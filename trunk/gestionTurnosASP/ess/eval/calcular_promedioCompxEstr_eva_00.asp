<%Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<script>
var xc = screen.availWidth;
var yc = screen.availHeight;
window.moveTo(xc,yc);	
</script>
<%
on error goto 0
'=====================================================================================
'Archivo  : calcular_promediocompxEstr_eva_00.asp
'Objetivo : calcular promedio competencias que NO sean potenciales
'Fecha	  : 03-11-2005
'Autor	  : CCRossi
'Modificacion:  - LA - arreglo de calculos
'=====================================================================================

dim l_Error
Dim l_rs_oblig
dim l_rs
dim l_sql

dim l_suma
dim l_promedio

'parametros de entrada
dim l_evldrnro
dim l_cantidad
dim l_usarporcen

l_evldrnro   = Request.QueryString("evldrnro")
'l_cantidad   = Request.QueryString("cantidad")
l_usarporcen = Request.QueryString("usarporcen")

Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_Error = 0

'___________________________________________________________________________________
function PasarComaAPunto(valor)
	dim l_numero
	dim l_ubicacion
	dim l_entero
	dim l_decimal
	l_numero = trim(valor)
	l_ubicacion = InStr(l_numero, ",")
	if l_ubicacion > 1 then
		l_ubicacion = l_ubicacion  - 1
		l_entero = left(l_numero, l_ubicacion)
		l_ubicacion = l_ubicacion  + 1
		l_decimal = right(l_numero, (len(l_numero) - l_ubicacion))
    	l_numero = l_entero & "." & l_decimal
    	PasarComaAPunto = l_numero
    else
		PasarComaAPunto = valor
	end if
end function	

'=========================================================================================
' Calcular Suma Total

' si usa porcentaje usar esta suma , sino la otra..

l_suma = 0
if cint(l_usarporcen)=-1 then
	 Set l_rs = Server.CreateObject("ADODB.RecordSet")
	 l_sql = " SELECT SUM(evarestot) AS suma "
	 l_sql = l_sql & " FROM  evaresultado "
	 l_sql = l_sql & " INNER JOIN evatipresu    ON evatipresu.evatrnro = evaresultado.evatrnro   "
	 l_sql = l_sql & " INNER JOIN evafactor		ON evafactor.evafacnro = evaresultado.evafacnro "
	 l_sql = l_sql & "		AND evafactor.evafacpot = 0 "
	 l_sql = l_sql & " WHERE evaresultado.evldrnro    = " & l_evldrnro
	 rsOpen l_rs, cn, l_sql, 0
	 if not l_rs.eof then 
		l_suma = l_rs("suma")
	 end if
	 l_rs.Close
else
	 Set l_rs = Server.CreateObject("ADODB.RecordSet")
	 l_sql = " SELECT SUM(evatrvalor) AS suma "
	 l_sql = l_sql & " FROM  evaresultado "
	 l_sql = l_sql & " INNER JOIN evatipresu    ON evatipresu.evatrnro = evaresultado.evatrnro   "
	 l_sql = l_sql & " INNER JOIN evafactor		ON evafactor.evafacnro = evaresultado.evafacnro "
	 l_sql = l_sql & "		AND evafactor.evafacpot = 0 "
	 l_sql = l_sql & " WHERE evaresultado.evldrnro    = " & l_evldrnro
	 rsOpen l_rs, cn, l_sql, 0
	 if not l_rs.eof then 
		l_suma = l_rs("suma")
	 end if
	 l_rs.Close
	
end if
'response.write "<script> alert('"& l_usarporcen &"')</script>"
'response.write "<script> alert('"& l_suma &"')</script>"
 
' Contar Cantidad de Competencias Evaluadas
l_sql = "SELECT COUNT(evaresultado.evafacnro) AS cantidad "
l_sql = l_sql & " FROM evaresultado "
l_sql = l_sql & " INNER JOIN evafactor ON evafactor.evafacnro = evaresultado.evafacnro AND evafactor.evafacpot = 0 "
l_sql = l_sql & " INNER JOIN evatipresu ON evatipresu.evatrnro = evaresultado.evatrnro "
l_sql = l_sql & " WHERE evldrnro =" & l_evldrnro

'response.write "<script> alert('"& l_sql &"')</script>"
  'l_sql = "SELECT COUNT(evadescomp.evafacnro) AS cantidad "
  'l_sql = l_sql & " FROM evadescomp "
  'l_sql = l_sql & " INNER JOIN evafactor ON evafactor.evafacnro = evadescomp.evafacnro "
  'l_sql = l_sql & "		AND evafactor.evafacpot = 0 "
  'l_sql = l_sql & " WHERE EXISTS "
  'l_sql = l_sql & " (SELECT * FROM evaresultado "
  'l_sql = l_sql & " INNER JOIN evatipresu ON evatipresu.evatrnro = evaresultado.evatrnro "
  'l_sql = l_sql & " AND evatipresu.evatrnro IS NOT NULL "
  'l_sql = l_sql & " WHERE evaresultado.evafacnro = evadescomp.evafacnro "
  'l_sql = l_sql & " AND   evaresultado.evldrnro = " & l_evldrnro
  'l_sql = l_sql & " AND evadescomp.estrnro IN ("& l_estrnros & ")"
  'l_sql = l_sql & ")"
  Set l_rs = Server.CreateObject("ADODB.RecordSet")
  rsOpen l_rs, cn, l_sql, 0
  if not l_rs.eof then 
	l_cantidad = l_rs("cantidad")
  end if
  l_rs.Close

 'response.write "<script> alert('"& l_cantidad &"')</script>"
 
if isnull(l_suma) then
l_suma=0
end if

l_promedio = 0
  
if cint(l_usarporcen)=-1 then
	if trim(l_suma)<>"" and not isnull(l_suma) then
		l_promedio = cdbl(l_suma)
	else
		l_promedio = 0
	end if
else
   if trim(l_cantidad)<>"" and not isnull(l_cantidad) and cint(l_cantidad)<>0 then
		if trim(l_suma)<>"" and not isnull(l_suma) then
			l_promedio = cdbl(l_suma) / cdbl(l_cantidad)
		else
			l_promedio = 0
		end if
   end if
end if
%>
	<script>
		
		window.returnValue='<%= PasarComaAPunto(round(l_promedio,2))%>,<%= PasarComaAPunto(round(l_suma,2))%>';
		window.close();
	</script>



