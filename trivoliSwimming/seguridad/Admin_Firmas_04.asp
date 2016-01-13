<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--#include virtual="/turnos/shared/inc/fecha.inc"-->
<% 
Dim l_rs
Dim l_sql
Dim l_tipo
Dim l_codigo
Dim l_Descripcion

l_tipo        = request("tipo")
l_codigo      = request("codigo")
l_Descripcion = request("descripcion") 

Set l_rs = Server.CreateObject("ADODB.RecordSet") 

l_sql = "select cystipprogweb from cystipo "
l_sql = l_sql & "where cystipnro = " & l_tipo

rsOpen l_rs, cn, l_sql, 0 

'if not l_rs.eof then
  response.redirect l_rs(0) & l_codigo
'else
'  response.write "<script>alert('Los comprobantes de tipo " & l_codigo & ", no tienen definodo un programa de visualizacion.');window.close();</script>"
'end if
'select Case l_tipo
'	Case "1":  'partes diarios de horas extras
'      response.redirect "cabecera_partes_02.asp?Tipo=C&gtpnro=4&gcpnro=" & l_codigo
	  
'    Case "5":  'Novedades por estructuras
'	  response.redirect "../liq/novedades_estructuras_liq_00.asp"
	  
'	Case "6":  'licencia 
'      response.redirect "licencias_gti_02.asp?tipo=C&cabnro=" & l_codigo

'	Case "11":  'partes diarios de movilidad
'      response.redirect "cabecera_partes_02.asp?Tipo=C&gtpnro=6&gcpnro=" & l_codigo
	  
'	Case "17":  'partes diarios de asignacion horaria
'      response.redirect "cabecera_partes_02.asp?Tipo=C&gtpnro=1&gcpnro=" & l_codigo

	  
'end select

%>
