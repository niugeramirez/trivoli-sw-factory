<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/adovbs.inc"-->
<!--#include file="vacaciones_calculo_gti.asp"-->
<%
'Archivo	: pedido_emp_vac_calculol_gti_00
'Descripción: contar dias habiles y feriados
'Autor		: Scarpa D.
'Fecha		: 08/10/2004
'Modificado	: 
on error goto 0

Dim l_rs
Dim l_rs2
Dim l_sql

Dim l_errores
Dim l_ternro
Dim l_desde
Dim l_hasta
Dim l_cant
Dim l_tipoVac
Dim l_tipo
Dim l_vdiapednro

'locales
Dim i
Dim factual
Dim l_hasta_str
Dim l_pais
Dim l_totalFer
Dim l_total

dim l_dia1
dim l_dia2
dim l_dia3
dim l_dia4
dim l_dia5
dim l_dia6
dim l_dia7
dim l_excFer

Dim l_totalCantidad
Dim l_totalCorr

l_desde      = request("desde")
l_hasta      = request("hasta")
l_cant       = request("cantidad")
l_tipoVac    = request("tipovac")
l_tipo       = request("tipo")
l_vdiapednro = request("vdiapednro")

'---------------------------------------------------------------------------------------------------------
' EMPIEZA EL MODULO
'---------------------------------------------------------------------------------------------------------

'Busco cual es el pais default
Set l_rs  = Server.CreateObject("ADODB.RecordSet")
Set l_rs2 = Server.CreateObject("ADODB.RecordSet")

dim leg
leg = Session("empleg")
if leg = "" then
    response.write "NO SE HA SELECCIONADO UN EMPLEADO<BR>"
	Response.End
end if

l_sql = "SELECT ternro FROM empleado WHERE empleado.empleg = " & leg
l_rs.Open l_sql, cn
if l_rs.eof then
    response.write "NO SE HA SELECCIONADO UN EMPLEADO<BR>"
	response.end
else 
  l_ternro = l_rs("ternro")
end if
l_rs.close


'Calculo el rango de fecha

l_errores = 0

if l_tipo = "SD" then

  if Trim(l_cant) <> "" then
	  'Suma dias a una fecha
	  'call busqFecha(l_desde,l_cant,l_hasta,l_total, l_totalFer)  
	  call busqFechaPedidos(l_ternro, l_desde, l_cant, l_hasta, l_total, l_totalFer, l_vdiapednro, l_errores)
  else
      l_hasta    = ""
	  l_total    = 0
	  l_totalFer = 0
	  l_cant     = ""
  end if

else

  if Trim(l_hasta) <> "" then
      'Cantidad de dias entre dos fecha
      'call cantDias(l_desde,l_hasta,l_cant,l_total,l_totalFer)
	  call cantDiasPedidos(l_ternro, l_desde, l_hasta, l_cant, l_total, l_totalFer, l_vdiapednro, l_errores)
  else
      l_hasta    = ""
	  l_total    = 0
	  l_totalFer = 0
	  l_cant     = ""
  end if

end if  

l_totalCantidad = 0
l_totalCorr     = 0

call cantidadDiasPedDisp(l_ternro,l_desde,l_totalCantidad, l_totalCorr, l_vdiapednro)	  

if l_errores = 0 then
%>
<script>
  parent.actualizarTotales(<%= l_totalCantidad%>,<%= l_totalCorr%>);  
  parent.actualizarRango('<%= l_hasta%>','<%= l_cant%>','<%= l_total%>','<%= l_totalFer%>');
</script>
<%else%>
<script>
  parent.actualizarTotales(<%= l_totalCantidad%>,<%= l_totalCorr%>);    
  parent.mostrarErrores('<%= l_errores%>');
</script>
<%end if%>
