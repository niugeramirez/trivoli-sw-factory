<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/adovbs.inc"-->
<!--#include file="vacaciones_calculo_gti.asp"-->
<script>
var xc = screen.availWidth;
var yc = screen.availHeight;
window.moveTo(xc,yc);	
//window.resizeTo(20,20);</script>
<% 
'---------------------------------------------------------------------------------------
'Archivo	: pedido_emp_vac_gti_03.asp
'Descripción: Pedido Vacaciones
'Autor: Scarpa D.
'Fecha: 08/10/2004
'Modificado:
'---------------------------------------------------------------------------------------
on error goto 0

Dim l_tipo
Dim l_cm
Dim l_sql
Dim l_rs
Dim l_rs2

dim l_vacnro
dim l_ternro
dim l_vdiapednro
dim l_vdiapeddesde
dim l_vdiapedhasta
dim l_vdiascorcant	
dim l_vdiapedcant
dim l_vdiaspedhabiles
dim l_vdiaspednohabiles
dim l_vdiaspedferiados  
dim l_vdiaspedestado

l_tipo			= Request.Form("tipo")

l_vacnro  		= request.Form("vacnro")
l_vdiapednro	= request.Form("vdiapednro")

l_vdiapeddesde	= request.Form("vdiapeddesde")
l_vdiapedhasta	= request.Form("vdiapedhasta")

l_vdiapedcant		= request.Form("vdiapedcant")
l_vdiaspedhabiles	= request.Form("vdiaspedhabiles")
l_vdiaspednohabiles	= request.Form("vdiaspednohabiles")
l_vdiaspedferiados	= request.Form("vdiaspedferiados")

l_vdiaspedestado = 0

'------------------------------------------------------------------------------------------------------
' SUB: guardarDatos
'------------------------------------------------------------------------------------------------------

sub guardarDatos(vdiapeddesde,vdiapedhasta,vdiapedcant,ternro,vdiaspedestado,vacnro,vdiaspedferiados,vdiaspedhabiles,vdiaspednohabiles)

    Dim l_vpeddesde
    Dim l_vpedhasta

	l_vpeddesde	= cambiaFecha(vdiapeddesde, "YMD", true)
    l_vpedhasta	= cambiaFecha(vdiapedhasta, "YMD", true)

	l_sql = "INSERT INTO vacdiasped "
	l_sql = l_sql & "(vdiapeddesde, vdiapedhasta, vdiapedcant, ternro, vdiaspedestado, "
	l_sql = l_sql & " vacnro, vdiaspedferiados, vdiaspedhabiles,vdiaspednohabiles ) "
	l_sql = l_sql & " values (" & l_vpeddesde  & "," & l_vpedhasta  & ", "
	l_sql = l_sql & ( CInt(vdiapedcant) + CInt(vdiaspednohabiles) + CInt(vdiaspedferiados) )  & "," & ternro &  ","  & vdiaspedestado & "," 
	l_sql = l_sql & vacnro          & "," & vdiaspedferiados & ","
	l_sql = l_sql & vdiaspedhabiles & "," & vdiaspednohabiles & ")"

'	response.write l_sql

	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	
'	l_sql = "UPDATE vacdiasped SET "
'	l_sql = l_sql & " vdiapeddesde =  " & l_vdiapeddesde & " ," 
'	l_sql = l_sql & " vdiapedhasta =  " & l_vdiapedhasta & " ," 
'	l_sql = l_sql & " vdiapedcant  =  " & l_vdiapedcant  & " ," 
'	l_sql = l_sql & " vacnro		= " & l_vacnro       & " ,"
'	l_sql = l_sql & " vdiaspedferiados = " & l_vdiaspedferiados & ","	
'	l_sql = l_sql & " vdiaspedhabiles  = " & l_vdiaspedhabiles  & ","	
'	l_sql = l_sql & " vdiaspednohabiles  = " & l_vdiaspednohabiles  & ","
'	l_sql = l_sql & " vdiaspedestado   = " & l_vdiaspedestado  
'	l_sql = l_sql & " WHERE vacdiasped.ternro	   = " & l_ternro
'	l_sql = l_sql & " AND   vacdiasped.vdiapednro  = " & l_vdiapednro

end sub 'guardarDatos()

'cn.beginTrans

Set l_rs  = Server.CreateObject("ADODB.RecordSet")
Set l_rs2 = Server.CreateObject("ADODB.RecordSet")
set l_cm = Server.CreateObject("ADODB.Command")

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


	if l_tipo = "M" then 
	    'Borro el pedido actual
		l_sql = "DELETE FROM vacdiasped "
		l_sql = l_sql & " WHERE vacdiasped.ternro	   = " & l_ternro
		l_sql = l_sql & " AND   vacdiasped.vdiapednro  = " & l_vdiapednro
	
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
	end if

	Dim totalCantidad
	Dim corresp
	Dim pedidos
	Dim cantidad
	Dim total2
	Dim totalFer2
	Dim cant
	Dim desde
	Dim hasta
	
    Dim m_rs 
	
	set m_rs  = Server.CreateObject("ADODB.RecordSet")
	
	desde = l_vdiapeddesde
	hasta = l_vdiapedhasta
	cant  = CInt(l_vdiapedcant)
	
	'Busco la cant. de dias corresp.
	l_sql = "SELECT vacdiascor.vacnro, vacdiascor.tipvacnro, "
	l_sql = l_sql & " vacdiascor.vdiascorcant "
	l_sql = l_sql & " FROM  vacdiascor "
	l_sql = l_sql & " INNER JOIN vacacion ON vacacion.vacnro = vacdiascor.vacnro "
	l_sql = l_sql & " WHERE vacdiascor.ternro =  " & l_ternro
	l_sql = l_sql & "   AND vacfecdesde <=  " & cambiafecha(desde,"YMD",true)
	l_sql = l_sql & "   AND vacfechasta >=  " & cambiafecha(desde,"YMD",true)
	l_sql = l_sql & " ORDER BY vacfecdesde ASC "
	
	rsOpen m_rs, cn, l_sql, 0 
	
	do until m_rs.eof 
	
	   call iniTipoVac(m_rs("tipvacnro"))
	
	   call datosPedVac(l_ternro,m_rs("vacnro"), corresp, pedidos, l_vdiapednro)
	   
	   cantidad = CInt(corresp) - CInt(pedidos)

	   if CInt(cant) <= CInt(cantidad) then
	      call busqFecha(desde,cant,hasta,total2,totalFer2)
		  call guardarDatos(desde,hasta,cant,l_ternro,l_vdiaspedestado,m_rs("vacnro"),totalFer2,cant,(CInt(total2) - CInt(cant)))
		  exit do
	   else
	      if CInt(cantidad) > 0 then
		      call busqFecha(desde,cantidad,hasta,total2,totalFer2)
			  cant     = cant - cantidad
			  call guardarDatos(desde,hasta,cantidad,l_ternro,l_vdiaspedestado,m_rs("vacnro"),totalFer2,cantidad,(CInt(total2) - CInt(cantidad)))
			  desde    = DateAdd("d",1,CDate(hasta))
		  end if
	   end if

	   m_rs.movenext 
	loop
	
	m_rs.close
	
'cn.commitTrans

Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location.reload();window.close();</script>"

%>
