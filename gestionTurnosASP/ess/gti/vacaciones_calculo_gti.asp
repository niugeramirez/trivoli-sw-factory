<%
'Archivo	: vacaciones_calculo_lic_gti_00.asp
'Descripción: contar dias habiles y feriados
'Autor		: Scarpa D.
'Fecha		: 21/09/2004
'Modificado	: 

'locales
Dim m_i
Dim m_factual
Dim m_hasta_str
Dim m_pais

dim m_dia1
dim m_dia2
dim m_dia3
dim m_dia4
dim m_dia5
dim m_dia6
dim m_dia7
dim m_excFer


'---------------------------------------------------------------------------------------------------------
' FUNCION: esFeriado - calcula si una fecha es feriado
'---------------------------------------------------------------------------------------------------------
function esFeriado(dia,pais)

  Dim l_salida
  
  l_salida = false

  l_sql =         " SELECT * FROM feriado "
  l_sql = l_sql & " WHERE feriado.ferifecha = " & cambiafecha(dia,"","")
  
  rsOpenCursor l_rs2, cn, l_sql, 0, adOpenDynamic
  
  if not l_rs2.eof then
	 l_salida = ((CInt(l_rs2("tipferinro")) = 1) AND (CInt(l_rs2("fericodext")) = CInt(pais) ) )
  end if
  l_rs2.Close
	
  esFeriado = l_salida
end function 'esFeriado(dia)

'---------------------------------------------------------------------------------------------------------
' SUB: iniTipoVac(tipoVac) se encarga de inicializar las variables de dias de acuerdo al tipo de vacacion
'---------------------------------------------------------------------------------------------------------
sub iniTipoVac(tipoVac)

	'Busco el pais por default
	l_sql =         " SELECT * FROM pais "
	l_sql = l_sql & " WHERE pais.paisdef = -1 " 
	
	rsOpenCursor l_rs, cn, l_sql, 0, adOpenDynamic
	
	if not l_rs.eof then
	   m_pais = l_rs("paisnro") 
	else
	   m_pais = -1
	end if
	
	l_rs.close
	
	'Busco la conf del tipo de vacacion
	
	l_sql =         " SELECT * FROM tipovacac "
	l_sql = l_sql & " WHERE tipvacnro = " & tipoVac
	
	rsOpenCursor l_rs, cn, l_sql, 0, adOpenDynamic 
	
	if not l_rs.eof then
	   m_dia1 = (CInt(l_rs("tpvhabiles__1")) = -1 )
	   m_dia2 = (CInt(l_rs("tpvhabiles__2")) = -1 )
	   m_dia3 = (CInt(l_rs("tpvhabiles__3")) = -1 )
	   m_dia4 = (CInt(l_rs("tpvhabiles__4")) = -1 )
	   m_dia5 = (CInt(l_rs("tpvhabiles__5")) = -1 )
	   m_dia6 = (CInt(l_rs("tpvhabiles__6")) = -1 )
	   m_dia7 = (CInt(l_rs("tpvhabiles__7")) = -1 )
	   m_excFer  = (CInt(l_rs("tpvferiado")) = -1 )
	else
	   m_dia1 = true
	   m_dia2 = true
	   m_dia3 = true
	   m_dia4 = true
	   m_dia5 = true
	   m_dia6 = true
	   m_dia7 = true
	   m_excFer  = true
	end if
	
	l_rs.close

end sub 'iniTipoVac(tipoVac)


'---------------------------------------------------------------------------------------------------------
' SUB: cantDias(desde,hasta) se encarga buscar la cant. de dias tomados de vacaciones
'---------------------------------------------------------------------------------------------------------
sub cantDias(desde,hasta, byRef cantidad, byRef total, byRef totalFer)

'Calculo el rango de fecha
Dim l_contador

   total = 0
   totalFer = 0

  'Cantidad de dias entre dos fecha
   m_i = 0
   m_factual = CDate(desde)
   l_contador = DateDiff("d",m_factual,hasta)
   do 
	 total = total + 1	 
   	 if (NOT m_excFer) AND esFeriado(m_factual,m_pais) then	 
	    'es un feriado y no hay que excluirlo
        totalFer = totalFer + 1	   
	 else
	    'Dependiendo si hay que considerar el dia o no incremento i
        select case WeekDay(m_factual)
		  case 1
		    if m_dia1 then
			   m_i = m_i + 1
			end if
		  case 2
		    if m_dia2 then
			   m_i = m_i + 1
			end if
		  case 3
		    if m_dia3 then
			   m_i = m_i + 1
			end if
		  case 4
		    if m_dia4 then
			   m_i = m_i + 1
			end if
		  case 5
		    if m_dia5 then
			   m_i = m_i + 1
			end if
		  case 6
		    if m_dia6 then
			   m_i = m_i + 1
			end if
		  case 7
		    if m_dia7 then
			   m_i = m_i + 1
			end if
		end select
	  end if
 	 m_factual = DateAdd("d", 1, m_factual)   
	 l_contador = l_contador - 1
   loop while l_contador >= 0
   
   cantidad = m_i

end sub 'cantDias(desde,hasta, byRef cantidad, byRef total, byRef totalFer)


' SUB: busqFecha(desde,cant) se encarga buscar la fecha de fin a partir de un dia
'---------------------------------------------------------------------------------------------------------
sub busqFecha(desde,cant,byRef hasta, byRef total, byRef totalFer)

  'Suma dias a una fecha
  
   total = 0
   totalFer = 0
   m_i = 0
   m_factual = CDate(desde)
   do 
     total = total + 1   
   	 if (NOT m_excFer) AND esFeriado(m_factual,m_pais) then	 
	    'es un feriado y no hay que excluirlo
        totalFer = totalFer + 1	   
	 else
	    'Dependiendo si hay que considerar el dia o no incremento i
        select case WeekDay(m_factual)
		  case 1
		    if m_dia1 then
			   m_i = m_i + 1
			end if
		  case 2
		    if m_dia2 then
			   m_i = m_i + 1
			end if
		  case 3
		    if m_dia3 then
			   m_i = m_i + 1
			end if
		  case 4
		    if m_dia4 then
			   m_i = m_i + 1
			end if
		  case 5
		    if m_dia5 then
			   m_i = m_i + 1
			end if
		  case 6
		    if m_dia6 then
			   m_i = m_i + 1
			end if
		  case 7
		    if m_dia7 then
			   m_i = m_i + 1
			end if
		end select
	  end if
	  if m_i < CInt(cant) then
	     m_factual = DateAdd("d", 1, m_factual)
	  end if
   loop while m_i < CInt(cant)
   
   hasta = m_factual

end sub 'busqFecha(desde,cant)


'---------------------------------------------------------------------------------------------------------
' SUB: datosPedVac(ternro,vacnro, byRef corresp, byRef pedidos) se encarga buscar los dias correspondientes y los pedidos de un periodo para un empleado
'---------------------------------------------------------------------------------------------------------
sub datosPedVac(ternro,vacnro, byRef corresp, byRef pedidos, vdiapednro)

	'Busco la cant. de dias corresp.
	l_sql = "SELECT vacdiascor.vacnro, vacdiascor.tipvacnro, "
	l_sql = l_sql & " vacdiascor.vdiascorcant "
	l_sql = l_sql & " FROM  vacdiascor "
	l_sql = l_sql & " WHERE vacdiascor.ternro =  " & ternro
	l_sql = l_sql & "   AND vacdiascor.vacnro =  " & vacnro
	
	rsOpenCursor l_rs, cn, l_sql, 0, adOpenDynamic 
	
	if not l_rs.eof then
	   corresp = CInt(l_rs("vdiascorcant"))
	else
	   corresp = 0
	end if
	
	l_rs.close
	
	'Busco la cant. de dias pedidos
	l_sql = "SELECT vacdiasped.vacnro, vacdiasped.ternro, SUM(vacdiasped.vdiaspedhabiles) as suma"
	l_sql = l_sql & " FROM  vacdiasped "
	l_sql = l_sql & " WHERE vacdiasped.ternro =  " & l_ternro
	l_sql = l_sql & "   AND vacdiasped.vdiaspedestado = -1 " 
	l_sql = l_sql & "   AND vacdiasped.vacnro = " & vacnro
	if Trim(vdiapednro) <> "" then
	   l_sql = l_sql & "   AND vacdiasped.vdiapednro <> " & vdiapednro
	end if
	l_sql = l_sql & " GROUP BY vacdiasped.vacnro, vacdiasped.ternro "
	
	rsOpenCursor l_rs, cn, l_sql, 0, adOpenDynamic 
	
	if not l_rs.eof then
	   if isNull(l_rs("suma")) then
	      pedidos = 0   
	   else
	      pedidos = CInt(l_rs("suma"))
	   end if
	else
	   pedidos = 0
	end if
	
	l_rs.close

end sub 'datosPedVac(ternro,vacnro, byRef corresp, byRef pedidos)

'---------------------------------------------------------------------------------------------------------
' SUB: cantidadDiasPedDisp(ternro,desde, byRef cantidad)
'      se encarga buscar la cantidad de dias pedidos disponibles
'---------------------------------------------------------------------------------------------------------
sub cantidadDiasPedDisp(ternro,desde, byRef cantidad,byRef correspondientes, vdiapednro)

	Dim corresp
	Dim pedidos
	Dim m_rs
	
	Set m_rs  = Server.CreateObject("ADODB.RecordSet")
	
	'Busco la cant. de dias corresp.
	l_sql = "SELECT * "
	l_sql = l_sql & " FROM  vacacion "
	l_sql = l_sql & " WHERE vacfecdesde <=  " & cambiafecha(desde,"YMD",true)
	l_sql = l_sql & "   AND vacfechasta >=  " & cambiafecha(desde,"YMD",true)
	
	rsOpenCursor m_rs, cn, l_sql, 0 , adOpenDynamic
	
	cantidad = 0
	correspondientes = 0
	
	do until m_rs.eof 
	
	   call datosPedVac(ternro,m_rs("vacnro"), corresp, pedidos, vdiapednro)
	
	   cantidad = cantidad + (corresp - pedidos)
	   correspondientes = correspondientes + corresp
	
	   m_rs.movenext 
	loop
	
	m_rs.close

end sub 'cantidadDiasPedDisp(ternro,desde, byRef cantidad)


'---------------------------------------------------------------------------------------------------------
' SUB:  pedMaximoHasta(ternro,desde, byRef hasta)
'       se encarga buscar el maximo dia hasta posible que se puede indicar en el pedido a partir del desde
'---------------------------------------------------------------------------------------------------------
sub pedMaximoHasta(ternro,fdesde, byRef hasta, vdiapednro)

	Dim corresp
	Dim pedidos
	Dim cantidad
	Dim total
	Dim totalFer
	Dim m_rs
	Dim desde
	
	desde = fdesde
	
	set m_rs  = Server.CreateObject("ADODB.RecordSet")
	
	'Busco la cant. de dias corresp.
	l_sql = "SELECT vacdiascor.vacnro, vacdiascor.tipvacnro, "
	l_sql = l_sql & " vacdiascor.vdiascorcant "
	l_sql = l_sql & " FROM  vacdiascor "
	l_sql = l_sql & " INNER JOIN vacacion ON vacacion.vacnro = vacdiascor.vacnro "
	l_sql = l_sql & " WHERE vacdiascor.ternro =  " & ternro
	l_sql = l_sql & "   AND vacfecdesde <=  " & cambiafecha(desde,"YMD",true)
	l_sql = l_sql & "   AND vacfechasta >=  " & cambiafecha(desde,"YMD",true)
	l_sql = l_sql & " ORDER BY vacfecdesde ASC "

	rsOpenCursor m_rs, cn, l_sql, 0 , adOpenDynamic
	
	hasta = DateAdd("d",-1,CDAte(desde))
	
	do until m_rs.eof 
	
	   desde = DateAdd("d",1,CDAte(hasta))
	
	   call iniTipoVac(m_rs("tipvacnro"))
	
	   call datosPedVac(ternro,m_rs("vacnro"), corresp, pedidos, vdiapednro)
	
	   cantidad = CInt(corresp) - CInt(pedidos)
	
	   if CInt(cantidad) > 0 then
	      call busqFecha(desde,cantidad,hasta,total,totalFer)
	   end if
	
	   m_rs.movenext 
	loop
	
	m_rs.close

end sub ' pedMaximoHasta(ternro,desde, byRef hasta)


'---------------------------------------------------------------------------------------------------------
' SUB:  busqFechaPedidos(ternro,desde,cant,byRef hasta, byRef total, byRef totalFer, vdiapednro, byRef errores)
'       considerando varios periodos
'---------------------------------------------------------------------------------------------------------
sub busqFechaPedidos(ternro, desde,cantDias,byRef hasta, byRef total, byRef totalFer, vdiapednro, byRef errores)

	Dim totalCantidad
	Dim corresp
	Dim pedidos
	Dim cantidad
	Dim total2
	Dim totalFer2
	Dim cant
	Dim totalCorr
	
	cant = cantDias
	errores = 0
	
	call cantidadDiasPedDisp(ternro,desde,totalCantidad, totalCorr, vdiapednro)
	
	if CInt(cant) > CInt(totalCantidad) then
	   errores = 1
	else
	    Dim m_rs 
		
		set m_rs  = Server.CreateObject("ADODB.RecordSet")
		
		'Busco la cant. de dias corresp.
		l_sql = "SELECT vacdiascor.vacnro, vacdiascor.tipvacnro, "
		l_sql = l_sql & " vacdiascor.vdiascorcant "
		l_sql = l_sql & " FROM  vacdiascor "
		l_sql = l_sql & " INNER JOIN vacacion ON vacacion.vacnro = vacdiascor.vacnro "
		l_sql = l_sql & " WHERE vacdiascor.ternro =  " & ternro
		l_sql = l_sql & "   AND vacfecdesde <=  " & cambiafecha(desde,"YMD",true)
		l_sql = l_sql & "   AND vacfechasta >=  " & cambiafecha(desde,"YMD",true)
		l_sql = l_sql & " ORDER BY vacfecdesde ASC "
		
		rsOpenCursor m_rs, cn, l_sql, 0 , adOpenDynamic
		
	    total = 0
	    totalFer = 0
		
		do until m_rs.eof 
		
		   call iniTipoVac(m_rs("tipvacnro"))
		
		   call datosPedVac(ternro,m_rs("vacnro"), corresp, pedidos, vdiapednro)
		   
		   cantidad = CInt(corresp) - CInt(pedidos)

		   if CInt(cant) <= CInt(cantidad) then
		      call busqFecha(desde,cant,hasta,total2,totalFer2)
			  total    = total + total2
			  totalFer = totalFer + totalFer2
			  exit do
		   else
		      if CInt(cantidad) > 0 then
			      call busqFecha(desde,cantidad,hasta,total2,totalFer2)
				  total    = total + total2
				  totalFer = totalFer + totalFer2
				  cant     = cant - cantidad
				  desde    = DateAdd("d",1,CDate(hasta))
			  end if
		   end if

		   m_rs.movenext 
		loop
		
		m_rs.close

	end if

end sub 'busqFechaPedidos(ternro,desde,cant,byRef hasta, byRef total, byRef totalFer, vdiapednro, byRef errores)


'---------------------------------------------------------------------------------------------------------
' SUB:  cantDiasPedidos(desde,hasta, byRef cantidad, byRef total, byRef totalFer, vdiapednro, byRef errores)
'       considerando varios periodos
'---------------------------------------------------------------------------------------------------------
sub cantDiasPedidos(ternro, desde,hasta, byRef cantidad, byRef total, byRef totalFer, vdiapednro, byRef errores)

	Dim totalCantidad
	Dim corresp
	Dim pedidos
	Dim cant2
	Dim total2
	Dim totalFer2
	Dim maxHasta
	
	errores = 0
	
	call pedMaximoHasta(ternro,desde,maxHasta,vdiapednro)
	
	if DateDiff("d",CDate(hasta),CDate(maxHasta)) < 0 then
	   errores = 1
	else
	    Dim m_rs 
		
		set m_rs  = Server.CreateObject("ADODB.RecordSet")
		
		'Busco la cant. de dias corresp.
		l_sql = "SELECT vacdiascor.vacnro, vacdiascor.tipvacnro, "
		l_sql = l_sql & " vacdiascor.vdiascorcant "
		l_sql = l_sql & " FROM  vacdiascor "
		l_sql = l_sql & " INNER JOIN vacacion ON vacacion.vacnro = vacdiascor.vacnro "
		l_sql = l_sql & " WHERE vacdiascor.ternro =  " & ternro
		l_sql = l_sql & "   AND vacfecdesde <=  " & cambiafecha(desde,"YMD",true)
		l_sql = l_sql & "   AND vacfechasta >=  " & cambiafecha(desde,"YMD",true)
		l_sql = l_sql & " ORDER BY vacfecdesde ASC "
		
		rsOpenCursor m_rs, cn, l_sql, 0 , adOpenDynamic
		
	    total = 0
	    totalFer = 0
		cantidad = 0
		
		do until m_rs.eof 
		
		   call iniTipoVac(m_rs("tipvacnro"))
		
		   call datosPedVac(ternro,m_rs("vacnro"), corresp, pedidos, vdiapednro)
		
		   cant2 = CInt(corresp) - CInt(pedidos)
		   
		   call busqFecha(desde,cant2,maxHasta,total2,totalFer2)
		   
 	       if DateDiff("d",CDate(hasta),CDate(maxHasta)) >= 0 then

			  call cantDias(desde,hasta,cant2,total2,totalFer2)
			  
			  total    = CInt(total)    + CInt(total2)
			  totalFer = CInt(totalFer) + CInt(totalFer2)
			  cantidad = CInt(cantidad) + CInt(cant2)
			  exit do
		   else
			  call cantDias(desde,maxHasta,cant2,total2,totalFer2)		   
			  total    = CInt(total)    + CInt(total2)
			  totalFer = CInt(totalFer) + CInt(totalFer2)
			  cantidad = CInt(cantidad) + CInt(cant2)
			  desde    = DateAdd("d",1,CDate(maxHasta))
		   end if

		   m_rs.movenext 
		loop
		
		m_rs.close

	end if

end sub 'cantDiasPedidos(ternro,desde,hasta, byRef cantidad, byRef total, byRef totalFer, vdiapednro, byRef errores)

%>

		
