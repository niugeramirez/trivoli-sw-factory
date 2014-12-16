<%

Dim m_tope_restrictivo
Dim m_tope_valor
Dim m_tope_acunro
Dim m_pliqmes
Dim m_pliqanio

'----------------------------------------------------------------------------------------------------
' Inicializa la busqueda de topes para los empleados
'----------------------------------------------------------------------------------------------------
sub inicializarObtTope(pliqnro)
    l_sql = " SELECT * FROM conf_vales "

    rsOpen l_rs, cn, l_sql, 0   

    if l_rs.eof then
      m_tope_restrictivo = false
    else
      m_tope_restrictivo = CInt(l_rs("restrictivo")) = -1
      m_tope_valor       = CInt(l_rs("topevale"))
      m_tope_acunro      = CInt(l_rs("acunro"))
    end if
   
    l_rs.close
   
    l_sql = " SELECT * FROM periodo WHERE pliqnro=" & pliqnro
	
	rsOpen l_rs, cn, l_sql, 0 
	
	m_pliqmes  = CInt(l_rs("pliqmes"))
	m_pliqanio = CInt(l_rs("pliqanio"))
	
	l_rs.close
	
	'Obtengo el mes anterior
	if m_pliqmes = 1 then
	   m_pliqmes = 12
	   m_pliqanio = m_pliqanio - 1
	else
	   m_pliqmes = m_pliqmes - 1
	end if

end sub 'inicializarObtTope(pliqnro)

'----------------------------------------------------------------------------------------------------
' Obtiene el tope del vale para un empleado
'----------------------------------------------------------------------------------------------------
function obtenerTope(ternro,pliqnro,valnro)
  Dim m_suma_vales
  Dim m_valor_acum
  Dim m_restante

  if m_tope_restrictivo then

	  'Obtengo todos los vales pedidos en el periodo y sumo sus valores
      l_sql = " SELECT valmonto FROM vales WHERE empleado=" & ternro & " AND pliqdto=" & pliqnro
	  if valnro <> "" then
	     l_sql = l_sql & " AND valnro <> " & valnro
	  end if
      rsOpen l_rs, cn, l_sql, 0

      m_suma_vales = 0	  
	  do until l_rs.eof
        m_suma_vales = m_suma_vales + CDbl(l_rs("valmonto"))
		l_rs.Movenext
      loop
	  
	  l_rs.close
	  
	  'Busco el valor del acumulador
	  l_sql = " SELECT ammonto,acunro "
      l_sql = l_sql & " FROM acu_mes"
      l_sql = l_sql & " WHERE acunro = " & m_tope_acunro
      l_sql = l_sql & " AND ammes =  " & m_pliqmes
      l_sql = l_sql & " AND amanio = " & m_pliqanio
      l_sql = l_sql & " AND ternro = " & ternro
	  
      rsOpen l_rs, cn, l_sql, 0   	  
	  
	  if not l_rs.eof then
	     m_valor_acum = CDbl(l_rs("ammonto"))
		 
		 'Obtengo un porcentaje del acum
		 m_valor_acum = (m_valor_acum * m_tope_valor) / 100
		 
		 m_restante = m_valor_acum - m_suma_vales
  	     l_rs.close		 
	  else
		 m_restante = 0
	     l_rs.close
	  end if
  else
      'Si no es restrictivo indico lo indico devolviendo -1
	  m_restante = -1
  end if
 
  obtenerTope = m_restante

end function 'obtenerTope(ternro)  

%>