<%
function cambiafecha (actual,formato,texto)
  Dim auxi
  Select Case l_base
        Case "1"
 			auxi  = "'" & mid(actual,7,4) & "/" & mid(actual,4,2) & "/" & mid(actual,1,2) & "'"
        Case "2" 
  	 		auxi  = "'" & mid(actual,4,2) & "/" & mid(actual,1,2) & "/" & mid(actual,7,4) & "'"
			'auxi  = "'" & mid(actual,7,2) & "/" & mid(actual,4,2) & "/" & mid(actual,1,2) & "'"
			'auxi  = "'" & mid(actual,1,2) & "/" & mid(actual,4,2) & "/" & mid(actual,7,4) &"'" 
        Case "3" 
	  	 	auxi  = "'" & mid(actual,1,2) & "/" & mid(actual,4,2) & "/" & mid(actual,7,4) &"'" 
        Case "8" 
  	 		auxi  = "'" & mid(actual,4,2) & "/" & mid(actual,1,2) & "/" & mid(actual,7,4) & "'"
        Case "9" 
  	 		auxi  = "'" & mid(actual,4,2) & "/" & mid(actual,1,2) & "/" & mid(actual,7,4) & "'"			
	Case "10" 
  	 		auxi  = "'" & mid(actual,4,2) & "/" & mid(actual,1,2) & "/" & mid(actual,7,4) & "'"
			'auxi  = "'" & mid(actual,7,2) & "/" & mid(actual,4,2) & "/" & mid(actual,1,2) & "'"
			'auxi  = "'" & mid(actual,1,2) & "/" & mid(actual,4,2) & "/" & mid(actual,7,4) &"'" 
        Case "12" 
  	 		auxi  = "'" & mid(actual,4,2) & "/" & mid(actual,1,2) & "/" & mid(actual,7,4) & "'"
			'auxi  = "'" & mid(actual,7,2) & "/" & mid(actual,4,2) & "/" & mid(actual,1,2) & "'"
			'auxi  = "'" & mid(actual,1,2) & "/" & mid(actual,4,2) & "/" & mid(actual,7,4) &"'" 
        Case "13" 
  	 		auxi  = "'" & mid(actual,4,2) & "/" & mid(actual,1,2) & "/" & mid(actual,7,4) & "'"
			'auxi  = "'" & mid(actual,7,2) & "/" & mid(actual,4,2) & "/" & mid(actual,1,2) & "'"
			'auxi  = "'" & mid(actual,1,2) & "/" & mid(actual,4,2) & "/" & mid(actual,7,4) &"'" 
        Case "14" 
  	 		auxi  = "'" & mid(actual,4,2) & "/" & mid(actual,1,2) & "/" & mid(actual,7,4) & "'"
			'auxi  = "'" & mid(actual,7,2) & "/" & mid(actual,4,2) & "/" & mid(actual,1,2) & "'"
			'auxi  = "'" & mid(actual,1,2) & "/" & mid(actual,4,2) & "/" & mid(actual,7,4) &"'" 
        Case "15" 
  	 		auxi  = "'" & mid(actual,4,2) & "/" & mid(actual,1,2) & "/" & mid(actual,7,4) & "'" 
  End Select
  cambiafecha = auxi
end function

function cambiafechahora (actual,formato,texto)
  Dim auxi
  Select Case l_base
        Case "1"
 			auxi  = "'" & mid(actual,7,4) & "/" & mid(actual,4,2) & "/" & mid(actual,1,2) & "'"
        Case "2" 
  	 		auxi  = "'" & mid(actual,4,2) & "/" & mid(actual,1,2) & "/" & mid(actual,7,4) & "'"
			'auxi  = "'" & mid(actual,7,2) & "/" & mid(actual,4,2) & "/" & mid(actual,1,2) & "'"
			'auxi  = "'" & mid(actual,1,2) & "/" & mid(actual,4,2) & "/" & mid(actual,7,4) &"'" 
        Case "3" 
	  	 	auxi  = "'" & mid(actual,1,2) & "/" & mid(actual,4,2) & "/" & mid(actual,7,4) &"'" 
        Case "8" 
  	 		auxi  = "'" & mid(actual,4,2) & "/" & mid(actual,1,2) & "/" & mid(actual,7,4) & "'"
        Case "9" 
  	 		auxi  = "'" & mid(actual,4,2) & "/" & mid(actual,1,2) & "/" & mid(actual,7,4) & "'"			
	Case "10" 
  	 		auxi  = "'" & mid(actual,4,2) & "/" & mid(actual,1,2) & "/" & mid(actual,7,4) & "'"
			'auxi  = "'" & mid(actual,7,2) & "/" & mid(actual,4,2) & "/" & mid(actual,1,2) & "'"
			'auxi  = "'" & mid(actual,1,2) & "/" & mid(actual,4,2) & "/" & mid(actual,7,4) &"'" 
        Case "12" 
  	 		auxi  = "'" & mid(actual,4,2) & "/" & mid(actual,1,2) & "/" & mid(actual,7,4) & "'"
			'auxi  = "'" & mid(actual,7,2) & "/" & mid(actual,4,2) & "/" & mid(actual,1,2) & "'"
			'auxi  = "'" & mid(actual,1,2) & "/" & mid(actual,4,2) & "/" & mid(actual,7,4) &"'" 
        Case "13" 
  	 		auxi  = "'" & mid(actual,4,2) & "/" & mid(actual,1,2) & "/" & mid(actual,7,4) & "'"
			'auxi  = "'" & mid(actual,7,2) & "/" & mid(actual,4,2) & "/" & mid(actual,1,2) & "'"
			'auxi  = "'" & mid(actual,1,2) & "/" & mid(actual,4,2) & "/" & mid(actual,7,4) &"'" 
        Case "14" 
  	 		auxi  = "'" & mid(actual,4,2) & "/" & mid(actual,1,2) & "/" & mid(actual,7,4) & "'"
			'auxi  = "'" & mid(actual,7,2) & "/" & mid(actual,4,2) & "/" & mid(actual,1,2) & "'"
			'auxi  = "'" & mid(actual,1,2) & "/" & mid(actual,4,2) & "/" & mid(actual,7,4) &"'" 
        Case "15" 
  	 		auxi  = "'" & mid(actual,4,2) & "/" & mid(actual,1,2) & "/" & mid(actual,7,4) & "'" 
  End Select
  cambiafechahora = auxi
end function

function cambiaformato (fecha,hora )
  Dim auxi
  Select Case l_base
        Case "1"
 			auxi  = "'" & mid(actual,7,4) & "/" & mid(actual,4,2) & "/" & mid(actual,1,2) & "'"
        Case "2" 
  	 		auxi  =  "'" & mid(fecha,4,2) & "/" & mid(fecha,1,2) & "/" & mid(fecha,7,4) & " " & ucase ( replace (hora,"." ,""  ))   & "'" 
			'auxi  =  "'" & mid(fecha,4,2) & "/" & mid(fecha,1,2) & "/" & mid(fecha,7,4) & " " & hora  & "'"
			'auxi  = "'" & mid(actual,7,2) & "/" & mid(actual,4,2) & "/" & mid(actual,1,2) & "'"
			'auxi  = "'" & mid(actual,1,2) & "/" & mid(actual,4,2) & "/" & mid(actual,7,4) &"'" 
        Case "3" 
	  	 	auxi  = "'" & mid(actual,1,2) & "/" & mid(actual,4,2) & "/" & mid(actual,7,4) &"'" 
        Case "8" 
  	 		auxi  = "'" & mid(actual,4,2) & "/" & mid(actual,1,2) & "/" & mid(actual,7,4) & "'"
        Case "9" 
  	 		auxi  = "'" & mid(actual,4,2) & "/" & mid(actual,1,2) & "/" & mid(actual,7,4) & "'"			
	Case "10" 
  	 		auxi  = "'" & mid(actual,4,2) & "/" & mid(actual,1,2) & "/" & mid(actual,7,4) & "'"
			'auxi  = "'" & mid(actual,7,2) & "/" & mid(actual,4,2) & "/" & mid(actual,1,2) & "'"
			'auxi  = "'" & mid(actual,1,2) & "/" & mid(actual,4,2) & "/" & mid(actual,7,4) &"'" 
        Case "12" 
  	 		auxi  = "'" & mid(actual,4,2) & "/" & mid(actual,1,2) & "/" & mid(actual,7,4) & "'"
			'auxi  = "'" & mid(actual,7,2) & "/" & mid(actual,4,2) & "/" & mid(actual,1,2) & "'"
			'auxi  = "'" & mid(actual,1,2) & "/" & mid(actual,4,2) & "/" & mid(actual,7,4) &"'" 
        Case "13" 
  	 		auxi  = "'" & mid(actual,4,2) & "/" & mid(actual,1,2) & "/" & mid(actual,7,4) & "'"
			'auxi  = "'" & mid(actual,7,2) & "/" & mid(actual,4,2) & "/" & mid(actual,1,2) & "'"
			'auxi  = "'" & mid(actual,1,2) & "/" & mid(actual,4,2) & "/" & mid(actual,7,4) &"'" 
        Case "14" 
  	 		auxi  = "'" & mid(actual,4,2) & "/" & mid(actual,1,2) & "/" & mid(actual,7,4) & "'"
			'auxi  = "'" & mid(actual,7,2) & "/" & mid(actual,4,2) & "/" & mid(actual,1,2) & "'"
			'auxi  = "'" & mid(actual,1,2) & "/" & mid(actual,4,2) & "/" & mid(actual,7,4) &"'" 
        Case "15" 
  	 		auxi  = "'" & mid(actual,4,2) & "/" & mid(actual,1,2) & "/" & mid(actual,7,4) & "'" 
  End Select
  cambiaformato = auxi
end function




function esmenor(fec1,fec2)
  Dim dif
  Dim f1
  Dim f2
  Select Case l_base
        Case "1" 
 			f1  = mid(fec1,4,2) & "/" & mid(fec1,1,2) & "/" & mid(fec1,7,4)
 			f2  = mid(fec2,4,2) & "/" & mid(fec2,1,2) & "/" & mid(fec2,7,4)
        Case "2" 
 			f1  = mid(fec1,4,2) & "/" & mid(fec1,1,2) & "/" & mid(fec1,7,4)
 			f2  = mid(fec2,4,2) & "/" & mid(fec2,1,2) & "/" & mid(fec2,7,4)
        Case "3" 
 			f1  = mid(fec1,4,2) & "/" & mid(fec1,1,2) & "/" & mid(fec1,7,4)
 			f2  = mid(fec2,4,2) & "/" & mid(fec2,1,2) & "/" & mid(fec2,7,4)
        Case "8" 
 			f1  = mid(fec1,4,2) & "/" & mid(fec1,1,2) & "/" & mid(fec1,7,4)
 			f2  = mid(fec2,4,2) & "/" & mid(fec2,1,2) & "/" & mid(fec2,7,4)
        Case "9" 
 			f1  = mid(fec1,4,2) & "/" & mid(fec1,1,2) & "/" & mid(fec1,7,4)
 			f2  = mid(fec2,4,2) & "/" & mid(fec2,1,2) & "/" & mid(fec2,7,4)
        Case "10" 
 			f1  = mid(fec1,4,2) & "/" & mid(fec1,1,2) & "/" & mid(fec1,7,4)
 			f2  = mid(fec2,4,2) & "/" & mid(fec2,1,2) & "/" & mid(fec2,7,4)			
        Case "12" 
 			f1  = mid(fec1,4,2) & "/" & mid(fec1,1,2) & "/" & mid(fec1,7,4)
 			f2  = mid(fec2,4,2) & "/" & mid(fec2,1,2) & "/" & mid(fec2,7,4)
        Case "13" 
 			f1  = mid(fec1,4,2) & "/" & mid(fec1,1,2) & "/" & mid(fec1,7,4)
 			f2  = mid(fec2,4,2) & "/" & mid(fec2,1,2) & "/" & mid(fec2,7,4)
        Case "14" 
 			f1  = mid(fec1,4,2) & "/" & mid(fec1,1,2) & "/" & mid(fec1,7,4)
 			f2  = mid(fec2,4,2) & "/" & mid(fec2,1,2) & "/" & mid(fec2,7,4)
        Case "15" 
 			f1  = mid(fec1,4,2) & "/" & mid(fec1,1,2) & "/" & mid(fec1,7,4)
 			f2  = mid(fec2,4,2) & "/" & mid(fec2,1,2) & "/" & mid(fec2,7,4)

  End Select
  'f1=cambiafecha(fec1,"MDY",false)
  'f2=cambiafecha(fec2,"MDY",false)
  dif = DateDiff("d", f2, f1)   
  if dif > 0 then
  	esmenor = false
  else
  	esmenor = true
  end if
  
end function

function calculardia(fecha)
	Dim nrodia
	nrodia = Weekday(fecha)
	if nrodia = 1 then
		calculardia = "DO"
	end if
	if nrodia = 2 then
		calculardia = "LU"
	end if
	if nrodia = 3 then
		calculardia = "MA"
	end if
	if nrodia = 4 then
		calculardia = "MI"
	end if
	if nrodia = 5 then
		calculardia = "JU"
	end if
	if nrodia = 6 then
		calculardia = "VI"
	end if
	if nrodia = 7 then
		calculardia = "SA"
	end if
	
end function

'Convierte la fecha a al formato estandar ISO
function fechaISO(str)
   Dim fec   
   
   if esFechaValida(str) then
      fec = CDate(str)
      fechaISO =  Year(fec) & "-" & Month(fec) & "-" & Day(fec)
   else
      fechaISO = str
   end if
end function 'fechaISO(str)

function esFechaValida(str)
  Dim salida 
  salida = false
  
  if not isNull(str) then
     if inStr(1,CStr(str),",") > 0 then
 	    salida = false
	 else
	    if inStr(1,CStr(str),".") > 0 then
		   salida = false
		else
 	       salida = isDate(str)
		end if
	 end if
  end if
  
  esFechaValida = salida
end function 'esFechaValida(str)

function anioBisiesto(anyo)
'si el año introducido es de dos cifras lo pasamos al periodo de 1900. Ejemplo: 25 > 1925
dim fin

   if anyo < 100 then
      fin = anyo + 2000
   else fin = anyo
   end if

' primera condicion: si el resto de dividir el año entre 4 no es cero > el año no es bisiesto
' es decir, obtenemos año modulo 4, teniendo que cumplirse anyo mod(4)=0 para bisiesto

   if (fin mod 4 <> 0) then
      anioBisiesto = false
   else
       if (fin mod 100 = 0) then
          ' si el año es divisible por 4 y por 100 y divisible por 400 > es bisiesto
            if (fin mod 400 = 0) then
                anioBisiesto = true
                ' si es divisible por 4 y por 100 pero no lo es por 400 > no es bisiesto
            else anioBisiesto false
            end if
           ' si es divisible por 4 y no es divisible por 100 > el año es bisiesto
       else anioBisiesto = true
       end if
   end if
end function

'-------------------------------------------------------------------------------------------
' Descripcion: campoFechaBD(campoBD,parteFecha)
' campoDB es el nombre del campo de la BD
' parteFecha indica que parte de la fecha queremos obtener y los valores posible son.
'     A - Año
'     M - Mes
'     D - Dia
'-------------------------------------------------------------------------------------------
function campoFechaBD(campoBD,parteFecha)
  Dim salida
  Dim TipoBD

  Select Case l_base
        Case "1"
 			TipoBD = "SQLServer"
' 			TipoBD = "Oracle"
        Case "2" 
 			TipoBD = "SQLServer"
        Case "3" 
 			TipoBD = "SQLServer"
        Case "8" 
 			TipoBD = "SQLServer"
        Case "9" 
 			TipoBD = "SQLServer"
        Case "10" 
 			TipoBD = "SQLServer"
        Case "12" 
 			TipoBD = "SQLServer"
        Case "13" 
 			TipoBD = "SQLServer"
        Case "14" 
 			TipoBD = "SQLServer"
        Case "15" 
 			TipoBD = "SQLServer"

  End Select

  Select Case TipoBD

     case "SQLServer"

	    Select Case parteFecha
		   case "A"
		      salida = "datepart(yy," & campoBD & ")"
		   case "M"
		      salida = "datepart(m,"  & campoBD & ")"
		   case "D"
		      salida = "datepart(d,"  & campoBD & ")"
		end select

	 case "Oracle"

        Select Case parteFecha
		   case "A"
		      salida = "TO_CHAR(YYYY," & campoBD & ")"
		   case "M"
		      salida = "TO_CHAR(MM,"  & campoBD & ")"
		   case "D"
		      salida = "TO_CHAR(DD,"  & campoBD & ")"
	    end select

  end select

  campoFechaBD = salida

end function

%>
