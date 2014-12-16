<%
function cambiafecha (actual,formato,texto)
  Dim auxi
  Select Case l_base
        Case "1" 
 			auxi  = "'" & mid(actual,7,4) & "/" & mid(actual,4,2) & "/" & mid(actual,1,2) & "'"
        Case "2" 
  	 		auxi  = "'" & mid(actual,4,2) & "/" & mid(actual,1,2) & "/" & mid(actual,7,4) & "'"
        Case "3" 
	  	 	auxi  = "'" & mid(actual,1,2) & "/" & mid(actual,4,2) & "/" & mid(actual,7,4) &"'" 
		Case "8" 
  	 		auxi  = "'" & mid(actual,4,2) & "/" & mid(actual,1,2) & "/" & mid(actual,7,4) & "'"
        Case "12" 
  	 		auxi  = "'" & mid(actual,4,2) & "/" & mid(actual,1,2) & "/" & mid(actual,7,4) & "'"

        Case "13" 
  	 		auxi  = "'" & mid(actual,4,2) & "/" & mid(actual,1,2) & "/" & mid(actual,7,4) & "'"

  End Select
  cambiafecha = auxi
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
        Case "12" 
 			f1  = mid(fec1,4,2) & "/" & mid(fec1,1,2) & "/" & mid(fec1,7,4)
 			f2  = mid(fec2,4,2) & "/" & mid(fec2,1,2) & "/" & mid(fec2,7,4)
        Case "13" 
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


%>
