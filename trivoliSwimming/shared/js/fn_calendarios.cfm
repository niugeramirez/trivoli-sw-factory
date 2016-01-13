function DiadeSemana (fecha)

	Select Case weekday(fecha.value)
           Case "1"  document.datos.caldia.value = "Domingo"
           Case "2"  document.datos.caldia.value = "Lunes"
           Case "3"  document.datos.caldia.value = "Martes"
           Case "4"  document.datos.caldia.value = "Miercoles"
           Case "5"  document.datos.caldia.value = "Jueves"
           Case "6"  document.datos.caldia.value = "Viernes"
		   Case "7"  document.datos.caldia.value = "Sabado"
		   Case Else document.datos.caldia.value = ""
	End Select
	
end function