function DiadeSemana (fecha)
	Select Case weekday(fecha)
           Case "1"  DiadeSemana = "Domingo"
           Case "2"  DiadeSemana = "Lunes"
           Case "3"  DiadeSemana = "Martes"
           Case "4"  DiadeSemana = "Miercoles"
           Case "5"  DiadeSemana = "Jueves"
           Case "6"  DiadeSemana = "Viernes"
		   Case "7"  DiadeSemana = "Sabado"
	End Select
end function

