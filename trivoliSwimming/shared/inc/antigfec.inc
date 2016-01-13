<%

Sub DIF_FECHAS(f1, f2, dias, meses, anios)

dim numdiasmes, d1, d2

d1 = dateserial( year(f1), month(f1), 01)
meses = (month(f1) mod 12) + 1
anios = year(f1) + fix (month(f1) / 12)
d2 = dateserial(anios, meses, 01)

numdiasmes = d2 - d1

meses = 0
anios = 0

dias = day(f2) - day(f1)
meses = month (f2) - month(f1)
anios = year(f2) - year(f1)
if dias < 0 then
   meses = meses - 1
   dias = dias + numdiasmes
end if
if meses < 0 then
   anios = anios - 1
   meses = meses + 12
end if 
end sub

SUB DIASTRAB(ternro, DESDE, HASTA, DIASH )
Dim d1, d2, aux, aux2, dxsem

dxsem = 5

d1 = weekday(desde)
d2 = weekday(hasta)

aux = (hasta - desde) + 1
if aux < 7 then
  if aux > dxsem then
    diash = dxsem
  else
    diash = aux
  end if
else
  if aux = 7 then
    diash = dxsem
  else
    aux2 = 8 - d1 + d2
    if aux2 < 7 then
      if aux2 > dxsem then
        aux2 = dxsem
      end if
    else
	  if aux = 7 then
	    aux2 = dxsem
	  end if
      if aux2 >= 7 then
	    aux2 = abs(aux2 - 7) + fix(aux2 / 7) * dxsem
        diash = (aux2 + fix ((aux - aux2) / 7) * dxsem )
      end if
	end if
  end if
end if
aux = 0

'find first per.tercero of per.empleado no-lock no-error.
'find first pais of per.tercero no-lock no-error.
'if available(pais)
'  then do:
'    for each feriado where (feriado.tipferinro = 2) and
'                           (feriado.fericodext = pais.paisnro) and
'                           (feriado.ferifecha > desde) and
'                           (feriado.ferifecha < hasta) no-lock:
'    
'    if weekday(feriado.ferifecha) > 1 then diash = diash - 1.
'    end.
'  end.

'find first per.convenio of per.empleado no-lock no-error.
'if available(per.convenio)
'  then do:
'    for each feriado where (feriado.tipferinro = 1) and
'                           (feriado.fericodext = per.convenio.convnro) and
'                           (feriado.ferifecha > desde) and
'                           (feriado.ferifecha < hasta) no-lock:
'      if weekday(feriado.ferifecha) > 1 
'      then 
'        diash = diash - 1.
'    end.
'  end.

'find first per.sucursal of per.empleado no-lock no-error.
'if available(per.sucursal)
'  then do:
'    for each feriado where (feriado.tipferinro = 3) and
'                           (feriado.fericodext = per.sucursal.ternro) and
'                           (feriado.ferifecha > desde) and
'                           (feriado.ferifecha < hasta) no-lock:    
'      if weekday(feriado.ferifecha) > 1 
'      then 
'        diash = diash - 1.
'    end.
'  end.

end sub

Sub antigfecvieja(ternro, fechafin, antdia, antmes, antanio, diashab)

Dim aux1, aux2, aux3, fecalta, fecbaja, seguir, q, sql,rsauxaux 

Set rsauxaux = Server.CreateObject("ADODB.RecordSet")
sql = "SELECT * FROM fases where fases.empleado = " & ternro & " AND fases.real = -1 " 
rsOpen rsauxaux, cn, sql, 0
do until rsauxaux.eof
    fecalta = rsauxaux("altfec")

'   Verificar si se trata de un registro completo (alta/baja) o solo de un alta */
    IF rsauxaux("estado") THEN
	  fecbaja = ""           ' solo es un alta, tomar el fechafin
    ELSE
	  fecbaja = rsauxaux("bajfec") ' se trata de un registro completo
	end if

    IF fecbaja <> "" THEN 
        Call DIF_FECHAS(fecalta, fecbaja, aux1, aux2, aux3)
        antdia  = antdia + aux1
        antmes  = antmes + aux2 + fix (antdia / 30)
        antanio = antanio + aux3 + fix (antmes / 12)
        antdia  = antdia mod 30
        antmes  = antmes mod 12.

        IF antanio = 0 THEN
           call diastrab (ternro ,fecalta, fecbaja, aux1)
           diashab = diashab + aux1
        END IF
    ELSE   ' TRABAJAR HATA LA FECHA INDICADA COMO PARAMETRO
        call DIF_FECHAS(fecalta, fechafin, aux1, aux2,  aux3)
        antdia  = antdia + aux1
        antmes  = antmes + aux2 + fix (antdia / 30)
        antanio = antanio + aux3 + fix (antmes / 12) 'SINO EN EL MES 12 DESDE EL INICIO TOMA UN ANIO
        antdia  = antdia mod 30
        antmes  = antmes mod 12
        IF antanio = 0 THEN
           call diastrab (ternro ,fecalta, fechafin, aux1)
           diashab = diashab + aux1
        END IF
    END IF
    rsauxaux.movenext
LOOP
rsauxaux.CLOSE

end sub



Sub antigfec (ternro, fechafin, antdia, antmes, antanio, diashab)
'----------------------------------------------------------------------------------
' 28/08/2003 - CCRossi - MODIFICADO.. lo vieja esta arriba con nombre antigfecvieja
'----------------------------------------------------------------------------------
Dim fecalta, fecbaja, seguir, q, sql, rs
Dim ultima, auxiliar

Set rs = Server.CreateObject("ADODB.RecordSet")
sql = "SELECT * FROM fases where fases.empleado = " & ternro & " AND fases.real = -1 " 
rsOpen rs, cn, sql, 0
do until rs.eof
    fecalta = rs("altfec")
    fecbaja = rs("bajfec") ' se trata de un registro completo

    IF fecbaja <> "" THEN 
	   'FASES CERRADAS --------------------------------------------
        antdia  = antdia + DateDiff("d", fecalta, fecbaja) 
        'response.write("antdia cerrada")
		'response.write(antdia)
    ELSE   
       'FASES ABIERTAS --------------------------------------------
        antdia  = antdia  + DateDiff("d", fecalta, fechafin)
        'response.write("antdia abierta")
		'response.write(antdia)
    END IF
    ultima = fecbaja
    rs.movenext
LOOP
rs.CLOSE
set rs = nothing

auxiliar=antdia
if antdia > 0 then
	' calcular año
	 if antdia > 365 then
		antanio = int(antdia / 365)
		antdia  = antdia - 365 * antanio + 1
	 end if	
	 if antdia > 30 then
		antmes = int(antdia/30.5)
		if month(ultima)=4 or month(ultima)=6 or month(ultima)=9 or month(ultima)=11 then
			antdia = antdia - int(antmes * 30.5) + 1
		else
			if month(ultima)=2 then
				antdia = antdia - int(antmes * 28) + 1
			else
				antdia = antdia - int(antmes * 30.5) + 1
			end if
		end if			
	 else
		antdia = antdia  + 1
	 end if	
end if 
end sub








%>



