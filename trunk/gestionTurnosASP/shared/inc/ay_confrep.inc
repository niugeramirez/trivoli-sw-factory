<%
on error goto 0

'-------------------------------------------------------------------------------------------------------------------
' Funciones Definidas en el Modulo:
'     * cargarConfRep(nroRep)     : carga los datos del reporte nro nroRep
'     * getColByVal(valor)        : Dado un valor(tipo) retorna la columna donde hay que almacenar el valor
'     * getEtiqueta(nroCol)       : Dado un nro de columna retorna la etiqueta de dicha columna
'     * getTipoByVal(valor)       : Dado un valor retorna el tipo de dicha columna
'     * getTipoByCol(nroCol)      : Dado un valor retorna el tipo de dicha columna
'     * getAccionByVal(valor)     : Dado un valor retorna la accion de dicha columna
'     * getValor2ByVal(valor)     : Dado un valor retorna el valor 2 de dicha columna
'     * getValorByCol(nroCol)     : Dado un nro de columna retorna el valor de dicha columna
'     * getCantCols()             : Retorna la cantidad de columnas que tiene el reporte
'     * getCantColsByTipo(tipo)   : Retorna la cantidad de columnas que tiene el reporte de un tipo dado
'     * getMaxCol()               : Retorna el mayor nro de columna
'     * iniValorCols()            : Inicializa el valor de las columnas con cero
'     * getValorCol(nro)          : Retorna el valor del nro de columna
'     * setValorCol(confval,val)  : Setea el valor de la columna nro con el valor val, de acuerdo con su configuracion
'-------------------------------------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------------
'funcion que se encarga de agregar a una lista de numero separados por coma, uno nuevo
function agregar(lista,dato)
Dim salida

   salida = lista

   if not existe(lista,dato) then
     if (lista = "") then
	   salida = dato
	 else
	   salida = lista & "," & dato
	 end if
   end if
   
   agregar = salida

end function 'agregar(lista,dato)

'-----------------------------------------------------------------------------------------------
'funcion que se encarga de agregar a una lista de numero separados por coma, uno nuevo
function existe(lista,dato)
Dim salida
Dim arr
Dim i
Dim tmp

   salida = false
   arr = split(lista,",")

   For i=0 To UBound(arr)	   

       tmp = split(arr(i),"@") 

       if (CStr(tmp(0)) = CStr(dato)) then
	      salida = true
   	   end if
   next

   existe = salida

end function 'existe(lista,dato)

'-----------------------------------------------------------------------------------------------
'funcion que se encarga de agregar a una lista de numero separados por coma, uno nuevo
function getAccionValorCol(valor,col)
Dim salida
Dim arr
Dim i
Dim tmp

   arr = split(li_valores_col(col),",")
   salida = "" 
   For i=0 To UBound(arr)	   
       tmp = split(arr(i),"@") 
	   
       if (CStr(tmp(0)) = CStr(valor)) then
	      salida = tmp(1)
   	   end if
   next

   getAccionValorCol = salida
end function 'getAccionValorCol(valor,col)


'-----------------------------------------------------------------------------------------------
Dim li_rscr
Dim li_sql
Dim li_cantcolumnas 

'li_nrocols, dado un valor almacena el nro de columna incremental
Dim li_nrocols(25000)

'li_columnas, dado un valor almacena el nro de columna definido en el confrep
Dim	li_columnas(25000)

'li_columnas, dado una columna(def.en el confrep) almacena su etiqueta
Dim li_etiquetas(25000)

'li_tipos, dado un valor almacena el tipo definido en el confrep
Dim li_tipos(25000)

'li_acciones, dado un valor almacena la accion definido en el confrep
Dim li_acciones(25000)

'li_valores2, dado un valor almacena el valor2 definido en el confrep
Dim li_valores2(25000)

'li_valores, dado una columna(def.en el confrep) almacena su valor
Dim li_valores(25000)

'li_valor_cols, dado una columna(def.en el confrep) almacena valores
Dim li_valor_cols(25000)

Dim li_valores_col(25000)
Dim li_titulo_col(25000)
Dim	li_tipos_col(25000)

Dim li_maxcolumna

'-------------------------------------------------------------------------------------------------------------------
' inicializar : inicializa los arreglos
'-------------------------------------------------------------------------------------------------------------------
sub inicializar
Dim li_i

  for li_i = 0 to 1536
    li_valores_col(li_i) = ""
    li_titulo_col(li_i)  = ""
	li_tipos_col(li_i)   = ""
  
	li_columnas(li_i)  = null
	li_etiquetas(li_i) = null
	li_tipos(li_i)     = null
	li_acciones(li_i)  = null
	li_nrocols(li_i)   = null
    li_valores(li_i)   = null
    li_valores2(li_i)  = null
  next

end sub 'inicializar

'-------------------------------------------------------------------------------------------------------------------
' cargarConfRep(nroRep) : carga los datos del reporte nro nroRep
'-------------------------------------------------------------------------------------------------------------------
sub cargarConfRep(nroRep)

	Set li_rscr = Server.CreateObject("ADODB.RecordSet")

	li_sql =          " SELECT confnrocol, conftipo, confaccion, confetiq, confval, confval2 " 
	li_sql = li_sql & " FROM confrep "
	li_sql = li_sql & " WHERE repnro = "& nroRep & " ORDER BY confnrocol"

	rsOpen li_rscr, cn, li_sql, 0 
	
	inicializar

	li_cantcolumnas = 0
	li_maxcolumna = 0
	do until li_rscr.eof
		li_cantcolumnas = li_cantcolumnas + 1

		if li_maxcolumna < li_rscr("confnrocol") then
		   li_maxcolumna = li_rscr("confnrocol")
		end if
		
		li_valores_col(li_rscr("confnrocol")) = agregar(li_valores_col(li_rscr("confnrocol")), (li_rscr("confval") & "@" & li_rscr("confaccion")))
        li_titulo_col(li_rscr("confnrocol")) = li_rscr("confetiq")
		li_tipos_col(li_rscr("confnrocol")) = agregar(li_tipos_col(li_rscr("confnrocol")), li_rscr("conftipo"))
		
        li_nrocols(li_rscr("confval"))		= li_cantcolumnas 
		li_columnas(li_rscr("confval"))     = li_rscr("confnrocol")
		li_etiquetas(li_rscr("confnrocol")) = li_rscr("confetiq")
		li_tipos(li_rscr("confval"))        = li_rscr("conftipo")
		li_acciones(li_rscr("confval"))     = li_rscr("confaccion")
		li_valores2(li_rscr("confval"))     = li_rscr("confval2")		
		li_valores(li_rscr("confnrocol"))   = li_rscr("confval")
		
		li_rscr.MoveNext
	loop
	
    li_rscr.close
	Set li_rscr = Nothing
end sub 'cargarConfRep 

'-------------------------------------------------------------------------------------------------------------------
' getColByVal(valor) : Dado un valor(tipo) retorna la columna donde hay que almacenar el valor
'-------------------------------------------------------------------------------------------------------------------
function getColByVal(valor)
  if isNull(li_columnas(valor)) then
     getColByVal = 0
  else
     getColByVal = li_columnas(valor)
  end if
end function 'getColByVal(valor)

'-------------------------------------------------------------------------------------------------------------------
' getEtiqueta(nroCol) : Dado un nro de columna retorna la etiqueta de dicha columna
'-------------------------------------------------------------------------------------------------------------------
function getEtiqueta(nroCol)
  if isNull(li_etiquetas(nroCol)) then
     getEtiqueta = ""
  else
     getEtiqueta = li_etiquetas(nroCol)
  end if
end function 'getEtiqueta(nroCol)

'-------------------------------------------------------------------------------------------------------------------
' getTipoByVal(valor) : Dado un valor retorna el tipo de dicha columna
'-------------------------------------------------------------------------------------------------------------------
function getTipoByVal(valor)
  if isNull(li_tipos(valor)) then
     getTipoByVal = ""
  else
     getTipoByVal = li_tipos(valor)	 
  end if
end function 'getTipoByVal(valor)

'-------------------------------------------------------------------------------------------------------------------
' getTipoByCol(nroCol) : Dado un valor retorna el tipo de dicha columna
'-------------------------------------------------------------------------------------------------------------------
function getTipoByCol(nroCol)
  if isNull(li_valores(nroCol)) then
     getTipoByCol = ""
  else
     getTipoByCol = li_tipos(li_valores(nroCol))	 
  end if
end function 'getTipoByCol(nroCol)

'-------------------------------------------------------------------------------------------------------------------
' getValorByCol(nroCol)     : Dado un nro de columna retorna el valor de dicha columna
'-------------------------------------------------------------------------------------------------------------------
function getValorByCol(nroCol)
  if isNull(li_valores(nroCol)) then
     getValorByCol = ""
  else
     getValorByCol = li_valores(nroCol) 
  end if
end function 'getValorByCol(nroCol)

'-------------------------------------------------------------------------------------------------------------------
' getAccionByVal(valor) : Dado un valor retorna la accion de dicha columna
'-------------------------------------------------------------------------------------------------------------------
function getAccionByVal(valor)
  if isNull(li_acciones(valor)) then
     getAccionByVal = ""
  else
     getAccionByVal = li_acciones(valor)	 
  end if

end function 'getAccionByVal(valor)

'-------------------------------------------------------------------------------------------------------------------
' getValor2ByVal(valor) : Dado un valor retorna valor2 de dicha columna
'-------------------------------------------------------------------------------------------------------------------
function getValor2ByVal(valor)
  if isNull(li_valores2(valor)) then
     getValor2ByVal = ""
  else
     getValor2ByVal = li_valores2(valor)	 
  end if

end function 'getValor2ByVal(valor)

'-------------------------------------------------------------------------------------------------------------------
' getCantCols() : Retorna la cantidad de columnas que tiene el reporte
'-------------------------------------------------------------------------------------------------------------------
function getCantCols()
  getCantCols = li_cantcolumnas
end function 'getCantCols()

'-------------------------------------------------------------------------------------------------------------------
' getCantColsByTipo(tipo) : Retorna la cantidad de columnas que tiene el reporte de un tipo dado
'-------------------------------------------------------------------------------------------------------------------
function getCantColsByTipo(tipo)
Dim li_i
Dim li_cant
  
  li_cant = 0 

'  for li_i = 0 to 1536
'     if not isNull(li_valores(li_i)) then
'       if CStr(tipo) = CStr(li_tipos(li_valores(li_i))) then
'	     li_cant = li_cant + 1
'	   end if
'	 end if	   
'  next 

  for li_i = 0 to 1536
     if existe(li_tipos_col(li_i),tipo) then
        li_cant = li_cant + 1
	 end if	   
  next 

  getCantColsByTipo = li_cant

end function 'getCantColsByTipo()

'-------------------------------------------------------------------------------------------------------------------
' getMaxCol() : Retorna el mayor nro de columna
'-------------------------------------------------------------------------------------------------------------------
function getMaxCol()
  getMaxCol = li_maxcolumna
end function 'getMaxCol()

'-------------------------------------------------------------------------------------------------------------------
' iniValorCols()         : Inicializa el valor de las columnas con cero
'-------------------------------------------------------------------------------------------------------------------
sub iniValorCols()
Dim li_i

  for li_i = 0 to 1536
	li_valor_cols(li_i) = 0
  next

end sub 'iniValorCols()

'-------------------------------------------------------------------------------------------------------------------
' getValorCol(nro)          : Retorna el valor del nro de columna
'-------------------------------------------------------------------------------------------------------------------
function getValorCol(nro)
  getValorCol = li_valor_cols(nro)
end function 'getValorCol(nro)

'-------------------------------------------------------------------------------------------------------------------
' setValorCol(confval,val) : Setea en todas las columnas que usan el tipo de hora confval, con el valor val
'                            de acuerdo al tipo de accion
'-------------------------------------------------------------------------------------------------------------------
sub setValorCol(confval,val)
Dim acc
Dim li_i
  for li_i = 0 to li_maxcolumna
    if existe(li_valores_col(li_i), confval) then 
       acc = getAccionValorCol(confval,li_i)	
       if CStr(acc) = "sumar" then
          li_valor_cols(li_i) = li_valor_cols(li_i) +  val  	 
  	   else
          if CStr(acc) = "restar" then
             li_valor_cols(li_i) = li_valor_cols(li_i) -  val  		
		  else
             li_valor_cols(li_i) = 543212345
		  end if
	   end if
	end if
  next   
end sub 'setValorCol(nro,conf,val)

%>
