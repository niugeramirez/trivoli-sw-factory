<%
'
'-----------------------------------------------------------------------------
'Archivo        : a_texto.inc
'Descripcion    : Funciones para exportaciones a archivo de texto
'Creador        : Alvaro Bayon
'Fecha Creacion : 27/01/2004
'-----------------------------------------------------------------------------


'Imprime el TEXTO de una determinada LONGITUD, alineado a izquierda. Los sobrantes se rellenan con blancos
sub imprimirTexto(texto,longitud)
	call imprimirCampo(texto,longitud," ",true)
end sub


'Imprime el TEXTO(número) de una determinada LONGITUD, alineado a derecha y relleno a izquierda con 0 
sub imprimirNumero(texto,longitud)
	call imprimirCampo(texto,longitud,"0",false)
end sub


'Salto de línea
sub saltarLinea()
	response.write vbcrlf
end sub


'Rutina genérica para imprimir un TEXTO, de una LONGITUD determinada. 
'Los sobrantes se rellenan con CARACTER
'Alineados a derecha (o no, dependiendo del valor)
sub imprimirCampo(texto,longitud,caracter,derecha)
'texto: obvio
'longitud: longitud a imprimir
'caracter: caracter de relleno para los que sobren
Dim txt
Dim u
Dim longTexto
if isnull(texto) then longTexto=0 else longTexto=len(texto) end if
u = CInt(longitud) - longTexto
if u<0 then
	response.write left(texto,longitud)
else
	if derecha then
		response.write texto & string(u,caracter)
	else
		response.write string(u,caracter) & texto
	end if
end if

end sub
%>