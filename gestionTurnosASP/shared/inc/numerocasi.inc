<%
on error goto 0
' 29/12/2003 - Scarpa D. - Se agrego la funcion NumerosALetras(NumeroA) 
' 12/01/2004 - Scarpa D. - Se agrego un control para determinar la conf. regional del int. explorer
' 11/03/2004 - Scarpa D. - Correcion en la funcion NumerosALetras

Dim SeparadorDecimal
Dim letras

'Me fijo cual es el separador de decimales del cliente
select case mid(CStr(Request.ServerVariables("HTTP_ACCEPT_LANGUAGE")),1,2)
  case "es"
     SeparadorDecimal = "."
  case else
     SeparadorDecimal = ","
end select

function cambiadecimal (actual)
  Dim aux
  actual = Round(actual, 2)
  aux = InStr(1,actual,".")
  if aux = 0 then
  	cambiadecimal = actual & ".00"
  else
  	if len(actual)-aux = 1 then
	  	cambiadecimal = actual & "0"
	else
    	cambiadecimal = actual
	end if
  end if
end function

Dim numeros(103) 

Function Centenas(VCentena ) 
    If VCentena = 1 Then
       Centenas = numeros(100)
    Else
       If VCentena = 5 Then
          Centenas = numeros(101)
          Else
             If VCentena = 7 Then
'                Centenas = letras & numeros(102)
                 Centenas = numeros(102)
             Else
                If VCentena = 9 Then
'                  Centenas = letras & numeros(103)
                   Centenas = numeros(103)
                Else
                   Centenas = numeros(VCentena)
                End If
             End If
       End If
    End If
End Function

Function Unidades(VUnidad ) 
   Unidades = numeros(VUnidad)
End Function

Function Decenas(VDecena )
   Decenas = numeros(VDecena)
End Function

Sub Inicializar()
   numeros(0) = "cero"
   numeros(1) = "uno"
   numeros(2) = "dos"
   numeros(3) = "tres"
   numeros(4) = "cuatro"
   numeros(5) = "cinco"
   numeros(6) = "seis"
   numeros(7) = "siete"
   numeros(8) = "ocho"
   numeros(9) = "nueve"
   numeros(10) = "diez"
   numeros(11) = "once"
   numeros(12) = "doce"
   numeros(13) = "trece"
   numeros(14) = "catorce"
   numeros(15) = "quince"
   numeros(16) = "dieciseis"
   numeros(17) = "diecisiete"
   numeros(18) = "dieciocho"
   numeros(19) = "diecinueve"
   numeros(20) = "veinte"
   numeros(30) = "treinta"
   numeros(40) = "cuarenta"
   numeros(50) = "cincuenta"
   numeros(60) = "sesenta"
   numeros(70) = "setenta"
   numeros(80) = "ochenta"
   numeros(90) = "noventa"
   numeros(100) = "ciento"
   numeros(101) = "quinientos"
   numeros(102) = "setecientos"
   numeros(103) = "novecientos"
End Sub

Function NumerosALetras(NumeroA ) 

    if CDbl(NumeroA) = 0 then
	   NumerosALetras = "cero"
	else
	    if CDbl(NumeroA) < 0 then
		   NumerosALetras = "menos " & NumerosALetras(CDbl(NumeroA) * (-1))
		else
	
		    Dim HuboCentavos 
		    Dim Decimales 
		    Decimales = NumeroA - Int(NumeroA)
		    NumeroA = Int(NumeroA)
		    Inicializar
		    letras = ""
		    Do
		
		       '*---> Validación si se pasa de 100 millones
		       If NumeroA >= 1000000000 Then
		          letras = "Error en Conversión a Letras"
		          NumeroA = 0
		          Decimales = 0
		       End If
		
		       '*---> Centenas de Millón
		       If (NumeroA < 1000000000) And (NumeroA >= 100000000) Then
		          If (Int(NumeroA / 100000000) = 1) And ((NumeroA - (Int(NumeroA / 100000000) * 100000000)) < 1000000) Then
		             letras = letras & "cien millones "
		          Else
		             letras = letras & Centenas(Int(NumeroA / 100000000))
		             If (Int(NumeroA / 100000000) <> 1) And (Int(NumeroA / 100000000) <> 5) And (Int(NumeroA / 100000000) <> 7) And (Int(NumeroA / 100000000) <> 9) Then
		                letras = letras & "cientos "
		             Else
		                letras = letras & " "
		             End If
		          End If
		          NumeroA = NumeroA - (Int(NumeroA / 100000000) * 100000000)
		       End If
		
		       '*---> Decenas de Millón
		       If (NumeroA < 100000000) And (NumeroA >= 10000000) Then
		          If Int(NumeroA / 1000000) < 20 Then
		             letras = letras & Decenas(Int(NumeroA / 1000000))
		             letras = letras & " millones "
		             NumeroA = NumeroA - (Int(NumeroA / 1000000) * 1000000)
		          Else
		             letras = letras & Decenas(Int(NumeroA / 10000000) * 10)
		             NumeroA = NumeroA - (Int(NumeroA / 10000000) * 10000000)
		             If NumeroA > 1000000 Then
		                letras = letras & " y "
		             End If
		          End If
		       End If
		
		       '*---> Unidades de Millón
		       If (NumeroA < 10000000) And (NumeroA >= 1000000) Then
		          If Int(NumeroA / 1000000) = 1 Then
		             letras = letras & " un millón "
		          Else
		             letras = letras & Unidades(Int(NumeroA / 1000000))
		             letras = letras & " millones "
		          End If
		          NumeroA = NumeroA - (Int(NumeroA / 1000000) * 1000000)
		       End If
		
		       '*---> Centenas de Millar
		       If (NumeroA < 1000000) And (NumeroA >= 100000) Then
		          If (Int(NumeroA / 100000) = 1) And ((NumeroA - (Int(NumeroA / 100000) * 100000)) < 1000) Then
		             letras = letras & "cien mil "
		          Else
		             letras = letras & Centenas(Int(NumeroA / 100000))
					 
		             If (Int(NumeroA / 100000) <> 1) And (Int(NumeroA / 100000) <> 5) And (Int(NumeroA / 100000) <> 7) And (Int(NumeroA / 100000) <> 9) Then
		                letras = letras & "cientos "
		             Else
		                letras = letras & " "
		             End If
		          End If
				  If ((NumeroA - (Int(NumeroA / 100000) * 100000)) = 0) then letras = letras & " mil"
		          NumeroA = NumeroA - (Int(NumeroA / 100000) * 100000)
		       End If
		
		       '*---> Decenas de Millar
			   If (NumeroA < 100000) And (NumeroA >= 10000) Then
		          If Int(NumeroA / 1000) < 20 Then
				  	 letras = letras & Decenas(Int(NumeroA / 1000))
		             letras = letras & " mil "
		             NumeroA = NumeroA - (Int(NumeroA / 1000) * 1000)
		          Else
		             letras = letras & Decenas(Int(NumeroA / 10000) * 10)
		             NumeroA = NumeroA - (Int((NumeroA / 10000)) * 10000)
		             If NumeroA > 1000 Then
		                letras = letras & " y "
		             Else
		                letras = letras & " mil "
		             End If
		          End If
		       End If
		
		       '*---> Unidades de Millar
		       If (NumeroA < 10000) And (NumeroA >= 1000) Then
		          If Int(NumeroA / 1000) = 1 Then
		             letras = letras & "un"
		             'letras = letras & ""			 
		          Else
		             letras = letras & Unidades(Int(NumeroA / 1000))
		          End If
		          letras = letras & " mil "
		          NumeroA = NumeroA - (Int(NumeroA / 1000) * 1000)
		       End If
		
		       '*---> Centenas
		       If (NumeroA < 1000) And (NumeroA > 99) Then
		          If (Int(NumeroA / 100) = 1) And ((NumeroA - (Int(NumeroA / 100) * 100)) < 1) Then
		             letras = letras & "cien "
		          Else
		             letras = letras & Centenas(Int(NumeroA / 100))
		             If (Int(NumeroA / 100) <> 1) And (Int(NumeroA / 100) <> 5) And (Int(NumeroA / 100) <> 7) And (Int(NumeroA / 100) <> 9) Then
		                letras = letras & "cientos "
		             Else
		                letras = letras & " "
		             End If
		          End If
		          NumeroA = NumeroA - (Int(NumeroA / 100) * 100)
		          
		       End If
		
		        '*---> Decenas
			   
			   If (NumeroA < 100) And (NumeroA > 9) Then
		          If NumeroA < 20 Then
				     letras = letras & Decenas(Int(NumeroA))
		             NumeroA = NumeroA - Int(NumeroA)
		          Else
				     letras = letras & Decenas(Int((NumeroA / 10)) * 10)
		             NumeroA = NumeroA - (Int((NumeroA / 10)) * 10)
		             If NumeroA > 0.99 Then
		                letras = letras & " y "
		             End If
		          End If
		    End If
		
		       '*---> Unidades
			   dim letra_unidad
			   
		       If (NumeroA < 10) And (NumeroA > 0.99) Then
			   	  if Unidades(Int(NumeroA)) = "uno" Then
				     letras = letras & "un"
				  else
				     letras = letras & Unidades(Int(NumeroA))
				  End If
		          NumeroA = NumeroA - Int(NumeroA)
		       End If
		
		    Loop Until (NumeroA = 0)
		    
		    '*---> Decimales
		    If (Decimales > 0) Then
		        letras = letras & " con "
		        letras = letras & FormatNumber(Decimales * 100, "00") & "/100"
		    Else
		        If (letras <> "Error en Conversión a Letras") And (Len(Trim(letras)) > 0) Then
		           letras = letras & " pesos"
		        End If
		    End If
		    
		    NumerosALetras = letras
			
		end if
	end if
	
End Function



%>
