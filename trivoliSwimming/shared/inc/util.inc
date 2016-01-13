<%
'Modificaciones:
' 12/04/2004 - Scarpa D. - Cambio del control al obtener un valor

'-----------------------------------------------------------------------------------------------
'FUNCION: se fija si un campo es nulo, si es nulo devuelve el default, sino el campo
function getValor(campo,valDef)

   if len(campo) = 0 or isnull(campo) then
      getValor = valDef
   else      
      getValor = campo
   end if
   
end function 'getValor(campo,valDef)

'-----------------------------------------------------------------------------------------------
'FUNCION: funcion que obtiene el valor de un checkbox
function getCheckbox(formName)
  if len(request.form(formName)) <> 0 then
     getCheckbox = "-1"
  else
     getCheckbox = "0"  
  end if
end function 'getCheckbox(formName)

'-----------------------------------------------------------------------------------------------
'FUNCION: funcion que devuelve un string con comillas si no es nulo  y en caso contrario NULL
function getString(formName)
  if len(request.form(formName)) <> 0 then
     getString = "'" & request.form(formName) & "'"
  else
     getString = "NULL"  
  end if
end function 'getString(formName)

'-----------------------------------------------------------------------------------------------
'FUNCION: funcion que devuelve un string con comillas si no es nulo  y en caso contrario NULL
function getFecha(formName)
  if len(request.form(formName)) <> 0 then
     getFecha = cambiafecha(request.form(formName),"YMD",true) 
  else
     getFecha = "NULL"  
  end if
end function 'getFecha(formName)

%>
