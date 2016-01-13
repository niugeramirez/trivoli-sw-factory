<%
Sub DobleBrowse(l_sql, l_cadenafiltro)

dim l_rs
Dim l_texto
Dim l_indice
Dim l_indiceder
Dim l_indiceizq

response.write "<script languaje='javascript'>" & vbCrLf & vbCrLf
response.write "function CargarArreglo() {" & vbCrLf

Set l_rs = Server.CreateObject("ADODB.RecordSet")
rsOpen l_rs, cn, l_sql, 0 
l_indice = 0
l_indiceder = 0
l_indiceizq = 0
do until l_rs.eof
	l_texto =  l_rs(0) & " - " & l_rs(1)
    response.write "obj = " & l_obj & ";" & vbCrLf
    response.write "newOp = new Option();" & vbCrLf
	response.write "newOp.value = " & l_rs(0) & ";" & vbCrLf
    response.write "newOp.text = '" & l_texto & "';" & vbCrLf
	
    if inStr(1,l_cadenafiltro,"," & trim(l_rs(0)) & ",") <> 0 then
		response.write "selfil.options[" & l_indiceder & "] = newOp;"  & vbCrLf
		l_indiceder = l_indiceder + 1.		
    else
        response.write "nselfil.options[" & l_indiceizq & "] = newOp;" & vbCrLf
		l_indiceizq = l_indiceizq + 1.		
    end if
	l_indice = l_indice + 1
	l_rs.MoveNext
loop
response.write "}" & vbCrLf

response.write "function Todos(fuente,destino){"  & vbCrLf &_
    	   "x=fuente.length;"  & vbCrLf &_
		   "    for (i=1;i<=x;i++){" & vbCrLf &_
           "var opcion = new Option();" & vbCrLf &_
           "opcion.value= fuente[0].value;" & vbCrLf &_
           "opcion.text  = fuente[ 0].text;" & vbCrLf &_
           "fuente.remove(0);" & vbCrLf &_
           "destino.add(opcion);" & vbCrLf &_
    			"}" & vbCrLf &_
		   "}" 

response.write "function Uno(fuente,destino){" & vbCrLf &_
    	   "    if (fuente.selectedIndex != -1) { " & vbCrLf &_
    	   "    var opcion = new Option();" & vbCrLf &_
    	   "    opcion.value= fuente[fuente.selectedIndex].value;" & vbCrLf &_
    	   "    opcion.text  = fuente[fuente.selectedIndex].text;" & vbCrLf &_
    	   "    fuente.remove(fuente.selectedIndex);" & vbCrLf &_
    	   "    destino.add(opcion);" & vbCrLf &_
    	   "    destino[destino.length-1].focus();" & vbCrLf &_
    	   "}" & vbCrLf &_
    	   "}" & vbCrLf

response.write "function Aceptar(obj){" & vbCrLf &_
    	   "    var cadena = ',';" & vbCrLf &_
    	   "    var i;" & vbCrLf &_
    	   "    var long = selfil.length-1;" & vbCrLf &_
    	   "    for (i=0;i<=long;i++){" & vbCrLf &_
    	   "      cadena = cadena + selfil[i].value + ','  ;" & vbCrLf &_
    	   "    }" & vbCrLf &_
    	   "    obj.value = cadena;" & vbCrLf &_
    	   "}" & vbCrLf

response.write "</script>" & vbCrLf

l_rs.Close
'l_rs = Nothing
cn.Close
'cn = Nothing

End Sub
%>