<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/inc/sec.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/const.inc"-->
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<!--
Archivo: pasos_00.asp
Descripción: wizzard
Autor: lisandro moro & Lic. Muzzolon Leandro
Fecha: 23/10/2003
Modificado: 
	FFavre - 22-01-04 - Se agrego la opcion de poder decidir quedarse en un paso del wizard. Es a traves de la funcion SigPaso
	Alvaro Bayon 22-01-04 - Modifiqué para que no tenga en cuenta al usuario al cargar la lista de pasos.
-->
<html>
<head>
<link href="/trivoliSwimming/shared/css/tables4.css" rel="StyleSheet" type="text/css">
<title><%= Session("Titulo")%>Menu - Asistente de Conceptos - RHPro &reg;</title>
<script src="/trivoliSwimming/shared/js/fn_windows.js"></script>
<script src="/trivoliSwimming/shared/js/fn_confirm.js"></script>
<script src="/trivoliSwimming/shared/js/fn_ayuda.js"></script>
<script>
var jsSelRow = null;
var pintado = "Navy";
var despintado = "#4682B4";
function TRPintar(trs, nrop){
	trs.childNodes[0].style.backgroundColor = despintado;
	trs.childNodes[1].style.backgroundColor = despintado;
	trs.style.cursor = "hand";
}
function TRDesPintar(atrs, nrod){
	if (nrod != parent.datos.menunro.value){
			atrs.childNodes[0].style.backgroundColor = pintado;
			atrs.childNodes[1].style.backgroundColor = pintado;
	}else{
			atrs.childNodes[0].style.backgroundColor = despintado;
			atrs.childNodes[1].style.backgroundColor = despintado;
	}
}
function ear(otro, estado){
    var count=0;
	if (estado == 0){					//se produce al clickear sobre el menu
	    for (i=0; i < document.all.tabla.rows.length; i++) {
	        for (j=0; j < document.all.tabla.rows(i).cells.length; j++) {
				if (document.all.tabla.rows(i).id == parent.datos.menunro.value){
					document.all.tabla.rows(i).cells(j).style.backgroundColor = pintado;
				}
            count++;
    		}
	    }
	}else{								//se carga la pagina
    	for (i=1; i < document.all.tabla.rows.length -1; i++) {
        	for (j=0; j < document.all.tabla.rows(i).cells.length; j++) {
				document.all.tabla.rows(i).cells(j).style.backgroundColor = pintado;
				if (document.all.tabla.rows(i).id == parent.datos.menunro.value){
					document.all.tabla.rows(i).cells(j).style.backgroundColor = despintado;
				}else{
					document.all.tabla.rows(i).cells(j).style.backgroundColor = pintado;
						if (parent.datos.menunro.value == ""){
							document.all.tabla.rows(1).cells(j).style.backgroundColor = despintado; // pinto el primero si no hay ninguno seleccionado, ya que por efecto se empieza por el primero no?
							parent.datos.menunro.value = document.all.tabla.rows(1).id;
						}
					//document.all.tabla.rows(1).cells(j).style.backgroundColor = despintado; // pinto el primero si no hay ninguno seleccionado, ya que por efecto se empieza por el primero no?
					//parent.datos.menunro.value = document.all.tabla.rows(1).id;
				}
            count++;
	        }
    	}
	}
}
// Esta funcion permite poder quedarse en un paso determinado del wizard
// Se debe definir la funcion SigPaso en el IFRM
function SigPaso(pasasp, codigo, pasnro){
	
	if (parent.ifrm.SigPaso){
		switch (parent.ifrm.SigPaso()){
			// Se pasa al siguiente paso sin realizar nada
			case "1":
				parent.Abrir(pasasp, codigo, pasnro);
				break;
			// No se pasa al siguiente paso.
			case "2":
				ear(parent.datos.menunroant.value, 0);
				TRPintar(document.all.tabla.rows(parent.datos.menunroant.value), parent.datos.menunroant.value);
				parent.datos.menunro.value = parent.datos.menunroant.value;		
				parent.datos.menunroant.value = pasnro; 
				break;
		}
	}else{
		// No esta definida la funcion SigPaso en el ifrm. Se pasa al siguiente paso normalmente.
		parent.Abrir(pasasp, codigo, pasnro);
	}
}
</script>
<%

'on error goto 0

Dim l_wiznro
Dim l_codigo
Dim l_label
Dim l_nombre
Dim l_rs
Dim l_rs2
Dim l_sql
Dim l_pasos(1000)
Dim l_i

l_wiznro = request("wiznro")
l_codigo = request("codigo")
l_label  = request("label")
l_nombre = request("nombre")

Set l_rs  = Server.CreateObject("ADODB.RecordSet")
Set l_rs2 = Server.CreateObject("ADODB.RecordSet")

'Inicializo el estado de cada paso
For l_i=1 To 1000
  l_pasos(l_i) = 0
next

'Almaceno el estado de cada paso
l_sql = " SELECT extestado,pasos.pasnro "
l_sql = l_sql & " FROM paso_ext "
l_sql = l_sql & " INNER JOIN pasos ON pasos.pasnro = paso_ext.pasnro AND pasos.wiznro=" & l_wiznro
l_sql = l_sql & " WHERE paso_ext.extnro = " & l_codigo
l_sql = l_sql & " ORDER BY paso_ext.pasnro "

rsOpen l_rs, cn, l_sql, 0 

do until l_rs.eof 
  l_pasos(CInt(l_rs("pasnro"))) = l_rs("extestado")
  l_rs.moveNext
loop

l_rs.close

'------------------------------------------------------------------------------------------------------------
'FUNCION: controla si las dependencia de un paso ya estan resueltas
function dependenciasListas(listaDep)
  Dim f_arr 
  Dim f_i
  Dim l_salida
  
  f_arr = split(listaDep,",")
  
  l_salida = true
  
  For f_i=0 To UBound(f_arr) 
     l_salida = l_salida AND (CInt(l_pasos(f_arr(f_i))) = -1)
  next
  
  dependenciasListas = l_salida

end function 'dependenciasListas(listaDep)

%>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onload="ear(parent.datos.menunro.value, 1)">
<table border="5" cellpadding="0" cellspacing="0" height="100%" width="100%" id="tabla" >
  <tr >
	<th colspan="3"  height="25" style="background-color: Navy; border-bottom: 1px solid White;border-right: 0px solid White;"><b> <%'= l_label & ": " & l_nombre%><%= l_nombre%> </b></th>
  </tr>	
<%
'------------------------------------------------------------------------------------------------------------
'EMPIEZA EL MODULO

   l_sql = "SELECT pasnro, pasdesabr, pasasp, pasoblig, pasdepende " &_
           "FROM pasos " &_
		   "WHERE pasos.wiznro = " & l_wiznro & " " &_
		   "ORDER BY pasos.pasorden "

Dim l_dep


rsOpen l_rs, cn, l_sql, 0 
do until l_rs.eof %>
<tr valign="top" style="border-bottom: 1px solid White;" nowrap id="<%= l_rs("pasnro") %>" onmouseover="TRPintar(this,'<%= l_rs("pasnro") %>');"  onmouseout="TRDesPintar(this,'<%= l_rs("pasnro") %>');"  onclick="ear(<%= l_rs("pasnro") %>,0);parent.datos.menunroant.value=parent.datos.menunro.value;parent.datos.menunro.value=<%= l_rs("pasnro") %>;">
	<td width="0" align="left" valign="middle"  nowrap style="background-color: Navy; color: white;"   > 
	
	<% 
	'<!--'Busco el estado del Paso-->
	if isNull(l_rs("pasdepende")) then
	  l_dep = ""
	else
	  l_dep = l_rs("pasdepende")
	end if
	
	if CInt(l_pasos(CInt(l_rs("pasnro")))) = -1 then
	   'El paso ya esta completo
	   if CInt(l_rs("pasoblig")) = -1 then
	      'Si es obligatorio
          response.write "<img title=""Obligatorio"" src=""../images/gen_rep/Oblig.gif"" >"
          response.write "<a href=""JavaScript:SigPaso('" & l_rs("pasasp") & "'," & l_codigo & "," & l_rs("pasnro") & ");"" class=""plano"" style=""color:LightGreen"" title=""Completo"" >"
	      response.write l_rs("pasdesabr") 
		  response.write "</a>"
	   else
	      'Si No es obligatorio
          response.write "<img title=""No Obligatorio"" src=""../images/gen_rep/Obligno.gif"">"
          response.write "<a href=""JavaScript:SigPaso('" & l_rs("pasasp") & "'," & l_codigo & "," & l_rs("pasnro") & ");"" class=""plano"" style=""color:LightGreen"" title=""Completo"" >"
	      response.write l_rs("pasdesabr") 
		  response.write "</a>"
	   end if
	   response.write "</td><td width=""0"" align=""left"" valign=""middle"" style=""background-color: Navy;border-left: 0px; border-bottom: 1px solid White;"" nowrap><img border=""0"" src=""../images/gen_rep/Check.gif"" title=""Completo""></td>"
	else
	   'El paso NO esta completo
	   if CInt(l_rs("pasoblig")) = -1 then
	      'Si es obligatorio
		  if dependenciasListas(l_dep) then
		     'Las dependencias estan resueltas
              response.write "<img title=""Obligatorio"" src=""../images/gen_rep/Oblig.gif"">"
              response.write "<a href=""JavaScript:SigPaso('" & l_rs("pasasp") & "'," & l_codigo & "," & l_rs("pasnro") & ");"" class=""plano"" style=""color:White"" title=""Incompleto"" >"
	          response.write l_rs("pasdesabr") 
		      response.write "</a>"
		  else
  		     'Las dependencias NO estan resueltas
              response.write "<img title=""Obligatorio"" src=""../images/gen_rep/Oblig.gif"">"
              response.write "<a href=""javascript:;"" class=""plano"" style=""color:LightBlue"" title=""Incompleto"" >"
	          response.write l_rs("pasdesabr") 
		      response.write "</a>"
		  end if
	   else
	      'Si No es obligatorio
		  if dependenciasListas(l_dep) then
		     'Las dependencias estan resueltas
              response.write "<img title=""No Obligatorio"" src=""../images/gen_rep/Obligno.gif"">"
              response.write "<a href=""JavaScript:SigPaso('" & l_rs("pasasp") & "'," & l_codigo & "," & l_rs("pasnro") & ");"" class=""plano"" style=""color:White"" title=""Incompleto"" >"
	          response.write l_rs("pasdesabr") 
		      response.write "</a>"
		  else
  		     'Las dependencias NO estan resueltas
              response.write "<img title=""No Obligatorio"" src=""../images/gen_rep/Obligno.gif"">"
              response.write "<a href=""javascript:"" class=""plano"" style=""color:LightBlue"" title=""Incompleto"" >"
	          response.write l_rs("pasdesabr") 
		      response.write "</a>"
		  end if
	   end if
	   response.write "<td width=""10"" style=""background-color: Navy;border-left: 0px; border-bottom: 1px solid White; "" valign=""middle"" align=""left"">&nbsp;</td>"
	end if
	%>
  </tr>
	<% 
	l_rs.MoveNext
loop
l_rs.MoveFirst
response.write "<script>parent.datos.pasonro.value =" & l_rs("pasnro") & "</script>"
l_rs.Close
set l_rs = Nothing
cn.Close
set cn = Nothing
%>
<tr ><td height="100%" colspan="3" style="background-color: Navy; border-top: 0px solid White; border-left: 0px solid White;">&nbsp;</td></tr>
</table>
</body>
</html>
