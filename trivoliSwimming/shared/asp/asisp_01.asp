<% Option Explicit %>
<html>
<head>
<link href="../css/tablesraul.css" rel="StyleSheet" type="text/css">
<title><%= Session("Titulo")%>Menu - Asistente de Conceptos - RHPro &reg;</title>
<script src="/trivoliSwimming/shared/js/fn_windows.js"></script>
<script src="/trivoliSwimming/shared/js/fn_confirm.js"></script>
<script src="/trivoliSwimming/shared/js/fn_ayuda.js"></script>
<script>
var jsSelRow = null;
var pintado = "#000000" //"Navy";
var despintado = "#7A7A7A"; //#4682B4";

function TRPintar(trs, nrop){
    //alert('mouse over' + trs);
	//trs.childNodes[0].style.backgroundColor = despintado;
	//trs.childNodes[1].style.backgroundColor = despintado;
	trs.style.backgroundColor = despintado;	
	trs.style.cursor = "hand";
}
function TRDesPintar(atrs, nrod){
    //alert('mouse out' + nrod);
	//alert((parent.datos.menunro.value == 151));
	//alert(document.all.tabla.rows(1).cells(3).id);

	//alert("A" + document.all.tabla.rows(1).cells(0).style.backgroundColor);
	//document.all.tabla.rows(1).cells(0).style.backgroundColor = despintado;
	//document.all.tabla.rows(1).cells(1).style.backgroundColor = despintado;
	//document.all.tabla.rows(1).cells(2).style.backgroundColor = despintado;
	//alert("A" + document.all.tabla.rows(1).cells(0).style.backgroundColor);
	
//	alert("B" + document.all.tabla.rows(1).cells(1).style.backgroundColor);
//	alert("C" + document.all.tabla.rows(1).cells(2).style.backgroundColor);	
	

	if (parent.datos.menunro.value == 151){
		document.all.tabla.rows(1).cells(0).style.backgroundColor = despintado;
		document.all.tabla.rows(1).cells(1).style.backgroundColor = pintado;
		document.all.tabla.rows(1).cells(2).style.backgroundColor = pintado;		
	}
	
	if (parent.datos.menunro.value == 152){
		document.all.tabla.rows(1).cells(0).style.backgroundColor = pintado;
		document.all.tabla.rows(1).cells(1).style.backgroundColor = despintado;
		document.all.tabla.rows(1).cells(2).style.backgroundColor = pintado;		
	}
	
	if (parent.datos.menunro.value == 153){
		document.all.tabla.rows(1).cells(0).style.backgroundColor = pintado;
		document.all.tabla.rows(1).cells(1).style.backgroundColor = pintado;
		document.all.tabla.rows(1).cells(2).style.backgroundColor = despintado;		
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
	//alert(parent.ifrm.SigPaso);
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


%>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onload="//ear(parent.datos.menunro.value, 1)">
<table border="5" cellpadding="0" cellspacing="0" height="100%" width="100%" id="tabla" >
  <tr id="5" >
	<th colspan="3"  height="25" style="background-color: #000000; border-bottom: 1px solid White;border-right: 0px solid White;"><b> <%'= l_label & ": " & l_nombre%><%= l_nombre%> </b></th>
  </tr>	

<tr valign="top" style="border-bottom: 1px solid White;" nowrap id="15" >
	<td width="0" align="left" valign="middle"  nowrap id="151" style="background-color: #000000; color: white;" onmouseover="TRPintar(this,'151');"  onmouseout="TRDesPintar(this,'151');"  onclick="ear(151,0);parent.datos.menunroant.value=parent.datos.menunro.value;parent.datos.menunro.value=151;" > 
		<!--<img title="No Obligatorio" src="../images/gen_rep/Obligno.gif">-->&nbsp;&nbsp;&nbsp;<a href="JavaScript:SigPaso('pasos_01.asp',1,151);" class="plano" style="color:White" title="Incompleto" >Paso 1</a>
	</td>	
	<!--	
	<td width="10" style="background-color: Navy;border-left: 0px; border-bottom: 1px solid White; " valign="middle" align="left">&nbsp;</td>
	-->
	
	<td width="0" align="left" valign="middle"  nowrap id="152" style="background-color: #000000; color: white;" onmouseover="TRPintar(this,'152');"  onmouseout="TRDesPintar(this,'152');"  onclick="ear(152,0);parent.datos.menunroant.value=parent.datos.menunro.value;parent.datos.menunro.value=152;" > 
		<!--<img title="No Obligatorio" src="../images/gen_rep/Obligno.gif">-->&nbsp;&nbsp;&nbsp;<a href="JavaScript:SigPaso('pasos_02.asp',1,152);" class="plano" style="color:White" title="Incompleto" >Paso 2</a>
	</td>		
	<!--			
	<td width="10" style="background-color: Navy;border-left: 0px; border-bottom: 1px solid White; " valign="middle" align="left">&nbsp;</td>		
	-->
	<td width="0" align="left" valign="middle"  nowrap id="153" style="background-color: #000000; color: white;" onmouseover="TRPintar(this,'153');"  onmouseout="TRDesPintar(this,'153');"  onclick="ear(153,0);parent.datos.menunroant.value=parent.datos.menunro.value;parent.datos.menunro.value=153;" > 
		<!--<img title="No Obligatorio" src="../images/gen_rep/Obligno.gif">-->&nbsp;&nbsp;&nbsp;<a href="JavaScript:SigPaso('pasos_03.asp',1,153);" class="plano" style="color:White" title="Incompleto" >Paso 3</a>
	</td>		
	<!--					
	<td width="10" style="background-color: Navy;border-left: 0px; border-bottom: 1px solid White; " valign="middle" align="left">&nbsp;</td>		
	-->	
  </tr>  

<!--
<tr valign="top" style="border-bottom: 1px solid White;" nowrap id="153" onmouseover="TRPintar(this,'153');"  onmouseout="TRDesPintar(this,'153');"  onclick="ear(153,0);parent.datos.menunroant.value=parent.datos.menunro.value;parent.datos.menunro.value=153;">
	<td width="0" align="left" valign="middle"  nowrap style="background-color: Navy; color: white;"   > 
		<img title="No Obligatorio" src="../images/gen_rep/Obligno.gif"><a href="JavaScript:SigPaso('pasos_03.asp',1,153);" class="plano" style="color:White" title="Incompleto" >Paso 3</a><td width="10" style="background-color: Navy;border-left: 0px; border-bottom: 1px solid White; " valign="middle" align="left">&nbsp;</td>
  </tr>

<tr valign="top" style="border-bottom: 1px solid White;" nowrap id="154" onmouseover="TRPintar(this,'154');"  onmouseout="TRDesPintar(this,'154');"  onclick="ear(154,0);parent.datos.menunroant.value=parent.datos.menunro.value;parent.datos.menunro.value=154;">
	<td width="0" align="left" valign="middle"  nowrap style="background-color: Navy; color: white;"   > 
		<img title="No Obligatorio" src="../images/gen_rep/Obligno.gif"><a href="JavaScript:SigPaso('pasos_04.asp',1,154);" class="plano" style="color:White" title="Incompleto" >Paso 4</a><td width="10" style="background-color: Navy;border-left: 0px; border-bottom: 1px solid White; " valign="middle" align="left">&nbsp;</td>
  </tr>  
  -->

	

	<% 



'response.write "<script>parent.datos.pasonro.value =" & l_rs("pasnro") & "</script>"
'l_rs.Close
'set l_rs = Nothing
'cn.Close
'set cn = Nothing
%>

<tr id="25"><td height="100%" colspan="3" style="background-color: #000000; border-top: 0px solid White; border-left: 0px solid White;">&nbsp;</td></tr>
</table>
</body>
</html>
