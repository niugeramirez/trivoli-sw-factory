// Autores: C. Testaseca, F. Favre
// Se definieron todas las funciones para el manejo de un Doble Browse.
// Para filtro, deben estar presentes los archivos filtro_doblebrowse.asp, filtro_doblebrowse_num.asp, filtro_doblebrowse_txt.asp y filtro_doblebrowse_fec.asp
// Para orden, deben esta presente el archivo orden_doblebrowse
// Deben declararse las variables l_Etiquetas, l_Campos, l_Tipos para filtro y l_EtiquetasOr, l_FuncionesOr para orden
// Debe declararse la constante c_IndiceLado en VScript para indicar en que columna del arreglo Datos esta el indicador de seleccionado o no. Por convencion es x + 1, donde x es max(l_rs(x)) 
 

var filtro_izquierda = "true";
var filtro_derecha = "true";

//-----------------------------------------------------------------------------------------------------------------------------------
// Carga los datos en los select, dependiendo si estan seleccionados o no.
function Cargar(){   
	Vaciar(document.all.nselfil);
	Vaciar(document.all.selfil);
    for (i=0;i<datos.length;i++) { 
       	newOp = new Option();
        newOp.value  = datos[i][0];
       	newOp.text = datos[i][1] + " " + datos[i][2];
		if (datos[i][jsIndiceLado]){
        	document.all.selfil.options.add(newOp);
		}
		else {
			document.all.nselfil.options.add(newOp);
		}
      }
	if (document.all.total)
		document.all.total.value = document.all.selfil.length; 
	if (document.all.filtro)
		document.all.filtro.value = document.all.selfil.length; 
	if (document.all.ntotal)
		document.all.ntotal.value = document.all.nselfil.length; 
	if (document.all.nfiltro)
		document.all.nfiltro.value = document.all.nselfil.length; 
	  
     if (nselfil.length > 0) nselfil.selectedIndex = 0;
     if (selfil.length > 0)  selfil.selectedIndex = 0;
}

//-----------------------------------------------------------------------------------------------------------------------------------
// Vacia el select (objeto).
function Vaciar(objeto) {   
	var cont = objeto.length;
    for (i=1;i<=cont;i++){
		objeto.remove(0);
	}    
}

//-----------------------------------------------------------------------------------------------------------------------------------
// Setea los filtros a null.
function LimpiarFiltro(lado) {
	if (lado) 
		filtro_derecha = "true";
	else 
		filtro_izquierda = "true";
}

//-----------------------------------------------------------------------------------------------------------------------------------
// Carga los datos en el select indicado por la variable lado.
function CargarLado(lado){   
	var cond;
	var filtro = lado?filtro_derecha:filtro_izquierda;
	if (lado){		// lado derecho
		Vaciar(document.all.selfil);
	}
	else{			// lado izquierdo
		Vaciar(document.all.nselfil);
	}
    for (i=0;i<datos.length;i++) { 
       	newOp = new Option();
        newOp.value  = datos[i][0];
       	newOp.text = datos[i][1] + " " + datos[i][2];
		eval ("cond = " + filtro);
		if (datos[i][jsIndiceLado] == lado && cond)
			if (lado) 
	        	document.all.selfil.options.add(newOp);
			else
				document.all.nselfil.options.add(newOp);
	}
	if (lado) {
		if (selfil.length > 0) selfil.selectedIndex = 0;
		if (document.all.filtro)
				document.all.filtro.value = selfil.length;
	}
	else {
		if (nselfil.length > 0) nselfil.selectedIndex = 0;
		if (document.all.nfiltro)
				document.all.nfiltro.value = nselfil.length;
	}
}

//-----------------------------------------------------------------------------------------------------------------------------------
// Filtra datos en el select (lado), respetando el orden.
function Filtrar(lado,condicion){
 	var objeto;
 	var cond;
	if (lado){	// lado derecho 
		objeto = document.all.selfil;
		filtro_derecha = condicion;
	}
	else {		// lado izquierdo
		objeto = document.all.nselfil;
		filtro_izquierda = condicion;		
	}
	Vaciar(objeto);
    for (i=0;i<datos.length;i++) { 
	    eval("cond = " + condicion);
		if (datos[i][jsIndiceLado] == lado && cond){
       		newOp = new Option();
	        newOp.value = datos[i][0];
    	   	newOp.text = datos[i][1] + " " + datos[i][2];
			objeto.options.add(newOp);
		}
    }
	if (lado) {
		if (selfil.length > 0) selfil.selectedIndex = 0;
		if (document.all.filtro)
			document.all.filtro.value = selfil.length;
	}
	else {
		if (nselfil.length > 0) nselfil.selectedIndex = 0;
		if (document.all.nfiltro)
			document.all.nfiltro.value = nselfil.length;
		
	}
}

//-----------------------------------------------------------------------------------------------------------------------------------
// Pasa todos los datos desde un select fuente (fuente) a un select destino (destino), respetando el orden del select destino.
function Todos(fuente,destino, lado){
    x=fuente.length;
    for (i=1;i<=x;i++){
        var opcion = new Option();
        opcion.value= fuente[0].value;
        opcion.text  = fuente[ 0].text;
	    for (j=0;j<datos.length;j++) { 
			if (opcion.value == datos[j][0]) {
				datos[j][jsIndiceLado] = lado;
			}
		}
        fuente.remove(0);
        destino.add(opcion);
    }
	if (lado) {
		Orden(jsOrdenDerecho, jsAscenDerecho, lado);
		if (document.all.total)
			document.all.total.value = parseInt(document.all.total.value) + parseInt(document.all.nfiltro.value) ;
		if (document.all.ntotal)
			document.all.ntotal.value = parseInt(document.all.ntotal.value) - parseInt(document.all.nfiltro.value);
		if (document.all.nfiltro)
			document.all.nfiltro.value = 0;
	}
	else {
		Orden(jsOrdenIzquierdo, jsAscenIzquierdo, lado);
		if (document.all.ntotal)
			document.all.ntotal.value = parseInt(document.all.ntotal.value) + parseInt(document.all.filtro.value) ;
		if (document.all.total)
			document.all.total.value = parseInt(document.all.total.value) - parseInt(document.all.filtro.value);
		if (document.all.filtro)
			document.all.filtro.value = 0;
	}

	CargarLado(lado);
}

//-----------------------------------------------------------------------------------------------------------------------------------
// Pasa un dato desde un select fuente (fuente) a un select destino (destino), respetando el orden del select destino. 
function Uno(fuente,destino,lado){
    if (fuente.selectedIndex == -1) return; 
	var selFuente = fuente.selectedIndex;
	var selDestino = ",";
	for (j=0;j<fuente.length;j++){
		if (fuente[j].selected) {
		    var opcion = new Option();		
		    opcion.value= fuente[j].value;
		    opcion.text  = fuente[j].text;
			selDestino = selDestino + opcion.value + ",";
		    for (i=0;i<datos.length;i++) { 
				if (opcion.value == datos[i][0]) {
					datos[i][jsIndiceLado] = lado;
					break;
				}
			}
		    fuente.remove(j);
		    destino.add(opcion);
			j--;
			if (lado) {
				if (document.all.total)
					document.all.total.value ++;
				if (document.all.ntotal)
					document.all.ntotal.value --;
				if (document.all.nfiltro)
					document.all.nfiltro.value --;
					
			}
			else {
				if (document.all.ntotal)
					document.all.ntotal.value ++;
				if (document.all.total)
					document.all.total.value --;
				if (document.all.filtro)
					document.all.filtro.value --;
			}			
		}
	}
	if (lado) {
		Orden(jsOrdenDerecho, jsAscenDerecho, lado);
	}
	else {
		Orden(jsOrdenIzquierdo, jsAscenIzquierdo, lado);
	}			

	
	
	CargarLado(lado);
	Reposicionar(destino, selDestino);
	fuente.selectedIndex = (selFuente==fuente.length)?selFuente - 1:selFuente;
}

//-----------------------------------------------------------------------------------------------------------------------------------
// Setea el focus en un items (codigo) del select (objeto).
function Reposicionar (objeto, codigo){
	objeto.selectedIndex = -1;
	for (i=0;i<objeto.length;i++){
		if (codigo.indexOf(objeto[i].value) != -1){
			objeto[i].selected = true;
		}
	}
}

//-----------------------------------------------------------------------------------------------------------------------------------
// Invierte la seleccion.
function InvertirSeleccion(objeto){
	for (i=0;i<objeto.length;i++)
		objeto[i].selected = !objeto[i].selected;
}

//-----------------------------------------------------------------------------------------------------------------------------------
// Ordena los datos en el select (lado), respetando el filtro.
function Orden(Menor, asc,lado){
	var menor;
	var aux = new Array;

	if (lado) {
		jsOredenDerecho = Menor;
		jsAscenDerecho  = asc;
	}
	else {
		jsOredenIzquierdo = Menor;
		jsAscenIzquierdo  = asc;
	}
	
    for (i=0;i<datos.length;i++) { 
		menor = i;
	    for (j=i;j<datos.length;j++) {
			if (eval(Menor+'(datos[j], datos[menor])') == asc){
				menor = j;
			}
		} 
		aux = datos[i];
		datos[i] = datos[menor];
		datos[menor] = aux;
	}
}

