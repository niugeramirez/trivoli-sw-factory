// Autores: C. Testaseca, F. Favre
// Se definieron todas las funciones para el manejo de un Doble Browse.
// Para filtro, deben estar presentes los archivos filtro_doblebrowse.asp, filtro_doblebrowse_num.asp, filtro_doblebrowse_txt.asp, filtro_doblebrowse_fec.asp y filtro_doblebrowse_bool.asp
// Para orden, deben esta presente el archivo orden_doblebrowse
// Deben declararse las variables l_Etiquetas, l_Campos, l_Tipos para filtro y l_EtiquetasOr, l_FuncionesOr para orden
// Debe declararse la constante c_IndiceLado en VScript para indicar en que columna del arreglo Datos esta el indicador de seleccionado o no. Por convencion es x + 1, donde x es max(l_rs(x)) 
 
var filtro_izquierda = "true";
var filtro_derecha = "true";

//-----------------------------------------------------------------------------------------------------------------------------------
// Genera un elemento.
function Opcion(indice){
   	newOp = new Option();
    newOp.value = datos[indice][jsClave1];
	for (k=jsClave1+1;k<=jsClave2;k++){
	    newOp.value = newOp.value + ' - ' + datos[indice][k];
	}
   	newOp.text = datos[indice][jsCampos1];
	for (l=jsCampos1+1;l<=jsCampos2;l++){
		if (datos[indice][l] != null){
			if (datos[indice][l] == -1)
			    newOp.text = newOp.text + ' - Sí';
			else{
				if (datos[indice][l] == 0)
				    newOp.text = newOp.text + ' - No';
				else
				    newOp.text = newOp.text + ' - ' + datos[indice][l];
			}
		}
	}
	newOp.subindice = indice;
	return (newOp);
}   

//-----------------------------------------------------------------------------------------------------------------------------------
// Carga los datos en los select, dependiendo si estan seleccionados o no.
function Cargar(){   
	Vaciar(document.all.nselfil);
	Vaciar(document.all.selfil);
    for (i=0;i<datos.length;i++) { 
		if (datos[i][jsIndiceLado]) 
        	document.all.selfil.options.add(Opcion(i));
		else
			document.all.nselfil.options.add(Opcion(i));
    }

	if (document.all.total)
		document.all.total.value = document.all.selfil.length; 
	if (document.all.filtro)
		document.all.filtro.value = document.all.selfil.length; 
	if (document.all.ntotal)
		document.all.ntotal.value = document.all.nselfil.length; 
	if (document.all.nfiltro)
		document.all.nfiltro.value = document.all.nselfil.length; 
	  
    if (document.all.nselfil.length > 0) document.all.nselfil.selectedIndex = 0;
    if (document.all.selfil.length > 0)  document.all.selfil.selectedIndex = 0;
}

//-----------------------------------------------------------------------------------------------------------------------------------
// Carga los datos en el select indicado por la variable lado.
function CargarLado(lado){  
	var cond;
	var filtro = lado?filtro_derecha:filtro_izquierda;
	if (lado)		// lado derecho
		Vaciar(document.all.selfil);
	else			// lado izquierdo
		Vaciar(document.all.nselfil);

    for (i=0;i<datos.length;i++) { 
		eval ("cond = " + filtro);
		if (datos[i][jsIndiceLado] == lado && cond)
			if (lado) 
	        	document.all.selfil.options.add(Opcion(i));
			else
				document.all.nselfil.options.add(Opcion(i));
	}
	if (lado) {
		if (document.all.selfil.length > 0) document.all.selfil.selectedIndex = 0;
		if (document.all.filtro)
			document.all.filtro.value = document.all.selfil.length;
	}
	else {
		if (document.all.nselfil.length > 0) document.all.nselfil.selectedIndex = 0;
		if (document.all.nfiltro)
			document.all.nfiltro.value = document.all.nselfil.length;
	}
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
// Filtra datos en el select (lado), respetando el orden.
function Filtrar(lado,condicion){
 	var objeto;
 	var cond;
	var i;
	
	if (lado){	// lado derecho 
		objeto = document.all.selfil;
		filtro_derecha = condicion;
	}
	else {		// lado izquierdo
		objeto = document.all.nselfil;
		filtro_izquierda = condicion;		
	}

	i = 0;
    while (i<objeto.length) { 
	    eval("cond = " + condicion);
		if (!cond)
			objeto.remove(i);
		else
			i++;
    }
	if (lado) {
		if (document.all.selfil.length > 0) document.all.selfil.selectedIndex = 0;
		if (document.all.filtro)			document.all.filtro.value = document.all.selfil.length;
	}
	else {
		if (document.all.nselfil.length > 0) document.all.nselfil.selectedIndex = 0;
		if (document.all.nfiltro)			 document.all.nfiltro.value = document.all.nselfil.length;
	}
}

//-----------------------------------------------------------------------------------------------------------------------------------
// Pasa todos los datos desde un select fuente (fuente) a un select destino (destino), respetando el orden del select destino.
function Todos(fuente, destino, lado){
	var selFuente = fuente.selectedIndex;
	var selDestino = ",";
	var asc;
	var Menor;
	var salir;
	var x;
	
	if (lado){
		Menor = jsOrdenDerecho;
		asc = jsAscenDerecho;
	}
	else{
		Menor = jsOrdenIzquierdo;
		asc = jsAscenIzquierdo;
	}
	x = fuente.length;
    for (j=0;j<x;j++){
		selDestino = selDestino + fuente[0].value + ",";
		datos[fuente[0].subindice][jsIndiceLado] = lado;
		i = 0;
		salir = false;
	    while (i<destino.length && !salir) {
			if ( eval(Menor+'(datos[fuente[0].subindice], datos[destino[i].subindice])') == asc){
				if (lado)
					document.all.selfil.add(Opcion(fuente[0].subindice), i);
				else
					document.all.nselfil.add(Opcion(fuente[0].subindice), i);
			    fuente.remove(0);
				salir = true;
			}
			i++;
		}
		if (!salir){
			if (lado)
				document.all.selfil.add(Opcion(fuente[0].subindice), i);
			else
				document.all.nselfil.add(Opcion(fuente[0].subindice), i);
		    fuente.remove(0);
		}
	}

	if (lado) {
		Orden(jsOrdenDerecho, jsAscenDerecho, lado);
		if (document.all.total)
			document.all.total.value = parseInt(document.all.total.value) + parseInt(document.all.nfiltro.value) ;
		if (document.all.filtro)
			document.all.filtro.value = parseInt(document.all.filtro.value) + parseInt(document.all.nfiltro.value) ;
		if (document.all.ntotal)
			document.all.ntotal.value = parseInt(document.all.ntotal.value) - parseInt(document.all.nfiltro.value);
		if (document.all.nfiltro)
			document.all.nfiltro.value = 0;
	}
	else {
		Orden(jsOrdenIzquierdo, jsAscenIzquierdo, lado);
		if (document.all.ntotal)
			document.all.ntotal.value = parseInt(document.all.ntotal.value) + parseInt(document.all.filtro.value) ;
		if (document.all.nfiltro)
			document.all.nfiltro.value = parseInt(document.all.nfiltro.value) + parseInt(document.all.filtro.value) ;
		if (document.all.total)
			document.all.total.value = parseInt(document.all.total.value) - parseInt(document.all.filtro.value);
		if (document.all.filtro)
			document.all.filtro.value = 0;
	}

	Reposicionar(destino, selDestino);
	fuente.selectedIndex = (selFuente==fuente.length)?selFuente - 1:selFuente;
}

//-----------------------------------------------------------------------------------------------------------------------------------
// Pasa un dato desde un select fuente (fuente) a un select destino (destino), respetando el orden del select destino. 
function Uno(fuente,destino,lado){
	var selFuente = fuente.selectedIndex;
	var selDestino = ",";
	var asc;
	var Menor;
	var salir;
	
	if (lado){
		Menor = jsOrdenDerecho;
		asc = jsAscenDerecho;
	}
	else{
		Menor = jsOrdenIzquierdo;
		asc = jsAscenIzquierdo;
	}
	
	while (fuente.selectedIndex != -1 ){
		selDestino = selDestino + fuente[fuente.selectedIndex].value + ",";
		datos[fuente[fuente.selectedIndex].subindice][jsIndiceLado] = lado;
		i = 0;
		salir = false;
		//Inserta en forma intercalada, si en el destino existen elementos
	    while (i<destino.length && !salir) {
			if ( eval(Menor+'(datos[fuente[fuente.selectedIndex].subindice], datos[destino[i].subindice])') == asc){
				if (lado)
					document.all.selfil.add(Opcion(fuente[fuente.selectedIndex].subindice), i);
				else
					document.all.nselfil.add(Opcion(fuente[fuente.selectedIndex].subindice), i);
			    fuente.remove(fuente.selectedIndex);
				salir = true;
			}
			i++;
		}
		//En este caso, no existían elementos en el destino
		if (!salir){
			if (lado)
				document.all.selfil.add(Opcion(fuente[fuente.selectedIndex].subindice), i);
			else
				document.all.nselfil.add(Opcion(fuente[fuente.selectedIndex].subindice), i);
		    fuente.remove(fuente.selectedIndex);
		}
		
		//Dependiendo que lado sea, actualizo los contadores de uno u otro lado
		if (lado) {
			if (document.all.total)
				document.all.total.value ++;
			if (document.all.filtro)
				document.all.filtro.value ++;
			if (document.all.ntotal)
				document.all.ntotal.value --;
			if (document.all.nfiltro)
				document.all.nfiltro.value --;
		}
		else {
			if (document.all.ntotal)
				document.all.ntotal.value ++;
			if (document.all.nfiltro)
				document.all.nfiltro.value ++;
			if (document.all.total)
				document.all.total.value --;
			if (document.all.filtro)
				document.all.filtro.value --;
		}			
	}
	//alert(selDestino);
	Reposicionar(destino, selDestino);
	fuente.selectedIndex = (selFuente==fuente.length)?selFuente - 1:selFuente;
}


//-----------------------------------------------------------------------------------------------------------------------------------
// Setea el focus en un items (codigo) del select (objeto).
function Reposicionar (objeto, codigo){
	objeto.selectedIndex = -1;
	for (i=0;i<objeto.length;i++){
		if (codigo.indexOf(","+objeto[i].value+",") != -1){
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
function Orden(Menor, asc, lado){
	var menor;
	var aux; 
	var a;

	if ((lado && jsOrdenDerecho == Menor && jsAscenDerecho == asc) ||
	   (!lado && jsOrdenIzquierdo == Menor && jsAscenIzquierdo == asc))
	   // si el arreglo ya esta ordenado por el criterio no hacer nada
		return false;

	if (lado) {
		jsOrdenDerecho = Menor;
		jsAscenDerecho  = asc;
		aux = document.all.selfil;
	}
	else {
		jsOrdenIzquierdo = Menor;
		jsAscenIzquierdo  = asc;
		aux = document.all.nselfil;
	}

	var x = aux.length - 1;
  	for (i=0;i<x;i++) { 
		menor = i;
	    for (j=i+1;j<aux.length;j++) {
			if (eval(Menor+'(datos[aux[j].subindice], datos[aux[menor].subindice])') == asc)
				menor = j;
		}
		aux.add(Opcion(aux[menor].subindice), i);
		aux.remove(menor+1);
	}
}