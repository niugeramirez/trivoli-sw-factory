/*
Archivo        : fn_sel_multiple.js
Descripcion    : Modulo con las funciones para implementar la seleccion multiplees.
Creador        : Scarpa D.
Fecha Creacion : 25/08/2003
*/

var listaDeDatos;
var listaDeDatosTodos;

/* Funcion que guarda los objetos del formulario que van a guardar los datos */
function setearObjDatos(objFormLista, objFormTodos){
   listaDeDatos = objFormLista;
   listaDeDatosTodos = objFormTodos;    
}

/* Funcion que selecciona los datos en la tabla y guarda la clave de la fila en la lista */
function Seleccionar(fila,cabnro)
{
 if (fila.className == "SelectedRow"){
   fila.className = "MouseOutRow"; 
   eliminarDeLista(cabnro);   
 }else{
   fila.className = "SelectedRow"; 
   agregarALista(cabnro);   
 }
}

/* Selecciona todos los elementos de la lista y de la tabla */
function selectTodos(){
  var allTRs = document.getElementsByTagName("tr");
  var i;
  var r;
  
  for(i=1; i< allTRs.length; i++){
     r = allTRs.item(i);
	 r.className = "SelectedRow"; 
  }

  listaDeDatos.value = listaDeDatosTodos.value;
}

/* Deselecciona todos los elementos de la lista y de la tabla */
function selectNinguno(){
  var allTRs = document.getElementsByTagName("tr");
  var i;
  var r;
  
  for(i=0; i< allTRs.length; i++){
     r = allTRs.item(i);
	 r.className = "MouseOutRow"; 
  }
  
  listaDeDatos.value = "";  
}

/* Se fija si existe una clave en la lista */
function existe(cabnro){
  var arreglo;
  var nro;
  var pos=0;
  var lista= listaDeDatos.value;
  
  if ((lista!="") && (lista!=null)){
      arreglo = lista.split(',');
	  while (pos < arreglo.length){
	     nro = arreglo[pos];
		 pos++;
		 if (nro == cabnro){
		   return 1;
		 }
	  }
  }
  return 0;  
}

/* Agrega una clave a la lista si no existe */
function agregarALista(cabnro){
  if (!existe(cabnro)){
    if (listaDeDatos.value == ''){
       listaDeDatos.value = cabnro	
	}else{
  	   listaDeDatos.value = listaDeDatos.value + ',' + cabnro
	}    
  }
}

/* Elimina una clave de la lista si no existe */
function eliminarDeLista(cabnro){
  var arreglo;
  var nro;
  var pos=0;
  var listatmp = "";
  var lista= listaDeDatos.value;
  
  if ((lista!="") && (lista!=null)){
	  arreglo = lista.split(',');
	  while (pos < arreglo.length){
	     nro = arreglo[pos];
		 pos++;
		 if (nro != cabnro){
			 if (listatmp == ''){
		  	   listatmp = nro;
			 }else{
			   listatmp = listatmp + ',' + nro;  
			 }
		 }
	  }
	  listaDeDatos.value = listatmp;
  }
}

