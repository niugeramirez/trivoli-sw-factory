// esta funcion deshabilita todos los objetos de una ventana.
function sololectura(){
	// Deshabilito todos los elementos mencionados en cad.
	var allTRs; var i; var r;var a;var j;var cad;
	var arreglo = new Array();
	cad = "INPUT,TEXTAREA,IMG";
	arreglo = cad.split(',');
	for (a=0; a <= arreglo.length -1 ; a++){
		allTRs = document.getElementsByTagName(arreglo[a]);
		for(i=0; i< allTRs.length; i++){
	    	r = allTRs.item(i);
			r.disabled = true;
		}		
	}
	// Deshabilito los doble click de los select
	allTRs = document.getElementsByTagName("SELECT");
	for(i=0; i< allTRs.length; i++){
	    	r = allTRs.item(i);
			r.ondblclick = "";
	}
	
	// Deshabilito todos los Links(A).
	allTRs = document.getElementsByTagName("A");
	for(i=0; i< allTRs.length; i++){
	   	r = allTRs.item(i);
		switch(allTRs.item(i).outerText){
			case "Aceptar":
				r.innerText = "Salir";
				r.href = "Javascript:window.close();";
				break;
				
			case "Guardar":
				r.innerText = "Salir";
				r.href = "Javascript:window.close();";
				break;

			case "Cancelar" :
				r.href = "#";
				r.style.visibility = "hidden";
				break;
				
			case "Confirmar" :
				r.href = "#";
				r.onclick = "";
				r.className = "sidebtnDSB";
				break;
				
			case "Orden":
				break;
				
			case "Filtro":
				break;
				
			case "Excel":
				break;

			case "Modifica":
				break;
				
			case "Calendarios":
				break;

			case "Ayuda":
				break;
				
			default :
				r.href = "#";
				r.className = "sidebtnDSB";
				break;
		}
	}
}
