function Validaciones_locales_pagos(){

	if ((document.datos_02_pagos.fecha.value == "")&&(!validarfecha(document.datos_02_pagos.fecha))){
		 document.datos_02_pagos.fecha.focus();
		 return false;
	}

	if (Trim(document.datos_02_pagos.idmediodepago.value) == "0"){
		alert("Debe ingresar el Medio de Pago.");
		document.datos_02_pagos.idmediodepago.focus();
		return false;
	}
	if (document.datos_02_pagos.mediodepagoos.value == document.datos_02_pagos.idmediodepago.value)  {
		if (Trim(document.datos_02_pagos.idobrasocial.value) == "0"){
			alert("Debe ingresar la Obra Social.");
			document.datos_02_pagos.idobrasocial.focus();
			return false;
		}
	}

	if ((document.datos_02_pagos.importe.value == "")||(document.datos_02_pagos.importe.value == "0")){
		alert("Debe ingresar un Importe mayor o igual a 0.");
		document.datos_02_pagos.importe.focus();
		return;
	}
	document.datos_02_pagos.importe2.value = document.datos_02_pagos.importe.value.replace(",", ".");
	  
	if (!validanumero(document.datos_02_pagos.importe2, 15, 4)){
			  alert("El Monto no es válido. Se permite hasta 15 enteros y 4 decimales.");	
			  document.datos_02_pagos.importe.focus();
			  document.datos_02_pagos.importe.select();
			  return false;
	}	

	return true;

}




function ctrolmetodopago_Pagos(){
	if (document.datos_02_pagos.mediodepagoos.value == document.datos_02_pagos.idmediodepago.value) {
			//document.datos_02_pagos.idobrasocial.readOnly = false;
			//document.datos_02_pagos.idobrasocial.className = 'habinp';			
			document.datos_02_pagos.idobrasocial.disabled = false;							
		}
		else {
			//document.datos_02_pagos.idobrasocial.readOnly = true;
			//document.datos_02_pagos.idobrasocial.className = 'deshabinp';		
			document.datos_02_pagos.idobrasocial.disabled = true;							
			//document.datos_02_pagos.idobrasocial.value = 0;	
		}	

}
