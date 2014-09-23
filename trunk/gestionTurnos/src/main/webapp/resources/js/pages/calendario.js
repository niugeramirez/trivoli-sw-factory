function controller($scope, $http) {
	// Se define el Modelo de la Página de Administración de Obras Sociales
	
	// Pagina solicitada al Backend
    $scope.nroPagina = 0;

    // Estado Actual de la Vista
    $scope.estado = 'busy';

    // Ultima Accion solictada por el Usuario
    $scope.ultimaAccion = '';

    // URL base de la Vista
    $scope.url = "/gestionTurnos/protected/calendarios/";

    // Flags diversos que manejan la interacción del Usuario con la Vista
    $scope.errorSubmit = false;
    $scope.errorAccesoIlegal = false;
    $scope.mostrarMensajesUsuario = false;
    $scope.mostrarErrorValidacion = false;
    $scope.mostrarMensajeBusqueda = false;
    $scope.mostrarBotonBuscar = false;
    $scope.mostrarBotonCrear = false;

    // Objetos JSON
    $scope.calendarioActual = {};   
    $scope.turnosActuales = {};
    $scope.recursosActuales = {};
    $scope.recursoActual = {};    
    
    // Filtro de Busqueda   
//    $scope.fechaActual = 	new Date().getDate().toString() + '/'+
//    						(new Date().getMonth()+1)  + '/'+
//    						new Date().getFullYear().toString();
    $scope.fechaActual = '23/09/2014';

   
    
    // Definición de Funciones del Controlador de la Página de Administración de Obras Sociales
/************************************************************************************************************************************************************************/    
    // Funcion que recupera del backend todos los Recursos
    $scope.listarTodo = function () {
    	// Se obtiene la URL actual
        var url = $scope.url;
        
        // Se fija la Ultima Accion Solicitada por el Usuario
        $scope.ultimaAccion = 'list';

        // Se abre el dialogo Loading .....
        $scope.startDialogAjaxRequest("#loadingModalRecursos");

        // Se fijan los parámetros de la llamada al servicio Rest (Página Solicitada por el Usuario)
        var config = {params: {nroPagina: $scope.nroPagina}};        

        // Se realiza un requerimiento HTTP a través de un método GET esperando 2 posibles resultados (callbacks invocados
        // asincronicamente) 
        $http.get(url, config)
            .success(function (data) {
                $scope.finishAjaxRecursos(data, null, false);    
                $scope.listarCalendario();
            })
            .error(function () {
                $scope.estado = 'error';
                $scope.mostrarBotonCrear = false;
            });
        
    };
/************************************************************************************************************************************************************************/    
    // Funcion que recupera del backend todos los Recursos
    $scope.listarCalendario = function () {
    	// Se obtiene la URL actual
        var url = $scope.url +  $scope.recursoActual.id;
        
        // Se fija la Ultima Accion Solicitada por el Usuario
        $scope.ultimaAccion = 'list';

        // Se abre el dialogo Loading .....
        $scope.startDialogAjaxRequest("#loadingModal");

        // Se fijan los parámetros de la llamada al servicio Rest (Página Solicitada por el Usuario)
        var config = {params: {nroPagina: $scope.nroPagina}};
        config.params.idRecurso = $scope.recursoActual.id;
        config.params.fechaTurnos = new Date();
        //config.params.fechaTurnos = $scope.fechaActual;

        // Se realiza un requerimiento HTTP a través de un método GET esperando 2 posibles resultados (callbacks invocados
        // asincronicamente) 
        $http.get(url, config)
            .success(function (data) {
                $scope.finishAjaxCallOnSuccess(data, null, false);    
                $scope.buscarTurnos();
            })
            .error(function () {
                $scope.estado = 'error';
                $scope.mostrarBotonCrear = false;
            });
        
    };
/************************************************************************************************************************************************************************/

    // Funcion que recupera del backend todos los Recursos
    $scope.buscarTurnos = function () {    
	    
		if	($scope.calendarioActual) {
	    	var url = $scope.url + $scope.calendarioActual.recurso.id +'/'+ $scope.calendarioActual.id;
			
			$scope.ultimaAccion = 'list';
			
			$scope.startDialogAjaxRequest("#loadingModalTurno");
			
			var config = {};		
			config.params = {};
			config.params.nroPagina = 0;
			config.params.idCalendario = $scope.calendarioActual.id;
			config.params.idRecurso = $scope.calendarioActual.recurso.id;
	  
			
			$http.get(url, config)
			    .success(function (data) {
			    	$scope.mostrarMensajeBusqueda = true;
			    	finishAjaxTurnos(data, null, false); 
			    })
			    .error(function(data, status, headers, config) {
	                $scope.estado = 'error';
	                $scope.mostrarBotonCrear = false;
			    });    		   
		}
	};
/************************************************************************************************************************************************************************/
    
    finishAjaxTurnos = function (data, modalId, isPagination) {
    	// Se muestran los datos en la Grilla de horarios        
        if (data.cantPaginas > 0) {
            $scope.estado = 'list';               

        }         
        $scope.turnosActuales   = data.registros;
        
        $("#loadingModalTurno").modal('hide');

        if(!isPagination){
            if(modalId){
                $scope.exit(modalId);
            }
        }

        $scope.ultimaAccion = '';
    };
/************************************************************************************************************************************************************************/
     
    $scope.finishAjaxRecursos = function (data, modalId, isPagination) {
    	// Se muestran los datos en la Grilla de turnos    	
    	
        $scope.recursosActuales = data.registros;//$scope.pagina.registros;
        $scope.recursoActual =data.registros[3];// $scope.pagina.registros[0];        
        
        $("#loadingModalRecursos").modal('hide');

        if(!isPagination){
            if(modalId){
                $scope.exit(modalId);
            }
        }

        $scope.ultimaAccion = '';
    };    
/************************************************************************************************************************************************************************/
    function pad(str, len, pad, dir) {
    	var STR_PAD_LEFT = 1;
    	var STR_PAD_RIGHT = 2;
    	var STR_PAD_BOTH = 3;    	

        if (typeof(len) == "undefined") { var len = 0; }
        if (typeof(pad) == "undefined") { var pad = ' '; }
        if (typeof(dir) == "undefined") { var dir = STR_PAD_RIGHT; }

        if (len + 1 >= str.length) {

            switch (dir){

                case STR_PAD_LEFT:
                    str = Array(len + 1 - str.length).join(pad) + str;
                break;

                case STR_PAD_BOTH:
                    var right = Math.ceil((padlen = len - str.length) / 2);
                    var left = padlen - right;
                    str = Array(left+1).join(pad) + str + Array(right+1).join(pad);
                break;

                default:
                    str = str + Array(len + 1 - str.length).join(pad);
                break;

            } // switch

        }

        return str;

    }  
/************************************************************************************************************************************************************************/    
//    // Funcion que elimina un Recurso del backend
//    $scope.eliminarObraSocial = function () {
//    	// Se fija la Ultima Accion Solicitada por el Usuario
//        $scope.ultimaAccion = 'delete';
//
//        // Se obtiene la URL actual, parametrizandola con el ID del Recurso a Eliminar
//        var url = $scope.url + $scope.obraSocial.id;
//
//        // Se abre el dialogo Loading .....
//        $scope.startDialogAjaxRequest();
//
//        var params = {filtroNombre: $scope.filtroNombre, nroPagina: $scope.nroPagina};
//
//        $http({
//            method: 'DELETE',
//            url: url,
//            params: params
//        }).success(function (data) {
//                $scope.resetObraSocial();
//                $scope.finishAjaxCallOnSuccess(data, "#eliminarObrasSocialesDialog", false);
//            }).error(function(data, status, headers, config) {
//                $scope.handleErrorInDialogs(status);
//            });
//    };

/************************************************************************************************************************************************************************/
 
    
//    // Funcion que crea un Recurso en el backend
//    $scope.crearObraSocial = function (crearObraSocialForm) {
//        if (!crearObraSocialForm.$valid) {
//            $scope.mostrarErrorValidacion = true;
//            return;
//        }
//
//        $scope.ultimaAccion = 'create';
//
//        var url = $scope.url;
//
//        var config = {headers: {'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8'}};
//
//        $scope.agregarFiltroBusqueda(config);
//
//        $scope.startDialogAjaxRequest();
//
//        // Se envía el Recurso actual serializado (via funcion $.param de jquery)
//        $http.post(url, $.param($scope.obraSocial), config)
//            .success(function (data) {
//                $scope.finishAjaxCallOnSuccess(data, "#crearObrasSocialesDialog", false);
//            })
//            .error(function(data, status, headers, config) {
//                $scope.handleErrorInDialogs(status);
//            });
//    };

/************************************************************************************************************************************************************************/    
    
//    // Funcion que Actualiza un Recurso en el Backend
//    $scope.updateObraSocial = function (updateObrasocialForm) {
//        if (!updateObrasocialForm.$valid) {
//            $scope.mostrarErrorValidacion = true;
//            return;
//        }
//
//        $scope.ultimaAccion = 'update';
//
//        var url = $scope.url + $scope.obraSocial.id;
//
//        $scope.startDialogAjaxRequest();
//
//        var config = {};
//
//        $scope.agregarFiltroBusqueda(config);
//
//        $http.put(url, $scope.obraSocial, config)
//            .success(function (data) {
//                $scope.finishAjaxCallOnSuccess(data, "#editarObrasSocialesDialog", false);
//            })
//            .error(function(data, status, headers, config) {
//                $scope.handleErrorInDialogs(status);
//            });
//    };
    
    
/************************************************************************************************************************************************************************/    
//    // Funcion que Busca Recursos segun un Filtro de Busqueda en el Backend
//    $scope.buscarObrasSociales = function (buscarObrasSocialesForm, isPagination) {
//        if (!($scope.filtroNombre) && (!buscarObrasSocialesForm.$valid)) {
//            $scope.mostrarErrorValidacion = true;
//            return;
//        }
//
//        $scope.ultimaAccion = 'search';
//
//        var url = $scope.url +  $scope.filtroNombre;
//
//        $scope.startDialogAjaxRequest();
//
//        var config = {};
//
//        if($scope.filtroNombre){
//            $scope.agregarFiltroBusqueda(config);
//        }
//
//        $http.get(url, config)
//            .success(function (data) {
//            	$scope.mostrarMensajeBusqueda = true;
//                $scope.finishAjaxCallOnSuccess(data, "#buscarObrasSocialesDialog", isPagination);
//            })
//            .error(function(data, status, headers, config) {
//                $scope.handleErrorInDialogs(status);
//            });
//    };
    
/************************************************************************************************************************************************************************/    
//    $scope.agregarFiltroBusqueda = function(config) {
//        if(!config.params){
//            config.params = {};
//        }
//
//        config.params.nroPagina = $scope.nroPagina;
//
//        if($scope.filtroNombre){
//            config.params.filtroNombre = $scope.filtroNombre;
//        }
//    };
/************************************************************************************************************************************************************************/
    
    $scope.startDialogAjaxRequest = function (modalId) {
        //TODO ver la posibilidad que quede el modal de ajax unificado para todas las llamadas
    	//TODO ver la posibilidad de usar distinto Scopes para los distintos niveles de llamadas AJAX
    	//TODO ver la unificacion de codigo javascript
    	$scope.mostrarErrorValidacion = false;
        $(modalId).modal('show');
        $scope.estadoAnterior = $scope.estado;
        $scope.estado = 'busy';
    };
    
/************************************************************************************************************************************************************************/    

//    $scope.handleErrorInDialogs = function (status) {
//        $("#loadingModal").modal('hide');
//        
//        $scope.estado = $scope.estadoAnterior;
//
//        // Acceso No Permitido
//        if(status == 403){
//            $scope.errorAccesoIlegal = true;
//            return;
//        }
//
//        $scope.errorSubmit = true;
//        $scope.ultimaAccion = '';
//    };
    
/************************************************************************************************************************************************************************/
    
    $scope.finishAjaxCallOnSuccess = function (data, modalId, isPagination) {
    	// Se muestran los datos en la Grilla de turnos    	
        $scope.populateTable(data);    	     
        
        if (data.totalRegistros > 0) {            
            $scope.calendarioActual = $scope.pagina.registros[0].calendario;             
        }         
        
        $("#loadingModal").modal('hide');

        if(!isPagination){
            if(modalId){
                $scope.exit(modalId);
            }
        }

        $scope.ultimaAccion = '';
    };
    
/************************************************************************************************************************************************************************/    
    $scope.populateTable = function (data) {
    	// Se diferencian los casos de Respuesta con datos vs. Respuesta vacía
        if (data.cantPaginas > 0) {
            $scope.estado = 'list';

            // Se define la Pagina de Datos a mostrar
            $scope.pagina = {registros: data.registros, paginaActual: $scope.nroPagina, cantPaginas: data.cantPaginas, totalRegistros : data.totalRegistros};                

        } 
    };
/************************************************************************************************************************************************************************/     
    $scope.seleccionarCalendario = function (registroActual) {
     
        // Se copia el objeto JSON seleccionado en la grilla al registro actual
        $scope.calendarioActual = angular.copy(registroActual.calendario);                              
        $scope.buscarTurnos();
        
    };
/************************************************************************************************************************************************************************/    

//    $scope.resetObraSocial = function(){
//        $scope.obraSocial = {};
//    };
    

/************************************************************************************************************************************************************************/
//    $scope.exit = function (modalId) {
//        $(modalId).modal('hide');
//        
//        $scope.resetObraSocial();
//        
//        $scope.errorSubmit = false;
//        $scope.errorAccesoIlegal = false;
//        $scope.mostrarErrorValidacion = false;
//    };

/************************************************************************************************************************************************************************/
    
//    $scope.resetearBusqueda = function(){
//        $scope.filtroNombre = "";
//        $scope.nroPagina = 0;
//        $scope.listarTodasLasObrasSociales();
//        $scope.mostrarMensajeBusqueda = false;
//    };
    

/************************************************************************************************************************************************************************/    
//    $scope.cambiarPagina = function (pagina) {
//        $scope.nroPagina = pagina;
//
//        if($scope.filtroNombre){
//            $scope.buscarObrasSociales($scope.filtroNombre, true);
//        } else{
//            $scope.listarTodasLasObrasSociales();
//        }
//    };
/************************************************************************************************************************************************************************/    
    // Codigo de Inicializacion del Controlador de la Página de Administración de Obras Sociales
    $scope.listarTodo();    
}
