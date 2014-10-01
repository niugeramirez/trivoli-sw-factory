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
    
    
    $scope.pagina={};
    $scope.pagina.mensajeAccion = "";
    $scope.pagina.mensajeBusqueda = '';       

    // Objetos JSON
    $scope.calendariosActuales = []; 
    $scope.calendarioActual = null;   
    $scope.turnosActuales = [];
    $scope.recursosActuales = {};
    $scope.recursoActual = {};    
    
    // Filtro de Busqueda   
    $scope.fechaActual = 	new Date().getDate().toString() + '-'+
    						(new Date().getMonth()+1)  + '-'+
    						new Date().getFullYear().toString();
   
    
    // Definición de Funciones del Controlador de la Página de Administración de Turnos
/************************************************************************************************************************************************************************/    
    // Funcion que recupera del backend todos los Recursos
    $scope.Inicializar = function () {
    	// Se obtiene la URL actual
        var url = $scope.url;
        //Se setea el modal de carga AJAX
		loadingId = "#loadingModalRecursos";
        
        // Se fija la Ultima Accion Solicitada por el Usuario
        $scope.ultimaAccion = 'list';

        // Se abre el dialogo Loading .....
		$scope.startDialogAjaxRequest(loadingId);           
        
		var config = {};		
		config.params = {};
		config.params.nroPagina = 0;
		
		//Parametros propios de esta llamada AJAX    
		//NO HAY PARAMETROS PROPIOS EN ESTA LLAMADA

        // Se realiza un requerimiento HTTP a través de un método GET esperando 2 posibles resultados (callbacks invocados asincronicamente) 
        $http.get(url, config)
            .success(function (data) {
                $scope.finishAjaxRecursos(data,loadingId);    
                $scope.listarCalendario();
            })
            .error(function () {
            	$scope.errorAjax();
            });
        
    };
/************************************************************************************************************************************************************************/    
    // Funcion que recupera del backend todos los Recursos
    $scope.listarCalendario = function () {
    	// Se obtiene la URL actual
        var url = $scope.url +  $scope.recursoActual.id;
        //Se setea el modal de carga AJAX
		loadingId = "#loadingModalCalendario";
        
        // Se fija la Ultima Accion Solicitada por el Usuario
        $scope.ultimaAccion = 'list';

        // Se abre el dialogo Loading .....
		$scope.startDialogAjaxRequest(loadingId);        

        // Se fijan los parámetros de la llamada al servicio Rest 
		var config = {};		
		config.params = {};
		config.params.nroPagina = 0;
		
		//Parametros propios de esta llamada AJAX        
        config.params.idRecurso = $scope.recursoActual.id;
        config.params.fechaTurnos = $scope.fechaActual;
        

        // Se realiza un requerimiento HTTP a través de un método GET esperando 2 posibles resultados (callbacks invocados asincronicamente) 
        $http.get(url, config)
            .success(function (data) {
                $scope.finishAjaxCalendario(data,loadingId);    
                $scope.buscarTurnos();
            })
            .error(function () {
            	$scope.errorAjax();
            });
        
    };
/************************************************************************************************************************************************************************/

    // Funcion que recupera del backend todos los Recursos
    $scope.buscarTurnos = function () {    
	    
		if	($scope.calendarioActual!=null) {
			// Se obtiene la URL actual
			var url = $scope.url + $scope.calendarioActual.recurso.id +'/'+ $scope.calendarioActual.id;
			//Se setea el modal de carga AJAX
			loadingId = "#loadingModalTurno";
			
	        // Se fija la Ultima Accion Solicitada por el Usuario
	        $scope.ultimaAccion = 'list';

	        // Se abre el dialogo Loading .....
			$scope.startDialogAjaxRequest(loadingId);        

	        // Se fijan los parámetros de la llamada al servicio Rest 			
			var config = {};		
			config.params = {};
			config.params.nroPagina = 0;
			
			//Parametros propios de esta llamada AJAX
			config.params.idCalendario = $scope.calendarioActual.id;
			config.params.idRecurso = $scope.calendarioActual.recurso.id;
	  
	        // Se realiza un requerimiento HTTP a través de un método GET esperando 2 posibles resultados (callbacks invocados asincronicamente) 			
			$http.get(url, config)
			    .success(function (data) {			    	
			    	 $scope.finishAjaxTurnos(data,loadingId); 
			    })
			    .error(function() {
			    	$scope.errorAjax();
			    });    		   
		}
	};
/************************************************************************************************************************************************************************/
    
	 $scope.finishAjaxTurnos = function (data,loadingId) {  
    	
        $scope.turnosActuales   = data.registros;

        $scope.finishAjaxGral(loadingId,data);
    };
/************************************************************************************************************************************************************************/
     
    $scope.finishAjaxRecursos = function (data,loadingId) {   	
        
    	$scope.recursosActuales = data.registros;
        $scope.recursoActual =data.registros[0];        
        
        $scope.finishAjaxGral(loadingId,data);
    };    

/************************************************************************************************************************************************************************/
    
    $scope.finishAjaxCalendario = function (data,loadingId) { 	     
        
        if (data.totalRegistros > 0) {            
        	$scope.calendarioActual = data.registros[0].calendario;   
        }    
        else	{
        	$scope.calendarioActual = null;     
        }
        $scope.calendariosActuales = data.registros;
        
        $scope.finishAjaxGral(loadingId,data);
    };   
/************************************************************************************************************************************************************************/
    
    $scope.finishAjaxGral = function (loadingId,data) { 	             
        
        $scope.estado = $scope.estadoAnterior;
        $(loadingId).modal('hide');      
        $scope.ultimaAccion = '';
        $scope.pagina.mensajeAccion = data.mensajeAccion;
        $scope.pagina.mensajeBusqueda = data.mensajeBusqueda;           
    };      
/************************************************************************************************************************************************************************/
    
    $scope.startDialogAjaxRequest = function (loadingId) {
        //TODO ver la posibilidad que quede el modal de ajax unificado para todas las llamadas
    	//TODO ver la posibilidad de usar distinto Scopes para los distintos niveles de llamadas AJAX
    	//TODO ver la unificacion de codigo javascript
    	$scope.mostrarErrorValidacion = false;
        $(loadingId).modal('show');
        $scope.estadoAnterior = $scope.estado;
        $scope.estado = 'busy';
    };
       
/************************************************************************************************************************************************************************/     
    $scope.seleccionarCalendario = function (registroActual) {
     
        // Se copia el objeto JSON seleccionado en la grilla al registro actual
        $scope.calendarioActual = angular.copy(registroActual.calendario);                              
        $scope.buscarTurnos();
        
    };
/************************************************************************************************************************************************************************/
    
	$scope.errorAjax = function () {  
        $scope.estado = 'error';     
   };    
   
/************************************************************************************************************************************************************************/    
    // Codigo de Inicializacion del Controlador de la Página de Administración de Obras Sociales
    $scope.Inicializar();    
}
