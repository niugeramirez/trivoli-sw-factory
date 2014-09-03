function obrasSocialesController($scope, $http) {
	// Se define el Modelo de la Página de Administración de Obras Sociales
	
	// Pagina solicitada al Backend
    $scope.nroPagina = 0;

    // Estado Actual de la Vista
    $scope.estado = 'busy';

    // Ultima Accion solictada por el Usuario
    $scope.ultimaAccion = '';

    // URL base de la Vista
    $scope.url = "/gestionTurnos/protected/obrasSociales/";

    // Flags diversos que manejan la interacción del Usuario con la Vista
    $scope.errorSubmit = false;
    $scope.errorAccesoIlegal = false;
    $scope.mostrarMensajesUsuario = false;
    $scope.mostrarErrorValidacion = false;
    $scope.mostrarMensajeBusqueda = false;
    $scope.mostrarBotonBuscar = false;
    $scope.mostrarBotonCrear = false;

    // Objeto JSON que almacena la obra social actual
    $scope.obraSocial = {};

    // Filtro de Busqueda
    $scope.filtroNombre = "";

    
   
    
    // Definición de Funciones del Controlador de la Página de Administración de Obras Sociales
/************************************************************************************************************************************************************************/    
    // Funcion que recupera del backend todos los Recursos
    $scope.listarTodasLasObrasSociales = function () {
    	// Se obtiene la URL actual
        var url = $scope.url;
        
        // Se fija la Ultima Accion Solicitada por el Usuario
        $scope.ultimaAccion = 'list';

        // Se abre el dialogo Loading .....
        $scope.startDialogAjaxRequest();

        // Se fijan los parámetros de la llamada al servicio Rest (Página Solicitada por el Usuario)
        var config = {params: {nroPagina: $scope.nroPagina}};

        // Se realiza un requerimiento HTTP a través de un método GET esperando 2 posibles resultados (callbacks invocados
        // asincronicamente) 
        $http.get(url, config)
            .success(function (data) {
                $scope.finishAjaxCallOnSuccess(data, null, false);
            })
            .error(function () {
                $scope.estado = 'error';
                $scope.mostrarBotonCrear = false;
            });
    };
/************************************************************************************************************************************************************************/
    
 
    // Funcion que elimina un Recurso del backend
    $scope.eliminarObraSocial = function () {
    	// Se fija la Ultima Accion Solicitada por el Usuario
        $scope.ultimaAccion = 'delete';

        // Se obtiene la URL actual, parametrizandola con el ID del Recurso a Eliminar
        var url = $scope.url + $scope.obraSocial.id;

        // Se abre el dialogo Loading .....
        $scope.startDialogAjaxRequest();

        var params = {filtroNombre: $scope.filtroNombre, nroPagina: $scope.nroPagina};

        $http({
            method: 'DELETE',
            url: url,
            params: params
        }).success(function (data) {
                $scope.resetObraSocial();
                $scope.finishAjaxCallOnSuccess(data, "#eliminarObrasSocialesDialog", false);
            }).error(function(data, status, headers, config) {
                $scope.handleErrorInDialogs(status);
            });
    };

/************************************************************************************************************************************************************************/
 
    
    // Funcion que crea un Recurso en el backend
    $scope.crearObraSocial = function (crearObraSocialForm) {
        if (!crearObraSocialForm.$valid) {
            $scope.mostrarErrorValidacion = true;
            return;
        }

        $scope.ultimaAccion = 'create';

        var url = $scope.url;

        var config = {headers: {'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8'}};

        $scope.agregarFiltroBusqueda(config);

        $scope.startDialogAjaxRequest();

        // Se envía el Recurso actual serializado (via funcion $.param de jquery)
        $http.post(url, $.param($scope.obraSocial), config)
            .success(function (data) {
                $scope.finishAjaxCallOnSuccess(data, "#crearObrasSocialesDialog", false);
            })
            .error(function(data, status, headers, config) {
                $scope.handleErrorInDialogs(status);
            });
    };

/************************************************************************************************************************************************************************/    
    
    // Funcion que Actualiza un Recurso en el Backend
    $scope.updateObraSocial = function (updateObrasocialForm) {
        if (!updateObrasocialForm.$valid) {
            $scope.mostrarErrorValidacion = true;
            return;
        }

        $scope.ultimaAccion = 'update';

        var url = $scope.url + $scope.obraSocial.id;

        $scope.startDialogAjaxRequest();

        var config = {};

        $scope.agregarFiltroBusqueda(config);

        $http.put(url, $scope.obraSocial, config)
            .success(function (data) {
                $scope.finishAjaxCallOnSuccess(data, "#editarObrasSocialesDialog", false);
            })
            .error(function(data, status, headers, config) {
                $scope.handleErrorInDialogs(status);
            });
    };
    
    
/************************************************************************************************************************************************************************/    
    // Funcion que Busca Recursos segun un Filtro de Busqueda en el Backend
    $scope.buscarObrasSociales = function (buscarObrasSocialesForm, isPagination) {
        if (!($scope.filtroNombre) && (!buscarObrasSocialesForm.$valid)) {
            $scope.mostrarErrorValidacion = true;
            return;
        }

        $scope.ultimaAccion = 'search';

        var url = $scope.url +  $scope.filtroNombre;

        $scope.startDialogAjaxRequest();

        var config = {};

        if($scope.filtroNombre){
            $scope.agregarFiltroBusqueda(config);
        }

        $http.get(url, config)
            .success(function (data) {
            	$scope.mostrarMensajeBusqueda = true;
                $scope.finishAjaxCallOnSuccess(data, "#buscarObrasSocialesDialog", isPagination);
            })
            .error(function(data, status, headers, config) {
                $scope.handleErrorInDialogs(status);
            });
    };
    
/************************************************************************************************************************************************************************/    
    $scope.agregarFiltroBusqueda = function(config) {
        if(!config.params){
            config.params = {};
        }

        config.params.nroPagina = $scope.nroPagina;

        if($scope.filtroNombre){
            config.params.filtroNombre = $scope.filtroNombre;
        }
    };
/************************************************************************************************************************************************************************/
    
    $scope.startDialogAjaxRequest = function () {
        $scope.mostrarErrorValidacion = false;
        $("#loadingModal").modal('show');
        $scope.estadoAnterior = $scope.estado;
        $scope.estado = 'busy';
    };
    
/************************************************************************************************************************************************************************/    

    $scope.handleErrorInDialogs = function (status) {
        $("#loadingModal").modal('hide');
        
        $scope.estado = $scope.estadoAnterior;

        // Acceso No Permitido
        if(status == 403){
            $scope.errorAccesoIlegal = true;
            return;
        }

        $scope.errorSubmit = true;
        $scope.ultimaAccion = '';
    };
    
/************************************************************************************************************************************************************************/
    
    $scope.finishAjaxCallOnSuccess = function (data, modalId, isPagination) {
    	// Se muestran los datos en la Grilla de Recursos
        $scope.populateTable(data);
        
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

            if($scope.pagina.cantPaginas <= $scope.pagina.paginaActual){
                $scope.nroPagina = $scope.pagina.cantPaginas - 1;
                $scope.pagina.paginaActual = $scope.pagina.cantPaginas - 1;
            }

            $scope.mostrarBotonCrear = true;
            $scope.mostrarBotonBuscar = true;
        } else {
            $scope.estado = 'noresult';
            $scope.mostrarBotonCrear = true;

            if(!$scope.filtroNombre){
                $scope.mostrarBotonBuscar = false;
            }
        }

        if (data.mensajeAccion || data.mensajeBusqueda) {
            $scope.mostrarMensajesUsuario = $scope.ultimaAccion != 'search';

            $scope.pagina.mensajeAccion = data.mensajeAccion;
            $scope.pagina.mensajeBusqueda = data.mensajeBusqueda;
        } else {
            $scope.mostrarMensajesUsuario = false;
        }
    };
/************************************************************************************************************************************************************************/    
 
    $scope.obraSocialSeleccionado = function (obraSocial) {
        var obraSocialSeleccionado = angular.copy(obraSocial);
        
        // Se copia el objeto JSON seleccionado en la grilla al registro actual
        $scope.obraSocial = obraSocialSeleccionado;
    };    
    

/************************************************************************************************************************************************************************/    

    $scope.resetObraSocial = function(){
        $scope.obraSocial = {};
    };
    

/************************************************************************************************************************************************************************/
    $scope.exit = function (modalId) {
        $(modalId).modal('hide');
        
        $scope.resetObraSocial();
        
        $scope.errorSubmit = false;
        $scope.errorAccesoIlegal = false;
        $scope.mostrarErrorValidacion = false;
    };

/************************************************************************************************************************************************************************/
    
    $scope.resetearBusqueda = function(){
        $scope.filtroNombre = "";
        $scope.nroPagina = 0;
        $scope.listarTodasLasObrasSociales();
        $scope.mostrarMensajeBusqueda = false;
    };
    

/************************************************************************************************************************************************************************/    
    $scope.cambiarPagina = function (pagina) {
        $scope.nroPagina = pagina;

        if($scope.filtroNombre){
            $scope.buscarObrasSociales($scope.filtroNombre, true);
        } else{
            $scope.listarTodasLasObrasSociales();
        }
    };
/************************************************************************************************************************************************************************/    
    // Codigo de Inicializacion del Controlador de la Página de Administración de Obras Sociales
    $scope.listarTodasLasObrasSociales();
}
