function recursosController($scope, $http) {
	// Se define el Modelo de la Página de Administración de Recursos
	
	// Pagina solicitada al Backend
    $scope.nroPagina = 0;

    // Estado Actual de la Vista
    $scope.estado = 'busy';

    // Ultima Accion solictada por el Usuario
    $scope.ultimaAccion = '';

    // URL base de la Vista
    $scope.url = "/gestionTurnos/protected/recursos/";

    // Flags diversos que manejan la interacción del Usuario con la Vista
    $scope.errorSubmit = false;
    $scope.errorAccesoIlegal = false;
    $scope.mostrarMensajesUsuario = false;
    $scope.mostrarErrorValidacion = false;
    $scope.mostrarMensajeBusqueda = false;
    $scope.mostrarBotonBuscar = false;
    $scope.mostrarBotonCrear = false;

    // Objeto JSON que almacena el Recurso actual
    $scope.recurso = {};

    // Filtro de Busqueda
    $scope.filtroDescripcion = "";

    
    
    
    // Definición de Funciones del Controlador de la Página de Administración de Recursos
    
    // Funcion que recupera del backend todos los Recursos
    $scope.listarTodosLosRecursos = function () {
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

    
    // Funcion que elimina un Recurso del backend
    $scope.eliminarRecurso = function () {
    	// Se fija la Ultima Accion Solicitada por el Usuario
        $scope.ultimaAccion = 'delete';

     // Se obtiene la URL actual, parametrizandola con el ID del Recurso a Eliminar
        var url = $scope.url + $scope.recurso.id;

        // Se abre el dialogo Loading .....
        $scope.startDialogAjaxRequest();

        var params = {filtroDescripcion: $scope.filtroDescripcion, nroPagina: $scope.nroPagina};

        $http({
            method: 'DELETE',
            url: url,
            params: params
        }).success(function (data) {
                $scope.resetRecurso();
                $scope.finishAjaxCallOnSuccess(data, "#eliminarRecursoDialog", false);
            }).error(function(data, status, headers, config) {
                $scope.handleErrorInDialogs(status);
            });
    };
    
    
    // Funcion que crea un Recurso en el backend
    $scope.crearRecurso = function (crearRecursoForm) {
        if (!crearRecursoForm.$valid) {
            $scope.mostrarErrorValidacion = true;
            return;
        }

        $scope.ultimaAccion = 'create';

        var url = $scope.url;

        var config = {headers: {'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8'}};

        $scope.agregarFiltroBusqueda(config);

        $scope.startDialogAjaxRequest();

        // Se envía el Recurso actual serializado (via funcion $.param de jquery)
        $http.post(url, $.param($scope.recurso), config)
            .success(function (data) {
                $scope.finishAjaxCallOnSuccess(data, "#crearRecursoDialog", false);
            })
            .error(function(data, status, headers, config) {
                $scope.handleErrorInDialogs(status);
            });
    };
    
    // Funcion que Actualiza un Recurso en el Backend
    $scope.updateRecurso = function (updateRecursoForm) {
        if (!updateRecursoForm.$valid) {
            $scope.mostrarErrorValidacion = true;
            return;
        }

        $scope.ultimaAccion = 'update';

        var url = $scope.url + $scope.recurso.id;

        $scope.startDialogAjaxRequest();

        var config = {};

        $scope.agregarFiltroBusqueda(config);

        $http.put(url, $scope.recurso, config)
            .success(function (data) {
                $scope.finishAjaxCallOnSuccess(data, "#editarRecursoDialog", false);
            })
            .error(function(data, status, headers, config) {
                $scope.handleErrorInDialogs(status);
            });
    };
    
    
    // Funcion que Busca Recursos segun un Filtro de Busqueda en el Backend
    $scope.buscarRecursos = function (buscarRecursosForm, isPagination) {
        if (!($scope.filtroDescripcion) && (!buscarRecursosForm.$valid)) {
            $scope.mostrarErrorValidacion = true;
            return;
        }

        $scope.ultimaAccion = 'search';

        var url = $scope.url +  $scope.filtroDescripcion;

        $scope.startDialogAjaxRequest();

        var config = {};

        if($scope.filtroDescripcion){
            $scope.agregarFiltroBusqueda(config);
        }

        $http.get(url, config)
            .success(function (data) {
            	$scope.mostrarMensajeBusqueda = true;
                $scope.finishAjaxCallOnSuccess(data, "#buscarRecursosDialog", isPagination);
            })
            .error(function(data, status, headers, config) {
                $scope.handleErrorInDialogs(status);
            });
    };
    
    
    $scope.agregarFiltroBusqueda = function(config) {
        if(!config.params){
            config.params = {};
        }

        config.params.nroPagina = $scope.nroPagina;

        if($scope.filtroDescripcion){
            config.params.filtroDescripcion = $scope.filtroDescripcion;
        }
    };
    
    $scope.startDialogAjaxRequest = function () {
        $scope.mostrarErrorValidacion = false;
        $("#loadingModal").modal('show');
        $scope.estadoAnterior = $scope.estado;
        $scope.estado = 'busy';
    };
    
    
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

            if(!$scope.filtroDescripcion){
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
    
     
    $scope.recursoSeleccionado = function (recurso) {
        var recursoSeleccionado = angular.copy(recurso);
        
        // Se copia el objeto JSON seleccionado en la grilla al Recurso actual
        $scope.recurso = recursoSeleccionado;
    };
    
    
    $scope.resetRecurso = function(){
        $scope.recurso = {};
    };
    
    
    $scope.exit = function (modalId) {
        $(modalId).modal('hide');
        
        $scope.resetRecurso();
        
        $scope.errorSubmit = false;
        $scope.errorAccesoIlegal = false;
        $scope.mostrarErrorValidacion = false;
    };
    
    
    $scope.resetearBusqueda = function(){
        $scope.filtroDescripcion = "";
        $scope.nroPagina = 0;
        $scope.listarTodosLosRecursos();
        $scope.mostrarMensajeBusqueda = false;
    };
    
    
    $scope.cambiarPagina = function (pagina) {
        $scope.nroPagina = pagina;

        if($scope.filtroDescripcion){
            $scope.buscarRecursos($scope.filtroDescripcion, true);
        } else{
            $scope.listarTodosLosRecursos();
        }
    };
    
    // Codigo de Inicializacion del Controlador de la Página de Administración de Recursos
    $scope.listarTodosLosRecursos();
}
