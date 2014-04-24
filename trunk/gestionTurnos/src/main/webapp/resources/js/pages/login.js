function loginController($scope, $location) {
	//$scope: 	 Objeto que define el Area de aplicación del Controller
	//$location: Servicio que expone la URL actual 
	
	// Se determina si se debe mostrar un Error de Login, 
	// en función de pagina a la que se redireccionó al Usuario
    var url = "" + $location.$$absUrl;
    
    // Se define en el Modelo un Flag que determina si mostrar o no un Error de Login
    $scope.displayLoginError = (url.indexOf("error") >= 0);
}