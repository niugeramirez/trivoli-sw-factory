function LocationController($scope, $location) {
    if($location.$$absUrl.lastIndexOf('/recursos') > 0){
        $scope.activeURL = 'recursos';
    } else{
    	if($location.$$absUrl.lastIndexOf('/modelos') > 0){
    		$scope.activeURL = 'modelos';
    	}
    	else{
    		$scope.activeURL = 'home';
    	}
    }
}