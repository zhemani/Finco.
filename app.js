angular.module('fincoApp', [])

.controller('tickerCtrl', ['$rootScope', '$scope', '$location', '$window', function($rootScope, $scope, $location, $window) {


  $scope.getReport = function () {
      console.log($scope.ticker);

      var str = 'https://dfxievhldg.localtunnel.me/csv/' + $scope.ticker;
      console.log(str);

      $window.open('https://dfxievhldg.localtunnel.me/csv/' + $scope.ticker, '_blank');
 

  }

}])



