var App = angular.module('drag-and-drop', ['ngAnimate', 'ngDragDrop']);

App.controller('oneCtrl', function($scope) {
  $scope.list1 = [
    { 'title': 'N', 'drag': true },
    { 'title': 'L', 'drag': true },
    { 'title': 'I', 'drag': true },
    { 'title': 'I', 'drag': true },
    { 'title': 'E', 'drag': true },
    { 'title': 'N', 'drag': true }
  ];
});

App.controller('twoCtrl', function($scope) {
  $scope.list1 = [
    { 'title': 'N', 'drag': true },
    { 'title': 'L', 'drag': true },
    { 'title': 'I', 'drag': true },
    { 'title': 'I', 'drag': true },
    { 'title': 'E', 'drag': true },
    { 'title': 'N', 'drag': true }
  ];
});