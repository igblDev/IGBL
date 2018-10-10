$.getScript('https://cdnjs.cloudflare.com/ajax/libs/fullcalendar/3.4.0/fullcalendar.min.js',function(){
  var date = new Date();
  var d = date.getDate();
  var m = date.getMonth();
  var y = date.getFullYear();
  
  $('#calendar').fullCalendar({
    header: {      left: 'prev,next',
      center: 'title',
      right: ''    },
    editable: true,
    events: [
			//OCTOBER SKED
      {        title: 'BOX',
        start: '2017-10-26',
        url: '/box_102617.asp'
      },
      {        title: 'BOX',
				start: '2017-10-28',
        url: '/box_102817.asp'
      },
			{        title: 'BOX',
				start: '2017-10-29',
        url: '/box_102917.asp'
      },
			{
        title: 'BOX',
				start: '2017-10-30',
        url: '/box_103017.asp'
      },

			//NOVEMBER SKED
			{
        title: 'BOX',				start: '2017-11-01',
        url: 'box_110117.asp'
      },
			{
       title: 'BOX',
       start: '2017-11-02',
       url: 'box_110217.asp'
      }, 
			{
        title: 'BOX',
        start: '2017-11-04',
        url: 'box_110417.asp'
      },
			{
        title: 'BOX',
        start: '2017-11-05',
        url: 'box_110517.asp'
      },
      {
        title: 'BOX',
       start: '2017-11-07',
        url: 'box_110717.asp'
      },
      {
        title: 'BOX',
        start: '2017-11-09',
        url: 'box_110917.asp'
      },
      {
        title: 'BOX',
        start: '2017-11-11',
        url: 'box_111117.asp'
      },			
      {
        title: 'BOX',
        start: '2017-11-12',
        url: 'box_111217.asp'
      },
      {
        title: 'BOX',
        start: '2017-11-14',
        url: 'box_111417.asp'
      },
      {
        title: 'BOX',
        start: '2017-11-16',
        url: 'box_111617.asp'
      },
      {
        title: 'BOX',
        start: '2017-11-18',
        url: 'box_111817.asp'
      },
      {
        title: 'BOX',
        start: '2017-11-19',
        url: 'box_111917.asp'
      },
      {
        title: 'BOX',
        start: '2017-11-21',
        url: 'box_112117.asp'
      }, 
      {
        title: 'BOX',
        start: '2017-11-23',
        url: 'box_112317.asp'
      },
      {
        title: 'BOX',
        start: '2017-11-25',
        url: 'box_112517.asp'
      },
      {
        title: 'BOX',
        start: '2017-11-27',
        url: 'box_112717.asp'
      },
      {
        title: 'BOX',
        start: '2017-11-28',
        url: 'box_112817.asp'
      },
      {
        title: 'BOX',
        start: '2017-11-30',
        url: 'box_113017.asp'
      },
			//DECEMBER SKED
			{
        title: 'BOX',
        start: '2017-12-02',
        url: 'box_120217.asp'
      },
			{
        title: 'BOX',
        start: '2017-12-03',
        url: 'box_120317.asp'
      }, 
			{
        title: 'BOX',
        start: '2017-12-05',
        url: 'box_120517.asp'
      },
			{
        title: 'BOX',
        start: '2017-12-07',
        url: 'box_120717.asp'
      },
      {
        title: 'BOX',
        start: '2017-12-09',
        url: 'box_120917.asp'
      },
      {
        title: 'BOX',
         start: '2017-12-10',
        url: 'box_121017.asp'
      },
      {
        title: 'BOX',
        start: '2017-12-12',
        url: 'box_121217.asp'
      },			
      {
        title: 'BOX',
        start: '2017-12-14',
        url: 'box_121417.asp'
      },
      {
        title: 'BOX',
        start: '2017-12-16',
        url: 'box_121617.asp'
      },
      {
        title: 'BOX',
        start: '2017-12-17',
        url: 'box_121717.asp'
      },
      {
        title: 'BOX',
       start: '2017-12-18',
        url: 'box_121817.asp'
      },
      {
        title: 'BOX',
        start: '2017-12-20',
        url: 'box_122018.asp'
      },
      {
        title: 'BOX',
        start: '2017-12-21',
        url: 'box_122117.asp'
      }, 
      {
        title: 'BOX',
        start: '2017-12-23',
        url: 'box_122317.asp'
      },
      {
        title: 'BOX',
       start: '2017-12-26',
        url: 'box_122617.asp'
      },
      {
        title: 'BOX',
        start: '2017-12-28',
        url: 'box_122817.asp'
      },
      {
        title: 'BOX',
        start: '2017-12-30',
        url: 'box_123017.asp'
      },
			//JANUARY SKED
			{
        title: 'BOX',
        start: '2018-01-02',
        url: 'box_010218.asp'
      },
			{
        title: 'BOX',
				start: '2018-01-03',
        url: 'box_010318.asp'
      }, 
			{
        title: 'BOX',
				start: '2018-01-04',
        url: 'box_010418.asp'
      },
			{
        title: 'BOX',
        start: '2018-01-05',
        url: 'box_010518.asp'
      },
      {
        title: 'BOX',
        start: '2018-01-06',
        url: 'box_010618.asp'
      },
      {
        title: 'BOX',
        start: '2018-01-07',
        url: 'box_010718.asp'
      },
      {
        title: 'BOX',
        start: '2018-01-08',
        url: 'box_010818.asp'
      },			
      {
        title: 'BOX',
        start: '2018-01-10',
        url: 'box_011018.asp'
      },
      {
        title: 'BOX',
        start: '2018-01-13',
        url: 'box_011318.asp'
      },
      {
        title: 'BOX',
        start: '2018-01-15',
        url: 'box_011518.asp'
      },
      {
        title: 'BOX',
        start: '2018-01-16',
        url: 'box_011618.asp'
      },
      {
        title: 'BOX',
        start: '2018-01-18',
        url: 'box_011818.asp'
      },
      {
        title: 'BOX',
        start: '2018-01-20',
        url: 'box_012018.asp'
      }, 
      {
        title: 'BOX',
        start: '2018-01-21',
        url: 'box_012118.asp'
      },
      {
        title: 'BOX',
        start: '2018-01-23',
        url: 'box_012318.asp'
      },
      {
        title: 'BOX',
        start: '2018-01-25',
        url: 'box_012518.asp'
      },
      {
        title: 'BOX',
        start: '2018-01-27',
        url: 'box_012718.asp'
      },
      {
        title: 'BOX',
        start: '2018-01-28',
        url: 'box_012818.asp'
      },
      {
        title: 'BOX',
        start: '2018-01-29',
        url: 'box_012918.asp'
      },
			//FEBURARY SKED
			{
        title: 'BOX',
        start: '2018-02-01',
        url: 'box_020118.asp'
      },
			{
        title: 'BOX',
        start: '2018-02-03',
        url: 'box_020318.asp'
      }, 
			{
        title: 'BOX',
        start: '2018-02-04',
        url: 'box_020418.asp'
      },
      {
        title: 'BOX',
        start: '2018-02-06',
        url: 'box_020618.asp'
      },
      {
        title: 'BOX',
        start: '2018-02-08',
        url: 'box_020818.asp'
      },			
      {
        title: 'BOX',
        start: '2018-02-10',
        url: 'box_021018.asp'
      },
      {
        title: 'BOX',
        start: '2018-02-11',
        url: 'box_021118.asp'
      },
      {
        title: 'BOX',
        start: '2018-02-13',
        url: 'box_021318.asp'
      },
      {
        title: 'BOX',
       start: '2018-02-15',
        url: 'box_021518.asp'
      },
      {
        title: 'BOX',
        start: '2018-02-24',
        url: 'box_022418.asp'
      },
      {
        title: 'BOX',
        start: '2018-02-25',
        url: 'box_022518.asp'
      }, 
      {
        title: 'BOX',
        start: '2018-02-26',
        url: 'box_022618.asp'
      },
      {
        title: 'BOX',
        start: '2018-02-27',
        url: 'box_022718.asp'
      },	  
			//March SKED
			{
        title: 'BOX',
        start: '2018-03-01',
        url: 'box_030118.asp'
      },
			{
        title: 'BOX',
        start: '2018-03-03',
        url: 'box_030318.asp'
      }, 
			{
        title: 'BOX',
        start: '2018-03-04',
        url: 'box_030418.asp'
      },
      {
        title: 'BOX',
        start: '2018-03-05',
        url: 'box_030518.asp'
      },
      {
        title: 'BOX',
        start: '2018-03-06',
        url: 'box_030618.asp'
      },			
      {
        title: 'BOX',
        start: '2018-03-08',
        url: 'box_030818.asp'
      },
      {
        title: 'BOX',
        start: '2018-03-10',
        url: 'box_031018.asp'
      },
	  {
        title: 'BOX',
        start: '2018-03-11',
        url: 'box_031118.asp'
      },
      {
        title: 'BOX',
        start: '2018-03-13',
        url: 'box_031318.asp'
      },
      {
        title: 'BOX',
       start: '2018-03-15',
        url: 'box_031518.asp'
      },
      {
        title: 'BOX',
       start: '2018-03-17',
        url: 'box_031718.asp'
      },
      {
        title: 'BOX',
       start: '2018-03-18',
        url: 'box_031818.asp'
      },
      {
        title: 'BOX',
       start: '2018-03-19',
        url: 'box_031918.asp'
      },
      {
        title: 'BOX',
       start: '2018-03-20',
        url: 'box_032018.asp'
      },
      {
        title: 'BOX',
       start: '2018-03-21',
        url: 'box_032118.asp'
      },
      {
        title: 'BOX',
       start: '2018-03-22',
        url: 'box_032218.asp'
      },	  
      {
        title: 'BOX',
       start: '2018-03-24',
        url: 'box_032418.asp'
      },	  	  
      {
        title: 'BOX',
        start: '2018-03-26',
        url: 'box_032618.asp'
      },
      {
        title: 'BOX',
        start: '2018-03-28',
        url: 'box_032818.asp'
      }, 
      {
        title: 'BOX',
        start: '2018-03-29',
        url: 'box_032918.asp'
      },
		]
  });
})