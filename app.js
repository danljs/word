document.getElementById('load').addEventListener('click', function (e){
	new Promise(function (resolve, reject){
		var x = document.createElement('INPUT');
		x.style = 'visibility:hidden';
	    x.setAttribute('type', 'file');
	    document.body.appendChild(x);
	    x.addEventListener('change', resolve);
	    x.click();
	    document.body.removeChild(x);
	})
	.then(
		function (e){
			return new Promise( function(resolve, reject){
		    	var file = e.target.files[0];
				if(!file){return;}
				var reader = new FileReader();
				reader.onload = resolve;
				reader.readAsText(file);
			});
    	},
    	function (){
			console.log('Something wrong...');
		}
	)
	.then(
		function (e){
			var lines = e.target.result.split('\n').map(function (c,i){
				return c.indexOf('Elvis') > -1 ? i + 1 : '' ;
			}).filter(function (c,i){return c > 0 })

			var wordApp = new ActiveXObject("Word.Application");
			wordApp.Visible = true;
		    var doc = wordApp.Documents.Add ();
		    var sel = wordApp.Selection;
		    sel.TypeText(lines.join('\n'));
		    doc.save();
		    doc.close();
		},
		function (){
			console.log('Something wrong...');
		}
	)
});