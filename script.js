$( document ).ready(function() {
	$('#export').css('visibility', 'hidden'); 
	$("#file-upload").bind("change", function(){
		var fileName = ''; 
		fileName = $(this).val(); 
		$('#file-selected').html(fileName.replace("C:\\fakepath\\", ""));
	});
});
var xcelData;
var sheet_name;

async function search(){
	$('#export').css('visibility', 'hidden'); 
	sheet_name="";
	xcelData="";
	if(document.getElementById("file-upload").files.length == 0 ){
		alert("File Not Selected");
	}else if($('#sheet-no').val().length === 0){
		alert("Sheet Name Not Selected");
	}else if($('#search-tb').val().length === 0){
		alert("Text Not Selected");
	}else{
		$("#dataTbl").html("");
		var flag = await(checkSheet());
		if(flag){
			searchTxt();
			if($("#dataTbl tr").length==""){
				alert("Text not found");
			}
		}else{
			alert("Sheet not found");
		}
	}
}
	
function checkSheet(){
	return new Promise((resolve, reject) => {
		var fileUpload = $("#file-upload")[0];
		var reader = new FileReader();
		var flag=false;
		fileUpload.value.toLowerCase();
		//file reader onload function is an asynchronous function so we wait for promise.
		if (reader.readAsBinaryString) {
			reader.onload = function (e) {
				data= e.target.result;
				//reading excel file in binary format
				var workbook = XLSX.read(data, {
					type: 'binary'
				});
				xcelData=data;
				//checking if sheet exists
				for (let x of workbook.SheetNames){
					if( x.toLowerCase() === ($('#sheet-no').val().toLowerCase()) ){
						sheet_name=x;
						flag=true; 
					}
				}
				resolve(flag);
			}
			reader.readAsBinaryString(fileUpload.files[0]);
		}
	}).then(function(result) {
		return result;
		});			
}

function removeStubs(wb) {
	Object.values(wb.Sheets).forEach(ws => {
		Object.values(ws).filter(v => v.t === 'z').forEach(v => Object.assign(v,{t:'s',v:''}));
	});
	return wb;
}		
function searchTxt(){

	var workbook =  removeStubs( XLSX.read(xcelData, {type: 'binary' ,sheetStubs:true}) );
	var json_object;
	var txt=$('#search-tb').val().toLowerCase();
	var NoOfRows=0;
	//	workbook.SheetNames.forEach(function(sheetName) { //if we want to check in aLL sheets we use this for Loop
		// Here is your object
		// var XL_row_object = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheet_name]);
		// json_object = JSON.stringify(XL_row_object);// Converting JSON-encoded string to JS object
		// var jsonObj = JSON.parse(json_object);

		//now using different function
		var jsonObj = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name]);
		if(Object.keys(jsonObj).length!=0){ 
			var objKeys = Object.keys(jsonObj);
			for(var i = 0; i < objKeys.length; i++)
			{
				var objValue = Object.values(jsonObj[i]);
				var str;
				for(var j = 0; j < objValue.length; j++){
					if(objValue[j]!=null){
						str=objValue[j].toLowerCase();
					}
					if(str.includes(txt)){
						//console.log(str);
						getRow(i,jsonObj,NoOfRows);
						NoOfRows++;
					}
					// if(txt==jsonObj[i][objKeys[j]]){
						// getRow(i,jsonObj,NoOfRows); 
						// NoOfRows++; 
					// } 
				}
			}
		}else{
			alert("Sheet is empty");
		}
				
		//	});
}
	
function getRow(row,jsonObj,NoOfRows){
	$('#export').css('visibility', 'visible'); 
	var objKeys = Object.keys(jsonObj[0]);		//console.log(objKeys);
	var objValue = Object.values(jsonObj[row]);
	// add table header
	// if(NoOfRows==0){ -->
		// for(var j = 0; j < objKeys.length; j++){ -->
			// $("#dataTbl").append("<td><b>"+objKeys[j]+"</b></td>"); -->
		// }
	// }
	
	//table header is not being added because it is assumed that column heading is in the first row but if it is in another row after some headings it will not get correct data.
	//adding number to columns
	
	// if(NoOfRows==0){
		// for(var j = 0; j < objValue.length; j++){
			// $("#dataTbl").append("<td><b>"+j+"</b></td>");
		// }
	// }
	$("#dataTbl").append("<tr style='padding:0%;'><td style='padding:0%;'></td></tr>");
	// add table rows
	for(var j = 0; j < objValue.length; j++){
		//adding row in a table
		if(objValue[j]==null){
			$("#dataTbl").append("<td>"+'&nbsp;'+"</td>");
		}else{
			$("#dataTbl").append("<td>"+objValue[j]+"</td>");
		}
	}			
}
		
/* HTML to Microsoft Word Export Demo 
	* This code demonstrates how to export an html element to Microsoft Word
	* with CSS styles to set page orientation and paper size.
	* Tested with Word 2010, 2013 and FireFox, Chrome, Opera, IE10-11
	* Fails in legacy browsers (IE<10) that lack window.Blob object
	code from https://stackoverflow.com/questions/36330859/export-html-table-as-word-file-and-change-file-orientation/36337284
*/

function createReport(){
	window.export.onclick = function() {
	if (!window.Blob) {
		alert('Your legacy browser does not support this action.');
		return;
	}
	var html, link, blob, url, css;
	// EU A4 use: size: 841.95pt 595.35pt;
	// US Letter use: size:11.0in 8.5in; 
	css = (
		'<style>' +
		'@page WordDoc{size: 841.95pt 595.35pt;mso-page-orientation: portrait;}' +
		'div.WordDoc {page: WordDoc;}' +
		'table{border-collapse:collapse;}td{border:1px gray solid;width:5em;padding:2px;}'+
		'</style>'
	);
	html = window.docx.innerHTML;
	blob = new Blob(['\ufeff', css + html], {
		type: 'application/msword'
	});
	url = URL.createObjectURL(blob);
	link = document.createElement('A');
	link.href = url;
	// Set default file name. 
	// Word will append file extension - do not add an extension here.
	link.download = 'Report';   
	document.body.appendChild(link);
	if (navigator.msSaveOrOpenBlob ) navigator.msSaveOrOpenBlob( blob, 'Document.doc'); // IE10-11
	else link.click();  // other browsers
		document.body.removeChild(link);
	};	
		
}	
		