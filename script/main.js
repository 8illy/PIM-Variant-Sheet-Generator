	let ExcelToJSON = function() {
	
		this.parseExcel = function(file,cb) {
			let reader = new FileReader();
			
			reader.onload = function(e) {
				let data = e.target.result;
				let workbook = XLSX.read(data, {
					type: 'binary'
				});
				cb(workbook);
			};
	
			reader.onerror = function(ex) {
				console.log(ex);
			};
			
			reader.readAsBinaryString(file);
		};
	};

	function handleFileSelect(evt) {
	    let files = evt.target.files; // FileList object
	    let xl2json = new ExcelToJSON();
	    xl2json.parseExcel(files[0],function(workbook){
			let sheets = workbook.Sheets;
			let firstSheet = Object.keys(workbook.Sheets)[0];
	    	let lines =  XLSX.utils.sheet_to_row_object_array(workbook.Sheets[firstSheet]);

			let outputLines = [];

			for(let line of lines){
				//first 3 columns copy to the output line
				let columns = Object.keys(line);
				for(let column in line){
					let value = line[column];
					if(value && value!="NULL" && columns.indexOf(column) > 2){
						outputLines.push({
							"Product" : line[columns[2]],
							"Option Type" : "rectangle",
							"Option Name" : column,
							"Option Value" : value,
							"Option Code" : value,
							"Order" : 1,
							"Parent Order" : 1,
						})
					}
				}
			}
          	
			//sort outputLines by product, optionName, optionValue.
			outputLines.sort(function(a,b){
				return ((a.Product > b.Product)?1:-1)||((a["Option Name"] > b["Option Name"])?1:-1);
			});
			
			//build our 2 sheets.
			let variantOptionLines = outputLines.map(function(e,i,a){
				if(i==0||a[i-1].Product != e.Product){
					return e;
				}
				if(a[i-1]["Option Name"] == e["Option Name"]){
					e["Parent Order"] = a[i-1]["Parent Order"];
					e["Order"] = a[i-1]["Order"] + 1;
				}else{
					e["Parent Order"] = a[i-1]["Parent Order"] + 1;
				}
				return e;
			});
			
			//add blank "header" lines.
			let variantOptionLinesWithHeaders =[];
			for(let lineIndex in variantOptionLines){
				let prevLine = variantOptionLines[lineIndex-1];
				prevLine = prevLine?prevLine:{};
				let line = variantOptionLines[lineIndex];
				let columns = Object.keys(line);
				if(prevLine.Product!=line.Product || prevLine["Option Name"]!=line["Option Name"]){
					variantOptionLinesWithHeaders.push({
						"Product" : line.Product,
						"Option Type" : "rectangle",
						"Option Name" : line["Option Name"],
						"Option Value" : "",
						"Option Code" : "",
						"Order" : "",
						"Parent Order" : line["Parent Order"],
					});
				}
				variantOptionLinesWithHeaders.push(line);
			}
			
			
			//build option names
			let variantOptionNameLines = variantOptionLines.filter(function(e,i,a){
				return a.findIndex(function(elem){
					return elem.Product == e.Product && elem["Option Name"] == e["Option Name"];
				})==i;
			}).map(function(e){
				return {
					"Product" : e.Product,
					"Option Name" : e["Option Name"],
					"Position" : e["Parent Order"],//maybe not needed?
				}
			});
			
			let optionNamesWorksheet = XLSX.utils.json_to_sheet(variantOptionNameLines);
			let optionNamesWorkbook = XLSX.utils.book_new();
			XLSX.utils.book_append_sheet(optionNamesWorkbook, optionNamesWorksheet, "Sheet1");
			
			let optionValuesWorksheet = XLSX.utils.json_to_sheet(variantOptionLinesWithHeaders);
			let optionValuesWorkbook = XLSX.utils.book_new();
			XLSX.utils.book_append_sheet(optionValuesWorkbook, optionValuesWorksheet, "Sheet1");
			
			XLSX.writeFile(optionValuesWorkbook, "Variant Option Values.xlsx", { compression: true });
			XLSX.writeFile(optionNamesWorkbook, "Variant Option Names.xlsx", { compression: true });


			
		});
	}
	
	document.addEventListener("DOMContentLoaded", () => {
		console.log("ready");
		document.getElementById('importSheet').addEventListener('change', handleFileSelect, false);
	});