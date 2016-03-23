var X = XLSX;
var XW = {
	/* worker message */
	msg: 'xlsx',
	/* worker scripts */
	rABS: './xlsxworker2.js',
	norABS: './xlsxworker1.js',
	noxfer: './xlsxworker.js'
};

var rABS = typeof FileReader !== "undefined" && typeof FileReader.prototype !== "undefined" && typeof FileReader.prototype.readAsBinaryString !== "undefined";
if(!rABS) {
	document.getElementsByName("userabs")[0].disabled = true;
	document.getElementsByName("userabs")[0].checked = false;
}

var use_worker = typeof Worker !== 'undefined';
if(!use_worker) {
	document.getElementsByName("useworker")[0].disabled = true;
	document.getElementsByName("useworker")[0].checked = false;
}

var transferable = use_worker;
if(!transferable) {
	document.getElementsByName("xferable")[0].disabled = true;
	document.getElementsByName("xferable")[0].checked = false;
}

var wtf_mode = false;

function fixdata(data) {
	var o = "", l = 0, w = 10240;
	for(; l<data.byteLength/w; ++l) o+=String.fromCharCode.apply(null,new Uint8Array(data.slice(l*w,l*w+w)));
	o+=String.fromCharCode.apply(null, new Uint8Array(data.slice(l*w)));
	return o;
}

function ab2str(data) {
	var o = "", l = 0, w = 10240;
	for(; l<data.byteLength/w; ++l) o+=String.fromCharCode.apply(null,new Uint16Array(data.slice(l*w,l*w+w)));
	o+=String.fromCharCode.apply(null, new Uint16Array(data.slice(l*w)));
	return o;
}

function s2ab(s) {
	var b = new ArrayBuffer(s.length*2), v = new Uint16Array(b);
	for (var i=0; i != s.length; ++i) v[i] = s.charCodeAt(i);
	return [v, b];
}

function xw_noxfer(data, cb) {
	var worker = new Worker(XW.noxfer);
	worker.onmessage = function(e) {
		switch(e.data.t) {
			case 'ready': break;
			case 'e': console.error(e.data.d); break;
			case XW.msg: cb(JSON.parse(e.data.d)); break;
		}
	};
	var arr = rABS ? data : btoa(fixdata(data));
	worker.postMessage({d:arr,b:rABS});
}

function xw_xfer(data, cb) {
	var worker = new Worker(rABS ? XW.rABS : XW.norABS);
	worker.onmessage = function(e) {
		switch(e.data.t) {
			case 'ready': break;
			case 'e': console.error(e.data.d); break;
			default: xx=ab2str(e.data).replace(/\n/g,"\\n").replace(/\r/g,"\\r"); console.log("done"); cb(JSON.parse(xx)); break;
		}
	};
	if(rABS) {
		var val = s2ab(data);
		worker.postMessage(val[1], [val[1]]);
	} else {
		worker.postMessage(data, [data]);
	}
}

function xw(data, cb) {
	transferable = document.getElementsByName("xferable")[0].checked;
	if(transferable) xw_xfer(data, cb);
	else xw_noxfer(data, cb);
}

function get_radio_value( radioName ) {
	var radios = document.getElementsByName( radioName );
	for( var i = 0; i < radios.length; i++ ) {
		if( radios[i].checked || radios.length === 1 ) {
			return radios[i].value;
		}
	}
}

function to_json(workbook) {
	var result = {};
	workbook.SheetNames.forEach(function(sheetName) {
		var roa = X.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
		if(roa.length > 0){
			result[sheetName] = roa;
		}
	});
	return result;
}

function to_csv(workbook) {
	var result = [];
	workbook.SheetNames.forEach(function(sheetName) {
		var csv = X.utils.sheet_to_csv(workbook.Sheets[sheetName]);
		if(csv.length > 0){
			result.push("SHEET: " + sheetName);
			result.push("");
			result.push(csv);
		}
	});
	return result.join("\n");
}

function to_formulae(workbook) {
	var result = [];
	workbook.SheetNames.forEach(function(sheetName) {
		var formulae = X.utils.get_formulae(workbook.Sheets[sheetName]);
		if(formulae.length > 0){
			result.push("SHEET: " + sheetName);
			result.push("");
			result.push(formulae.join("\n"));
		}
	});
	return result.join("\n");
}

var tarea = document.getElementById('b64data');
function b64it() {
	if(typeof console !== 'undefined') console.log("onload", new Date());
	var wb = X.read(tarea.value, {type: 'base64',WTF:wtf_mode});
	process_wb(wb);
}

var scriptLines = {};
var scriptLinesList = [];
var DBG_wb;
var segmentLetters = [
  'D', 'F', 'H', 'J', 'L', 'N', 'P', 'R', 'T', 'V', 'W', 'X', 'Z',
  'AB', 'AD', 'AF'];

function toHtml(scriptLines) {
  var out = '';
  for (var lineIdx=0; lineIdx < scriptLinesList.length; lineIdx++) {
    var key = scriptLinesList[lineIdx];
    var data = scriptLines[key];
    out += '<div class=lineBlock>';
    out += '<div class=filename>' + data.filename + '</div>';
    if (data.characters) {
      out += '<div class=chars>';
      out += data.characters.join(', ');
      out += '</div>';
    }
    out += '<div class=all-notes>';
    if (data.prompt) {
      out += '<div class=prompt>' + data.prompt + '</div>';
    }
    if (data.notes) {
      out += '<div class=notes>' + data.notes + '</div>';
    }
    if (data.disposition) {
      out += '<div class=disposition>' + data.disposition + '</div>';
    }
    out += '</div>';
    out += '<div class=lines>'
    for (var i = 0; i < data.segments.length; i++) {
      out += '<p>' + data.segments[i] + '</p>';
    }
    out += '</div>';

    out += '</div>';
  }
  return out;
}

function process_wb(wb) {
  DBG_wb = wb;
	var output = "";
  var sheet = wb.Sheets[wb.SheetNames[0]];

  function getBlank(cellId) {
    var cell = sheet[cellId];
    if (typeof cell !== 'undefined') {
      return cell.v;
    }
    return null;
  }

  var lines = to_json(wb)[wb.SheetNames[0]];
  var dataLineCount = lines.length - 1;
  var npcName = sheet['M1'].v.replace(/\*/g, '');
  for (var i = 3; i <= (dataLineCount+2); i++) {
    var filename = sheet['B'+i].v;
    var segments = [];
    var prompt = getBlank('A'+i);
    var notes = getBlank('C'+i);
    var disposition = getBlank('E'+i);

    for (var j = 0; j < segmentLetters.length; j++) {
      var segment = sheet[segmentLetters[j]+i];
      if (typeof segment !== 'undefined') {
        segments.push(segment.v);
      }
    }

    if (filename in scriptLines) {
      scriptLines[filename].characters.push(npcName);
    } else {
      scriptLines[filename] = {
        characters: [npcName],
        filename: filename,
        prompt: prompt,
        disposition: disposition,
        notes: notes,
        segments: segments,
      };
      scriptLinesList.push(filename);
    }
    
  }
	// switch(get_radio_value("format")) {
	// 	case "json":
	// 		// output = JSON.stringify(to_json(wb), 2, 2);
	// 		output = JSON.stringify(scriptLines, 2, 2);
	// 		break;
	// 	case "form":
	// 		output = to_formulae(wb);
	// 		break;
	// 	default:
	// 	output = to_csv(wb);
	// }
  output = toHtml(scriptLines);
  out.innerHTML = output;
	// if(out.innerText === undefined) out.textContent = output;
	// else out.innerText = output;
	if(typeof console !== 'undefined') console.log("output", new Date());
}

var drop = document.getElementById('drop');
function handleDrop(e) {
	e.stopPropagation();
	e.preventDefault();
	rABS = document.getElementsByName("userabs")[0].checked;
	use_worker = document.getElementsByName("useworker")[0].checked;
	var files = e.dataTransfer.files;
	var f = files[0];
	{
		var reader = new FileReader();
		var name = f.name;
		reader.onload = function(e) {
			if(typeof console !== 'undefined') console.log("onload", new Date(), rABS, use_worker);
			var data = e.target.result;
			if(use_worker) {
				xw(data, process_wb);
			} else {
				var wb;
				if(rABS) {
					wb = X.read(data, {type: 'binary'});
				} else {
				var arr = fixdata(data);
					wb = X.read(btoa(arr), {type: 'base64'});
				}
				process_wb(wb);
			}
		};
		if(rABS) reader.readAsBinaryString(f);
		else reader.readAsArrayBuffer(f);
	}
}

function handleDragover(e) {
	e.stopPropagation();
	e.preventDefault();
	e.dataTransfer.dropEffect = 'copy';
}

if(drop.addEventListener) {
	drop.addEventListener('dragenter', handleDragover, false);
	drop.addEventListener('dragover', handleDragover, false);
	drop.addEventListener('drop', handleDrop, false);
}


var xlf = document.getElementById('xlf');
function handleFile(e) {
	rABS = document.getElementsByName("userabs")[0].checked;
	use_worker = document.getElementsByName("useworker")[0].checked;
	var files = e.target.files;
	var f = files[0];
	{
		var reader = new FileReader();
		var name = f.name;
		reader.onload = function(e) {
			if(typeof console !== 'undefined') console.log("onload", new Date(), rABS, use_worker);
			var data = e.target.result;
			if(use_worker) {
				xw(data, process_wb);
			} else {
				var wb;
				if(rABS) {
					wb = X.read(data, {type: 'binary'});
				} else {
				var arr = fixdata(data);
					wb = X.read(btoa(arr), {type: 'base64'});
				}
				process_wb(wb);
			}
		};
		if(rABS) reader.readAsBinaryString(f);
		else reader.readAsArrayBuffer(f);
	}
}

if(xlf.addEventListener) xlf.addEventListener('change', handleFile, false);
