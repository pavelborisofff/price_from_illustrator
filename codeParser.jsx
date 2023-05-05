var scriptVersion = '0.02';
var scriptDate = '2023-02-20';

// default actions
// make coordinates system relative to artboard
app.coordinateSystem = CoordinateSystem.ARTBOARDCOORDINATESYSTEM;
// set document color mode to RGB
app.executeMenuCommand('doc-color-rgb');

// default prefs
var defaultLayerName = 'prices';
var layoutType = 'landscape';
var toPNG = false;
var layoutPrefs = {
  landscape: {
    width: 1920,
    height: 1080
  },
  portrait: {
    width: 1080,
    height: 1920
  }
}

// current document (opened in Illustrator)
var doc = app.activeDocument;
// get path to document
var docPath = doc.path + '/';
// replace spaces with underscores and remove file extension
var docName = doc.name.replace(/(.+)\.[aieps]+$/, "$1").replace(/ +/g, "_");


// convert color as array of three numbers to hex string
// [128, 128, 0] => '#808000'
function rgbToHex(r, g, b) {
  return "#" + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1);
}


// save text file 
function saveToFile(path, data) {
  var fd = new File(path);

  fd.open('w', 'TEXT', 'TEXT');
  fd.lineFeed = 'Unix';
  fd.encoding = 'UTF-8';
  fd.writeln(data);
  fd.close();
}

// save to jpg 
// function saveToJpg(path, i) {
//   var exportOptions = new ExportOptionsJPEG();
//   // exportOptions.saveMultipleArtboards = false;
//   exportOptions.artboardClipping = true;
//   exportOptions.antiAliasing = false;
//   exportOptions.quality = 80; 

//   var type = ExportType.JPEG;

//   exportOptions.artboardRange = abs[i].name;


// compare current artboard with defined sizes depend of layout type
function checkArtoboardSize(i) {
  var artboard = app.activeDocument.artboards[i];

  var width = artboard.artboardRect[2] - artboard.artboardRect[0];
  var height = artboard.artboardRect[1] - artboard.artboardRect[3];

  if (width < height) layoutType = 'portrait';

  var widthCheck = layoutPrefs[layoutType].width;
  var heightCheck = layoutPrefs[layoutType].height;

  if (width !== widthCheck || height !== heightCheck) {
    alert('Check size on artboard ' + (i+1) + ':' + width + 'x' + height + 'px\nShould be ' + widthCheck + 'x' + heightCheck + 'px');
  }

  return true;
}


// create layer with name
function createLayer(layerName) {
  var layer = doc.layers.add();
  layer.name = layerName;
  // layer.locked = true;
  // layer.visible = false;
}


// unused
function getLayerIdByName(layerName) {
  var layers = doc.layers;
  var layerId = false;

  for (var i = 0; i < layers.length; i++) {
    if (layers[i].name === layerName) {
      layerId = i;
      break;
    }
  }

  return layerId;
}

// hide layer by name
// function hideLayer(doc, name, option) {
//   doc.layers.getByName(name).visible = option;
//   redraw();
// };


// check if layer with name exists to prevent duplicates
function checkLayerNameExists(layerName) {
  var layers = doc.layers;
  var layerExists = false;

  for (var i = 0; i < layers.length; i++) {
    if (layers[i].name.toLowerCase() === 'png') {
      toPNG = true;
    }

    if (layers[i].name === layerName) {
      layerExists = true;
      layers[i].name = layerName + '-backup';
    }
  }

  return layerExists;
}


// input prefered name for document
function inputDocName() {
  var input = prompt('Enter document name\nPlease use patern "YYYYMMDD_Ntv_OPTIONS" or "YYYYMMDD_price_OPTIONS" without spaces.', docName);

  if (input === null) {
    alert('Script has been stopped\nPlease try again');
    return false;
  }

  if (input === false || input === '' || !input) {
    alert('Document name is empty\nPlease try again');
    inputDocName();
  }

  docName = input.replace(/ +/g, "_");

  // var regExpRes = docName.match(/(\d{6})_(\w+)(?:_(.+))?/);
  // var regExpRes = docName.match(/(\d{6}_\dtv)(?:_?(.+))?/gi);

  // if (regExpRes && regExpRes.length === 3) {
  //   var baseName = regExpRes[1];
  //   var options = regExpRes[2] || '';
  // }
  
  return true;
};
  

// create new layer if not exists, if exists prompt to input new name
function createOrPromptLayer(layerName) {
  var layerExists = checkLayerNameExists(layerName);

  if (layerExists) {
    var nameOfLayer = layerName;
    while (layerExists) {
      var newName = prompt('Layer "' + nameOfLayer + '" already exists.\nPlease enter a new name:', nameOfLayer);

      if (newName == null || newName == '' || newName == false || !newName) {
        alert('Canceled\nBetter use default name for layer "prices"');
        break;
      }

      layerExists = checkLayerNameExists(newName);
      defaultLayerName = newName;
      nameOfLayer = newName;
    }
  }
  createLayer(layerName);
}


// move selected items to later with name
function moveToLayer(layerName) {
  var layer = doc.layers[layerName];
  var items = doc.selection;

  for (var i = 0; i < items.length; i++) {
    items[i].move(layer, ElementPlacement.PLACEATBEGINNING);
  }
}


// save prices to file with coordinates and styles
function cycleThrowArtboardsTextFrames(i) {
  doc.layers.getByName(defaultLayerName).visible = true;
  redraw();

  doc.artboards.setActiveArtboardIndex(i); // select artboard by index
  doc.artboards[i].selected = true; // select artboard

  var objToSave = [];
  // var artNumber = ('0' + (i + 1)).slice(-2);
  // var regexName = docName.match(/(\d{6}_\d*[a-zA-Z]+)(?:_?(.+))?/gi);

  // if (regexName && regexName.length === 3) {
  //   var artBoardName = regexName.slice(2);
  //   artBoardName.splice(1, 0, artNumber);
  //   artBoardName = artBoardName.join('_').replace(/^(.*)_$/g, '$1');
  // } else {
  //   var artBoardName = docName + '_' + artNumber;
  // }

  // var input = prompt('Enter artboard name\nPlease use patern "YYMMDD_Ntv_0X_OPTIONS" or "YYYYMMDD_price_OPTIONS". Use "_" as a space.\nCancel to use current name.', artBoardName);
  // input ? artBoardName = input.replace(/ +/g, "_") : null;


  // doc.artboards[i].name = artBoardName; // set name for artboard
  var artBoardName = doc.artboards[i].name;
  doc.selection = null; // drop previous selection

  for (var j = 0; j < doc.textFrames.length; j++) {
    var currentTextFrame = doc.textFrames[j];
  
    if (
      currentTextFrame.layer.name !== defaultLayerName &&
      currentTextFrame.layer.visible === true &&
      currentTextFrame.left <= doc.artboards[i].artboardRect[2] &&
      currentTextFrame.left >= doc.artboards[i].artboardRect[0] &&
      currentTextFrame.top >= doc.artboards[i].artboardRect[3] &&
      currentTextFrame.top <= doc.artboards[i].artboardRect[1] &&
      /^\d{4}$/.test(currentTextFrame.contents)
    ) {
      currentTextFrame.selected = true;

      var align = currentTextFrame.textRange.justification.toString().split('.').pop().toLowerCase();
      var text = currentTextFrame.contents;
      var charactersToCheck = currentTextFrame.characters;
      var currentCharacter = charactersToCheck[0];  // get the first character

      try {
        var color = currentCharacter.fillColor; // get character color
      } catch (e) {
        var color = '#000000';
      }

      try {
        var font = currentCharacter.textFont.name; // get character style
      } catch (e) {
        var font = 'CeraKFC_Cyr-CondensedBlack';
      }
      var size = Math.floor(currentCharacter.size);

      // get x coordinate of text frame
      // if align === 'right' add width of text frame
      var x = Math.floor(currentTextFrame.left);
      if (align === 'right') x += Math.floor(currentTextFrame.width);
      
      // get y coordinate of text frame
      var y = Math.abs(Math.floor(currentTextFrame.top));
      
      // convert color to hex
      if (color.typename === 'SpotColor') {
        color = '#C8002D';
      } else if (color.typename === 'RGBColor') {
        color = rgbToHex(color.red, color.green, color.blue);
      }
      // add code from text frame to array
      objToSave.push( text + ';' + font + ';' + size + ';' + color + ';' + x + ';' + y + ';' + align);
    }
  }

  if (objToSave.length === 0) return;

  moveToLayer(defaultLayerName);
  var data = objToSave.join('\r');
  var path = docPath + artBoardName + '.csv'
  saveToFile(path , data);
  
  var type;
  var fileType;
  var exportOptions
  var ext;

  if (!toPNG) {
    fileType = new File(docPath + artBoardName + '.jpg');
    type = ExportType.JPEG;
    ext = 'jpg';
  } else {
    fileType = new File(docPath + artBoardName + '.png');
    type = ExportType.PNG24;
    ext = 'png';
  }

  doc.layers.getByName(defaultLayerName).visible = false;
  redraw();
  // redraw().toString;

  if (!toPNG) {
    exportOptions = new ExportOptionsJPEG();
    exportOptions.artBoardClipping = true;
    exportOptions.antiAliasing = true;
    exportOptions.qualitySetting = 80;
    exportOptions.optimization = true;
    exportOptions.artboardRange = doc.artboards[i].name;  
  } else {
    exportOptions = new ExportOptionsPNG24();
    exportOptions.artBoardClipping = true;
    exportOptions.antiAliasing = true;
    exportOptions.transparency = false;
    exportOptions.artboardRange = doc.artboards[i].name;
  }

  doc.exportFile(fileType, type, exportOptions);

  alert('Saved SCV and ' + ext.toUpperCase() + ' to: ' + docPath + artBoardName);
  
  doc.layers.getByName(defaultLayerName).visible = true;
  redraw();
}


function main() {
  // if (!inputDocName()) return false;

  createOrPromptLayer(defaultLayerName);
  doc.selection = null;

  var errors = [];

  for (var i = 0; i < doc.artboards.length; i++) {
    try {
      checkArtoboardSize(i);
      cycleThrowArtboardsTextFrames(i);
    } catch (e) {
      var errorMessage = i + ':' + doc.artboards[i].name + ' - ' + e.line + ': ' + e;
      errors.push(errorMessage);
      alert(e);
      continue;
    }
  }

  if (errors.length > 0) {
    var data = errors.join('\r');
    saveToFile(docPath + docName + '_errors.log', data);
    alert('Errors occured. Check errors.txt file');
  }
}

main();
