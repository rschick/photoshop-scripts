var template = "~/Desktop/S2S/templates/Student_Card_Template.pdf";
var sheetName	= "student_card";
var saveDir		= "~/Desktop/";
var schoolName = 'School Name';
var year = '2020-2021';
var sourceFolder = new Folder('~/Desktop/Work_Space/RussTestInput/');
var saveFolder = new Folder('~/Desktop/Work_Space/RussTestOutput/');
var logoPath = '~/Desktop/Work_Space/BWC/blue_head/13012042789_Rudi_Beierlein.jpg';

var fontSize = 18;
var photoSize = [25, 33.3];
var photoPosition = [0, 0];
var padding = [2, 2];
var logoSize =[20, 20];

var docRef;
var sheetNumber = 1;

var cardSize = [80, 48];
var columns = [12.7, 104.8, 241.4, 333.5];
var rows = [12.9, 73.6, 165.4, 225.5];

var logoFile;
var logoWidth = new UnitValue(logoSize[0], 'mm');
var logoHeight = new UnitValue(logoSize[1], 'mm');

main();

function parseFilename(filename) {
	var decoded = File.decode(filename);
	var parts = decoded.split('_');
	if (parts.length !== 3) {
		return null;
	}
	return {
		id: parts[0],
		homeRoom: parts[0].substring(2, 4),
		firstName: parts[1],
		lastName: parts[2].substring(0, parts[2].lastIndexOf('.')),
	};
}

function loadLogo() {
	logoFile = open(File(logoPath));
	logoFile.resizeImage(logoWidth);
	logoFile.resizeCanvas(logoWidth, logoHeight, AnchorPosition.MIDDLECENTER);
	logoFile.flatten();
}

function putLogo(left, top, width) {
	app.activeDocument = logoFile;
	logoFile.activeLayer.copy();
	app.activeDocument = docRef;
	
	var layer = docRef.paste();
	var bounds = layer.bounds;
	layer.translate(
		left - bounds[0] + (width - logoWidth) / 2,
		top - bounds[1],
	);
}

function processSingleFile(file, left, top) {
	var info = parseFilename(file.name);
	if (!info) {
		return false;
	}
	
	var img = open(file);
	img.resizeImage(photoSize[0]);
	img.resizeCanvas(photoSize[0], photoSize[1], AnchorPosition.MIDDLECENTER);

	var color = app.backgroundColor;
	color.rgb.red = 30;
	color.rgb.green = 41;
	color.rgb.blue = 83;	

	img.backgroundColor = color;
	img.flatten();
	img.activeLayer.copy();
	img.close(SaveOptions.DONOTSAVECHANGES);
	app.activeDocument = docRef;
	
	var layer = docRef.paste();
	var bounds = layer.bounds;
	layer.translate(left - bounds[0], top - bounds[1]);

	var cardWidth = new UnitValue(cardSize[0], 'mm');
	var cardHeight = new UnitValue(cardSize[1], 'mm');
	var photoWidth = new UnitValue(photoSize[0], 'mm');
	var photoHeight = new UnitValue(photoSize[1], 'mm');
	var paddingX = new UnitValue(padding[0], 'mm');
	var paddingY = new UnitValue(padding[1], 'mm');
	var textLeft = left + photoWidth + paddingX;
	var textWidth = cardWidth - textLeft + left;
	var textHeight = cardHeight;
	var lineHeight = new UnitValue(fontSize, 'pt');
	
	var y = top;
	
	writeText(
		year,
		textLeft,
		y,
		textWidth,
		textHeight,
	);

	y += lineHeight + paddingY;
	
	writeText(
		schoolName,
		textLeft,
		y,
		textWidth,
		textHeight,
	);
	
	y += lineHeight + paddingY;

	if (logoPath) {	
		putLogo(
			textLeft,
			y,
			textWidth,
		);
	}
		
	y = top + cardHeight - lineHeight - paddingY;
	
	writeText(
		info.homeRoom + '   ' + info.firstName + ' ' + info.lastName,
		left,
		y,
		cardWidth,
		cardHeight - y + top,
	);
	
	return true;
}

function processFiles(fileList) {
	var fileIndex = 0;
	var row = 0;
	var column = 0;
	docRef = open(File(template));
	while (fileIndex < fileList.length) {
		var left = new UnitValue(columns[column], 'mm');
		var top = new UnitValue(rows[row], 'mm');
		var file = fileList[fileIndex];
		var result = processSingleFile(file, left, top);
		fileIndex += 1;
		if (fileIndex >= fileList.length) {
			break;
		}
		
		if (!result) {
			continue;
		}
		
		column += 1;
		if (column >= columns.length) {
			row += 1;
			column = 0;
		}
		
		if (row >= rows.length) {
			docRef.flatten();
			saveSheet(sheetNumber);
			sheetNumber += 1;
			row = 0;
			docRef = open(File(template));
		}								
	}
	docRef.flatten();
	saveSheet(sheetNumber);
}


function main() {
	var rulerUnits = app.preferences.rulerUnits;

	app.displayDialogs = DialogModes.NO;
	app.preferences.rulerUnits = Units.MM;

	var sourceFolder = Folder.selectDialog( "Source Folder" );
	if (!sourceFolder) {
		alert("You must select a source folder");
		return;
	}

	var saveFolder = Folder.selectDialog( "Destination Folder" );
	if (!saveFolder) {
		alert("You must select a destination folder");
		return;
	}
	
// 	var logoPath = File.selectDialog("Logo file");
// 	alert(logoPath);
	
	schoolName = prompt('School Name', 'School Name');
	
	if (logoPath) {
		loadLogo();
	}
	
	saveDir = File.decode(saveFolder);

	var fileList = sourceFolder
		.getFiles();

	processFiles(fileList);
	
	if (logoPath) {
		logoFile.close(SaveOptions.DONOTSAVECHANGES);
	}
	
	app.preferences.rulerUnits = rulerUnits;
	app.displayDialogs = DialogModes.ALL;
}


function writeText(text, left, top, width, height) {
	var layer = docRef.artLayers.add();	
	layer.kind = LayerKind.TEXT;
	layer.textItem.kind = TextType.PARAGRAPHTEXT;
	layer.textItem.contents = text;
	layer.textItem.position = [left, top];
	layer.textItem.size = fontSize;
	layer.textItem.width = width;
	layer.textItem.height = height;
	layer.textItem.justification = Justification.CENTER;
}

function saveSheet(n) {
	var filename = sheetName + "_" + n + ".jpeg";
	var path = File(saveDir + "/" + filename);
	var options = new JPEGSaveOptions();
	options.embedColorProfile = true;
	options.quality = 8;
	activeDocument.saveAs(path, options);
	activeDocument.close();
}



