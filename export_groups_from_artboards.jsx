// Export groups from artboards as PNG files
// This script exports all groups within artboards as separate PNG files
// Files are organized in folders named after the artboards

// Check if we have a document open
if (app.documents.length > 0) {
    var doc = app.activeDocument;
    
    // Get system information
    var systemInfo = {
        date: new Date().toLocaleString(),
        user: $.getenv("USER") || $.getenv("USERNAME") || "Unknown",
        photoshop: app.version,
        os: $.os.match(/windows/i) ? "Windows" : "macOS",
        osVersion: $.os
    };
    
    // Store original ruler units and change to pixels
    var originalRulerUnits = app.preferences.rulerUnits;
    app.preferences.rulerUnits = Units.PIXELS;
    
    // Create a base folder for exports
    var baseFolder = new Folder(doc.path + "/Exported_Groups");
    if (!baseFolder.exists) {
        baseFolder.create();
    }
    
    // Initialize report arrays
    var exportedGroups = [];
    var skippedGroups = [];
    var blankGroups = [];
    var unchangedGroups = [];
    var modifiedGroups = [];
    
    // Get all layers in the document
    var layers = doc.layers;
    var artboardCount = 0;
    
    // Store all original visibilities at the start
    var originalVisibilities = [];
    function storeVisibilities(layers) {
        for (var i = 0; i < layers.length; i++) {
            var layer = layers[i];
            originalVisibilities.push({
                layer: layer,
                visible: layer.visible
            });
            // Store visibilities for nested layers
            if (layer.typename === "LayerSet" && layer.layers) {
                storeVisibilities(layer.layers);
            }
        }
    }
    
    // Restore all visibilities at the end
    function restoreVisibilities() {
        for (var i = 0; i < originalVisibilities.length; i++) {
            var item = originalVisibilities[i];
            item.layer.visible = item.visible;
        }
    }
    
    // Function to compare two files byte by byte
    function compareFiles(file1, file2) {
        if (!file1.exists || !file2.exists) return false;
        
        try {
            // Get file sizes
            var size1 = file1.length;
            var size2 = file2.length;
            
            // If sizes are different, files are different
            if (size1 !== size2) return false;
            
            // Open both files in binary mode
            file1.encoding = 'BINARY';
            file2.encoding = 'BINARY';
            file1.open('r');
            file2.open('r');
            
            // Read files in chunks to handle large files
            var chunkSize = 4096;
            var areEqual = true;
            
            while (!file1.eof && areEqual) {
                var chunk1 = file1.read(chunkSize);
                var chunk2 = file2.read(chunkSize);
                
                if (chunk1 !== chunk2) {
                    areEqual = false;
                }
            }
            
            // Close files
            file1.close();
            file2.close();
            
            return areEqual;
        } catch(e) {
            if (file1.exists) file1.close();
            if (file2.exists) file2.close();
            return false;
        }
    }
    
    // Function to check if a group is marked with a specific color
    function getGroupColor(group) {
        try {
            var color = group.layerColor;
            return color;
        } catch(e) {
            return null;
        }
    }
    
    // Function to check if image is blank by checking bounds
    function isLayerEmpty(layer) {
        try {
            var bounds = layer.bounds;
            return bounds[2].value - bounds[0].value <= 1 || bounds[3].value - bounds[1].value <= 1;
        } catch(e) {
            return true;
        }
    }
    
    // Function to check if layer has visible content
    function hasVisibleContent(layer) {
        if (!layer.visible) return false;
        
        if (layer.typename === "LayerSet") {
            for (var i = 0; i < layer.layers.length; i++) {
                if (hasVisibleContent(layer.layers[i])) return true;
            }
            return false;
        }
        
        return !isLayerEmpty(layer);
    }
    
    try {
        // Store original visibilities before starting
        storeVisibilities(layers);
        
        // Process each layer in the document
        for (var i = 0; i < layers.length; i++) {
            var layer = layers[i];
            
            // Check if the layer is an artboard (LayerSet)
            if (layer.typename === "LayerSet") {
                artboardCount++;
                var artboardName = layer.name;
                
                // Get the bounds of the artboard
                var artboardBounds = layer.bounds;
                var left = artboardBounds[0].value;
                var top = artboardBounds[1].value;
                var right = artboardBounds[2].value;
                var bottom = artboardBounds[3].value;
                
                // Create folder for this artboard
                var artboardFolder = new Folder(baseFolder + "/" + artboardName);
                if (!artboardFolder.exists) {
                    artboardFolder.create();
                }
                
                // Hide all layers initially
                for (var k = 0; k < layers.length; k++) {
                    layers[k].visible = false;
                }
                
                // Make the artboard visible
                layer.visible = true;
                
                // Process all sublayers (groups) in this artboard
                for (var j = 0; j < layer.layers.length; j++) {
                    var sublayer = layer.layers[j];
                    
                    // Check if sublayer is a group
                    if (sublayer.typename === "LayerSet") {
                        var groupColor = getGroupColor(sublayer);
                        var groupInfo = {
                            name: sublayer.name,
                            artboard: artboardName,
                            color: groupColor
                        };
                        
                        // Skip if group is marked red and has no content
                        if (groupColor === 'red') {
                            if (!hasVisibleContent(sublayer)) {
                                skippedGroups.push(groupInfo);
                                continue;
                            }
                        }
                        
                        // Hide all groups in the current artboard
                        for (var m = 0; m < layer.layers.length; m++) {
                            var currentLayer = layer.layers[m];
                            if (currentLayer.typename === "LayerSet") {
                                currentLayer.visible = false;
                            }
                            // Hide Crop Mask layer if it exists
                            if (currentLayer.name === "Crop Mask") {
                                currentLayer.visible = false;
                            }
                        }
                        
                        // Show only the current group and its contents
                        sublayer.visible = true;
                        if (sublayer.layers) {
                            for (var l = 0; l < sublayer.layers.length; l++) {
                                var contentLayer = sublayer.layers[l];
                                // Don't show Crop Mask layers within groups
                                if (contentLayer.name !== "Crop Mask") {
                                    contentLayer.visible = true;
                                }
                            }
                        }
                        
                        // Check if group has visible content
                        if (!hasVisibleContent(sublayer)) {
                            blankGroups.push(groupInfo);
                            continue;
                        }
                        
                        // Create the file path
                        var fileName = sublayer.name.replace(/[^a-zA-Z0-9]/g, "_");
                        var filePath = new File(artboardFolder + "/" + fileName + ".png");
                        
                        // Check if file is new
                        var isNewFile = !filePath.exists;
                        
                        // Store original document crop
                        var originalCrop = doc.cropBox;
                        
                        // Set crop to artboard bounds
                        doc.crop([left, top, right, bottom]);
                        
                        // Create export options
                        var exportOptions = new ExportOptionsSaveForWeb();
                        exportOptions.format = SaveDocumentType.PNG;
                        exportOptions.PNG8 = false;
                        exportOptions.transparency = true;
                        exportOptions.interlaced = false;
                        exportOptions.quality = 100;
                        
                        // Export the group
                        doc.exportDocument(filePath, ExportType.SAVEFORWEB, exportOptions);
                        
                        // Add to exported groups with new file status
                        groupInfo.isNew = isNewFile;
                        exportedGroups.push(groupInfo);
                        
                        // Restore original crop
                        doc.cropBox = originalCrop;
                    }
                }
            }
        }
    } finally {
        // Always restore visibilities, even if an error occurs
        restoreVisibilities();
        
        // Restore original ruler units
        app.preferences.rulerUnits = originalRulerUnits;
        
        // Read existing report if it exists
        var existingReport = "";
        var reportFile = new File(baseFolder + "/export_report.txt");
        if (reportFile.exists) {
            reportFile.encoding = "UTF-8";
            reportFile.open("r");
            existingReport = reportFile.read();
            reportFile.close();
        }
        
        // Generate new report
        var report = "Relatório de Exportação\n";
        report += "----------------------\n\n";
        report += "Informações do Sistema:\n";
        report += "Data: " + systemInfo.date + "\n";
        report += "Usuário: " + systemInfo.user + "\n";
        report += "Versão do Photoshop: " + systemInfo.photoshop + "\n";
        report += "Sistema Operacional: " + systemInfo.os + "\n";
        report += "Versão do SO: " + systemInfo.osVersion + "\n\n";
        
        report += "Grupos Exportados (" + exportedGroups.length + "):\n";
        for (var i = 0; i < exportedGroups.length; i++) {
            var group = exportedGroups[i];
            report += "- " + group.artboard + " > " + group.name + 
                     (group.color === 'green' ? " (Pronto)" : "") +
                     (group.isNew ? " (Novo)" : "") + "\n";
        }
        
        report += "\nGrupos em Branco (" + blankGroups.length + "):\n";
        for (var i = 0; i < blankGroups.length; i++) {
            var group = blankGroups[i];
            report += "- " + group.artboard + " > " + group.name + "\n";
        }
        
        report += "\nGrupos Ignorados (" + skippedGroups.length + "):\n";
        for (var i = 0; i < skippedGroups.length; i++) {
            var group = skippedGroups[i];
            report += "- " + group.artboard + " > " + group.name + " (Marcado como não pronto)\n";
        }
        
        // Add separator and previous report if it exists
        if (existingReport) {
            report += "\n\n----------------------------------------\n\n";
            report += existingReport;
        }
        
        // Save report to file
        reportFile.encoding = "UTF-8";
        reportFile.open('w');
        reportFile.write(report);
        reportFile.close();
        
        alert("Exportação concluída!\n\n" +
              "Exportados: " + exportedGroups.length + " grupos\n" +
              "Em branco: " + blankGroups.length + " grupos\n" +
              "Ignorados: " + skippedGroups.length + " grupos\n\n" +
              "Consulte export_report.txt para mais detalhes");
    }
} else {
    alert("Please open a document first!");
} 