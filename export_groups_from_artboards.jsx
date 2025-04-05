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
    
    // Function to calculate MD5 hash of a file
    function calculateMD5(file) {
        try {
            if (!file.exists) return null;
            
            // Read file as binary
            file.encoding = 'BINARY';
            file.open('r');
            var content = file.read();
            file.close();
            
            // Simple hash function (since we can't use MD5 directly in ExtendScript)
            var hash = 0;
            for (var i = 0; i < content.length; i++) {
                hash = ((hash << 5) - hash) + content.charCodeAt(i);
                hash = hash & hash; // Convert to 32-bit integer
            }
            return hash.toString();
        } catch(e) {
            return null;
        }
    }
    
    // Function to check if file should be exported
    function shouldExportGroup(group, filePath, artboardBounds) {
        var file = new File(filePath);
        
        // If file doesn't exist, should export
        if (!file.exists) {
            return { shouldExport: true, reason: "New file" };
        }
        
        // Store current visibility states and crop
        var tempVisibilities = [];
        var originalCrop = doc.cropBox;
        
        function storeTemp(layer) {
            if (layer.typename === "LayerSet" && layer.layers) {
                for (var i = 0; i < layer.layers.length; i++) {
                    tempVisibilities.push({
                        layer: layer.layers[i],
                        visible: layer.layers[i].visible
                    });
                    storeTemp(layer.layers[i]);
                }
            }
        }
        storeTemp(group);
        
        // Calculate hash of existing file
        var existingHash = calculateMD5(file);
        
        // Export to temporary file
        var tempPath = filePath + ".temp";
        var tempFile = new File(tempPath);
        
        // Set crop to artboard bounds
        doc.crop([artboardBounds[0], artboardBounds[1], artboardBounds[2], artboardBounds[3]]);
        
        // Set up temporary export
        var exportOptions = new ExportOptionsSaveForWeb();
        exportOptions.format = SaveDocumentType.PNG;
        exportOptions.PNG8 = false;
        exportOptions.transparency = true;
        exportOptions.interlaced = false;
        exportOptions.quality = 100;
        
        // Export temporary file
        doc.exportDocument(tempFile, ExportType.SAVEFORWEB, exportOptions);
        
        // Calculate hash of new file
        var newHash = calculateMD5(tempFile);
        
        // Clean up temp file
        tempFile.remove();
        
        // Restore crop
        doc.cropBox = originalCrop;
        
        // Restore visibility states
        for (var i = 0; i < tempVisibilities.length; i++) {
            var item = tempVisibilities[i];
            item.layer.visible = item.visible;
        }
        
        // Compare hashes
        if (existingHash !== newHash) {
            return { shouldExport: true, reason: "Content modified" };
        }
        
        return { shouldExport: false, reason: "No changes" };
    }
    
    // Store original visibilities before starting
    storeVisibilities(layers);
    
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
                        
                        // Check if we should export this group
                        var exportCheck = shouldExportGroup(sublayer, filePath, [left, top, right, bottom]);
                        
                        if (exportCheck.shouldExport) {
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
                            
                            // Add to exported and modified groups
                            exportedGroups.push(groupInfo);
                            groupInfo.reason = exportCheck.reason;
                            modifiedGroups.push(groupInfo);
                            
                            // Restore original crop
                            doc.cropBox = originalCrop;
                        } else {
                            // Add to unchanged groups
                            groupInfo.reason = exportCheck.reason;
                            unchangedGroups.push(groupInfo);
                        }
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
        var report = "Export Report\n";
        report += "-------------\n\n";
        report += "System Information:\n";
        report += "Date: " + systemInfo.date + "\n";
        report += "User: " + systemInfo.user + "\n";
        report += "Photoshop Version: " + systemInfo.photoshop + "\n";
        report += "Operating System: " + systemInfo.os + "\n";
        report += "OS Version: " + systemInfo.osVersion + "\n\n";
        
        report += "Modified Groups (" + modifiedGroups.length + "):\n";
        for (var i = 0; i < modifiedGroups.length; i++) {
            var group = modifiedGroups[i];
            report += "- " + group.artboard + " > " + group.name + 
                     (group.color === 'green' ? " (Ready)" : "") +
                     " - " + group.reason + "\n";
        }
        
        report += "\nUnchanged Groups (" + unchangedGroups.length + "):\n";
        for (var i = 0; i < unchangedGroups.length; i++) {
            var group = unchangedGroups[i];
            report += "- " + group.artboard + " > " + group.name + "\n";
        }
        
        report += "\nBlank Groups (" + blankGroups.length + "):\n";
        for (var i = 0; i < blankGroups.length; i++) {
            var group = blankGroups[i];
            report += "- " + group.artboard + " > " + group.name + "\n";
        }
        
        report += "\nSkipped Groups (" + skippedGroups.length + "):\n";
        for (var i = 0; i < skippedGroups.length; i++) {
            var group = skippedGroups[i];
            report += "- " + group.artboard + " > " + group.name + " (Marked as not ready)\n";
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
        
        alert("Export completed!\n\n" +
              "Modified: " + modifiedGroups.length + " groups\n" +
              "Unchanged: " + unchangedGroups.length + " groups\n" +
              "Blank: " + blankGroups.length + " groups\n" +
              "Skipped: " + skippedGroups.length + " groups\n\n" +
              "See export_report.txt for details");
    }
} else {
    alert("Please open a document first!");
} 