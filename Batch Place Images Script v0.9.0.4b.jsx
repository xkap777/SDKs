#target indesign

// --- Store selected image files for each group ---
var groupImageFiles = [];

// --- SHARED FUNCTIONS FOR BOTH AUTO AND MANUAL MODES ---

// Helper function to check if an edge is outer (not adjacent to another frame)
function isEdgeOuter(myBounds, allBounds, edge, tolerance) {
  for (var i = 0; i < allBounds.length; i++) {
    var other = allBounds[i];
    if (other === myBounds) continue;
    if (edge === "top" && Math.abs(myBounds[0] - other[2]) < tolerance && myBounds[1] < other[3] && myBounds[3] > other[1]) return false;
    if (edge === "bottom" && Math.abs(myBounds[2] - other[0]) < tolerance && myBounds[1] < other[3] && myBounds[3] > other[1]) return false;
    if (edge === "left" && Math.abs(myBounds[1] - other[3]) < tolerance && myBounds[0] < other[2] && myBounds[2] > other[0]) return false;
    if (edge === "right" && Math.abs(myBounds[3] - other[1]) < tolerance && myBounds[0] < other[2] && myBounds[2] > other[0]) return false;
  }
  return true;
}

// Draw line helper function (page-aware version)
function drawCropMarkLine(bounds, strokeWeight, registrationColor, noneSwatch, layer, targetFrame) {
  try {
    // Validate the targetFrame object
    if (!targetFrame || !targetFrame.isValid) {
      return; // Skip if frame is invalid
    }
    
    // Get the parent page of the frame to ensure marks are drawn on the correct page
    var parentPage = targetFrame.parentPage;
    if (!parentPage || !parentPage.isValid) {
      return; // Skip if parent page is invalid
    }
    
    var targetSpread = parentPage.parent;
    if (!targetSpread || !targetSpread.isValid) {
      return; // Skip if target spread is invalid
    }
    
    targetSpread.graphicLines.add(layer, undefined, undefined, {
      strokeWeight: strokeWeight,
      fillColor: noneSwatch,
      strokeColor: registrationColor,
      geometricBounds: bounds
    });
  } catch (e) {
    // Silently continue if crop mark drawing fails
  }
}

// Main crop marks drawing function (direct implementation from MyCropMarks.jsx)
// Single consolidated crop marks function (based on MyCropMarks.jsx)
function applyCropMarksToDocument(doc, cropMarkLength, cropMarkOffset, cropMarkWidth, bleedOffset) {
  try {
    var myOldRulerOrigin = doc.viewPreferences.rulerOrigin;
    doc.viewPreferences.rulerOrigin = RulerOrigin.spreadOrigin;
    var myOldXUnits = doc.viewPreferences.horizontalMeasurementUnits;
    var myOldYUnits = doc.viewPreferences.verticalMeasurementUnits;
    doc.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.points;
    doc.viewPreferences.verticalMeasurementUnits = MeasurementUnits.points;
    
    var myLayer = doc.layers.item("myCropMarks");
    try { 
      var myLayerName = myLayer.name; 
    }
    catch (myError) { 
      myLayer = doc.layers.add({ name: "myCropMarks" }); 
    }
    var myRegistrationColor = doc.colors.item("Registration");
    var myNoneSwatch = doc.swatches.item("None");

    // Process each page separately to avoid cross-page adjacency issues
    for (var pageIdx = 0; pageIdx < doc.pages.length; pageIdx++) {
      var page = doc.pages[pageIdx];
      var pageFrames = [];
      
      // Collect frames with images for this page only
      for (var rectIdx = 0; rectIdx < page.rectangles.length; rectIdx++) {
        var rect = page.rectangles[rectIdx];
        if (rect && rect.isValid && rect.images && rect.images.length > 0) {
          pageFrames.push(rect);
        }
      }
      
      if (pageFrames.length === 0) continue; // No frames on this page
      
      // Collect bounds for this page only
      var pageBounds = [];
      for (var i = 0; i < pageFrames.length; i++) {
        pageBounds.push(pageFrames[i].visibleBounds);
      }
      var tolerance = 0.5; // points

      // Apply crop marks to frames on this page
      for (var myCounter = 0; myCounter < pageFrames.length; myCounter++) {
        var myObject = pageFrames[myCounter];
        
        // Validate the frame object before processing
        if (!myObject || !myObject.isValid) {
          continue; // Skip invalid frames
        }
        
        var b = myObject.visibleBounds; // [y1, x1, y2, x2]
        var y1 = b[0], x1 = b[1], y2 = b[2], x2 = b[3];

        // Determine which edges are outer (only considering frames on this page)
        var topOuter = isEdgeOuter(b, pageBounds, "top", tolerance);
        var bottomOuter = isEdgeOuter(b, pageBounds, "bottom", tolerance);
        var leftOuter = isEdgeOuter(b, pageBounds, "left", tolerance);
        var rightOuter = isEdgeOuter(b, pageBounds, "right", tolerance);

        // Perpendicular "tick" marks
        if (topOuter) {
          drawCropMarkLine([y1 - cropMarkOffset, x1, y1 - (cropMarkOffset + cropMarkLength), x1], cropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer, myObject); // upper left
          drawCropMarkLine([y1 - cropMarkOffset, x2, y1 - (cropMarkOffset + cropMarkLength), x2], cropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer, myObject); // upper right
        }
        if (bottomOuter) {
          drawCropMarkLine([y2 + cropMarkOffset, x1, y2 + cropMarkOffset + cropMarkLength, x1], cropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer, myObject); // lower left
          drawCropMarkLine([y2 + cropMarkOffset, x2, y2 + cropMarkOffset + cropMarkLength, x2], cropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer, myObject); // lower right
        }
        if (leftOuter) {
          drawCropMarkLine([y1, x1 - cropMarkOffset, y1, x1 - (cropMarkOffset + cropMarkLength)], cropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer, myObject); // upper left
          drawCropMarkLine([y2, x1 - cropMarkOffset, y2, x1 - (cropMarkOffset + cropMarkLength)], cropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer, myObject); // lower left
        }
        if (rightOuter) {
          drawCropMarkLine([y1, x2 + cropMarkOffset, y1, x2 + cropMarkOffset + cropMarkLength], cropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer, myObject); // upper right
          drawCropMarkLine([y2, x2 + cropMarkOffset, y2, x2 + cropMarkOffset + cropMarkLength], cropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer, myObject); // lower right
        }

        // Parallel marks ONLY at outermost corners
        if (topOuter && leftOuter) {
          // Upper left horizontal (parallel to top)
          drawCropMarkLine([y1, x1 - cropMarkOffset, y1, x1 - (cropMarkOffset + cropMarkLength)], cropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer, myObject);
        }
        if (topOuter && rightOuter) {
          // Upper right horizontal
          drawCropMarkLine([y1, x2 + cropMarkOffset, y1, x2 + cropMarkOffset + cropMarkLength], cropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer, myObject);
        }
        if (bottomOuter && leftOuter) {
          // Lower left horizontal
          drawCropMarkLine([y2, x1 - cropMarkOffset, y2, x1 - (cropMarkOffset + cropMarkLength)], cropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer, myObject);
        }
        if (bottomOuter && rightOuter) {
          // Lower right horizontal
          drawCropMarkLine([y2, x2 + cropMarkOffset, y2, x2 + cropMarkOffset + cropMarkLength], cropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer, myObject);
        }

        // Draw bleed marks if bleed offset is specified and greater than 0
        if (bleedOffset && bleedOffset > 0) {
          // Perpendicular bleed marks (offset inward from crop marks by bleed amount)
          if (topOuter) {
            drawCropMarkLine([y1 - cropMarkOffset, x1 + bleedOffset, y1 - (cropMarkOffset + cropMarkLength), x1 + bleedOffset], cropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer, myObject); // upper left bleed
            drawCropMarkLine([y1 - cropMarkOffset, x2 - bleedOffset, y1 - (cropMarkOffset + cropMarkLength), x2 - bleedOffset], cropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer, myObject); // upper right bleed
          }
          if (bottomOuter) {
            drawCropMarkLine([y2 + cropMarkOffset, x1 + bleedOffset, y2 + cropMarkOffset + cropMarkLength, x1 + bleedOffset], cropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer, myObject); // lower left bleed
            drawCropMarkLine([y2 + cropMarkOffset, x2 - bleedOffset, y2 + cropMarkOffset + cropMarkLength, x2 - bleedOffset], cropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer, myObject); // lower right bleed
          }
          if (leftOuter) {
            drawCropMarkLine([y1 + bleedOffset, x1 - cropMarkOffset, y1 + bleedOffset, x1 - (cropMarkOffset + cropMarkLength)], cropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer, myObject); // upper left bleed
            drawCropMarkLine([y2 - bleedOffset, x1 - cropMarkOffset, y2 - bleedOffset, x1 - (cropMarkOffset + cropMarkLength)], cropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer, myObject); // lower left bleed
          }
          if (rightOuter) {
            drawCropMarkLine([y1 + bleedOffset, x2 + cropMarkOffset, y1 + bleedOffset, x2 + cropMarkOffset + cropMarkLength], cropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer, myObject); // upper right bleed
            drawCropMarkLine([y2 - bleedOffset, x2 + cropMarkOffset, y2 - bleedOffset, x2 + cropMarkOffset + cropMarkLength], cropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer, myObject); // lower right bleed
          }

          // Parallel bleed marks at outermost corners (offset inward by bleed amount)
          if (topOuter && leftOuter) {
            // Upper left horizontal bleed mark
            drawCropMarkLine([y1 + bleedOffset, x1 - cropMarkOffset, y1 + bleedOffset, x1 - (cropMarkOffset + cropMarkLength)], cropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer, myObject);
          }
          if (topOuter && rightOuter) {
            // Upper right horizontal bleed mark
            drawCropMarkLine([y1 + bleedOffset, x2 + cropMarkOffset, y1 + bleedOffset, x2 + cropMarkOffset + cropMarkLength], cropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer, myObject);
          }
          if (bottomOuter && leftOuter) {
            // Lower left horizontal bleed mark
            drawCropMarkLine([y2 - bleedOffset, x1 - cropMarkOffset, y2 - bleedOffset, x1 - (cropMarkOffset + cropMarkLength)], cropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer, myObject);
          }
          if (bottomOuter && rightOuter) {
            // Lower right horizontal bleed mark
            drawCropMarkLine([y2 - bleedOffset, x2 + cropMarkOffset, y2 - bleedOffset, x2 + cropMarkOffset + cropMarkLength], cropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer, myObject);
          }
        }
      }
    }
    
    doc.viewPreferences.rulerOrigin = myOldRulerOrigin;
    doc.viewPreferences.horizontalMeasurementUnits = myOldXUnits;
    doc.viewPreferences.verticalMeasurementUnits = myOldYUnits;
  } catch (e) {
    // Continue if crop mark application fails
  }
}

// Shared function to create frame with standard properties
function createImageFrame(page, bounds, doc) {
  return page.rectangles.add({
    geometricBounds: bounds,
    strokeWeight: 0,
    strokeColor: doc.swatches.itemByName("None"),
    fillColor: doc.swatches.itemByName("None")
  });
}

// Shared function to apply fitting and rotation to frame
function applyImageFitting(frame, placedImage, enableFitting, fittingOption, rotation) {
  // Apply fitting
  if (enableFitting) {
    if (fittingOption === "Content-Aware") {
      frame.fit(FitOptions.CONTENT_AWARE_FIT);
    } else if (fittingOption === "Proportional") {
      frame.fit(FitOptions.PROPORTIONALLY);
    }
  } else {
    frame.fit(FitOptions.CENTER_CONTENT);
  }
  
  // Apply rotation
  if (rotation && rotation !== 0) {
    placedImage.absoluteRotationAngle = rotation;
    frame.fit(FitOptions.CENTER_CONTENT);
  }
}

function placeImagesOnDocumentPages() {
  try {
    if (app.documents.length === 0) {
      alert("Please open a document before running this script.");
      return;
    }

    var doc = app.activeDocument;

    // --- UI DIALOG SETUP ---
    var dialog = new Window("dialog", "Set Frame Layout");
    dialog.orientation = "column";
    dialog.alignChildren = "fill";
    dialog.spacing = 10;

    // --- Tabbed panel for global auto controls ---
    var mainTabs = dialog.add('tabbedpanel');
    mainTabs.alignChildren = 'fill';
    mainTabs.preferredSize.width = 500;
    var globalTab = mainTabs.add('tab', undefined, 'Global Auto Features');
    globalTab.orientation = 'column';
    globalTab.alignChildren = 'left';
    globalTab.margins = 10;
    var globalAutoGroup = globalTab.add('group');
    globalAutoGroup.orientation = 'row';
    globalAutoGroup.alignChildren = 'left';
    globalAutoGroup.margins = 0;
    var autoPlaceCheckbox = globalAutoGroup.add('checkbox', undefined, 'Auto Place & Center Images');
    autoPlaceCheckbox.value = false;

    // --- Add global Keep Image Sizes Separate checkbox ---
    var globalKeepSeparateCheckbox = globalAutoGroup.add('checkbox', undefined, 'Keep Image Sizes Separate');
    globalKeepSeparateCheckbox.value = false;
    globalKeepSeparateCheckbox.enabled = true;

    // --- Auto Place checkbox logic ---
    autoPlaceCheckbox.onClick = function () {
      setManualFieldsEnabled(!this.value);
      // Keep separate checkbox is now always available, not tied to auto place
      updateAllGroupFeaturesDisplays(); // Update features display in all groups
    };
    
    // --- Keep Image Sizes Separate is always available ---
    
    // --- Initial state will be set after all functions are defined ---

    // --- Keep Image Sizes Separate checkbox logic ---
    globalKeepSeparateCheckbox.onClick = function () {
      updateAllGroupFeaturesDisplays(); // Update features display in all groups
    };

    var autoRotateBestFitCheckbox = globalAutoGroup.add('checkbox', undefined, 'Auto Rotate for Best Fit (Auto Place only)');
    autoRotateBestFitCheckbox.value = false;
    autoRotateBestFitCheckbox.enabled = true; // Always enabled

    // Frame Dimensions
    var frameGroup = dialog.add("panel", undefined, "Frame Dimensions");
    frameGroup.orientation = "column";
    frameGroup.alignChildren = "left";
    frameGroup.margins = 10;

    var frameRow1 = frameGroup.add("group");
    frameRow1.orientation = "row";
    frameRow1.add("statictext", undefined, "Frame Width (in):");
    var frameWidthInput = frameRow1.add("edittext", undefined, "0");
    frameWidthInput.characters = 5;
    frameRow1.add("statictext", undefined, "Frame Height (in):");
    var frameHeightInput = frameRow1.add("edittext", undefined, "0");
    frameHeightInput.characters = 5;

    var frameRow2 = frameGroup.add("group");
    frameRow2.orientation = "row";
    frameRow2.add("statictext", undefined, "Rows:");
    var rowsInput = frameRow2.add("edittext", undefined, "0");
    rowsInput.characters = 3;
    frameRow2.add("statictext", undefined, "Columns:");
    var columnsInput = frameRow2.add("edittext", undefined, "0");
    columnsInput.characters = 3;

    var frameRow3 = frameGroup.add("group");
    frameRow3.orientation = "row";
    frameRow3.add("statictext", undefined, "Horizontal Gutter (in):");
    var horizontalGutterInput = frameRow3.add("edittext", undefined, "0");
    horizontalGutterInput.characters = 4;
    frameRow3.add("statictext", undefined, "Vertical Gutter (in):");
    var verticalGutterInput = frameRow3.add("edittext", undefined, "0");
    verticalGutterInput.characters = 4;

    // Frame Margins
    var frameMarginGroup = dialog.add("panel", undefined, "Frame Margins (for positioning the frames)");
    frameMarginGroup.orientation = "column";
    frameMarginGroup.alignChildren = "left";
    frameMarginGroup.margins = 10;

    var linkFrameMarginsGroup = frameMarginGroup.add("group");
    linkFrameMarginsGroup.orientation = "row";
    var linkFrameMarginsCheckbox = linkFrameMarginsGroup.add("checkbox", undefined, "Link All");
    linkFrameMarginsCheckbox.value = true;

    var frameMarginRow1 = frameMarginGroup.add("group");
    frameMarginRow1.orientation = "row";
    frameMarginRow1.add("statictext", undefined, "Top (in):");
    var frameTopMarginInput = frameMarginRow1.add("edittext", undefined, "0");
    frameTopMarginInput.characters = 5;
    frameMarginRow1.add("statictext", undefined, "Bottom (in):");
    var frameBottomMarginInput = frameMarginRow1.add("edittext", undefined, "0");
    frameBottomMarginInput.characters = 5;

    var frameMarginRow2 = frameMarginGroup.add("group");
    frameMarginRow2.orientation = "row";
    frameMarginRow2.add("statictext", undefined, "Left (in):");
    var frameLeftMarginInput = frameMarginRow2.add("edittext", undefined, "0");
    frameLeftMarginInput.characters = 5;
    frameMarginRow2.add("statictext", undefined, "Right (in):");
    var frameRightMarginInput = frameMarginRow2.add("edittext", undefined, "0");
    frameRightMarginInput.characters = 5;

    // --- Link Frame Margins Logic ---
    function setAllFrameMargins(val) {
      frameTopMarginInput.text = val;
      frameBottomMarginInput.text = val;
      frameLeftMarginInput.text = val;
      frameRightMarginInput.text = val;
    }
    function syncFrameMargins(source) {
      if (linkFrameMarginsCheckbox.value) {
        setAllFrameMargins(source.text);
      }
    }
    linkFrameMarginsCheckbox.onClick = function () {
      if (linkFrameMarginsCheckbox.value) {
        setAllFrameMargins(frameTopMarginInput.text);
      }
    };
    frameTopMarginInput.onChanging = function () { syncFrameMargins(frameTopMarginInput); };
    frameBottomMarginInput.onChanging = function () { syncFrameMargins(frameBottomMarginInput); };
    frameLeftMarginInput.onChanging = function () { syncFrameMargins(frameLeftMarginInput); };
    frameRightMarginInput.onChanging = function () { syncFrameMargins(frameRightMarginInput); };

    // Page Margins
    var pageMarginGroup = dialog.add("panel", undefined, "Page Margins (for adjusting margins)");
    pageMarginGroup.orientation = "column";
    pageMarginGroup.alignChildren = "left";
    pageMarginGroup.margins = 10;

    var linkPageMarginsGroup = pageMarginGroup.add("group");
    linkPageMarginsGroup.orientation = "row";
    var linkPageMarginsCheckbox = linkPageMarginsGroup.add("checkbox", undefined, "Link All");
    linkPageMarginsCheckbox.value = true;

    var pageMarginRow1 = pageMarginGroup.add("group");
    pageMarginRow1.orientation = "row";
    pageMarginRow1.add("statictext", undefined, "Top (in):");
    var pageTopMarginInput = pageMarginRow1.add("edittext", undefined, "0");
    pageTopMarginInput.characters = 5;
    pageMarginRow1.add("statictext", undefined, "Bottom (in):");
    var pageBottomMarginInput = pageMarginRow1.add("edittext", undefined, "0");
    pageBottomMarginInput.characters = 5;

    var pageMarginRow2 = pageMarginGroup.add("group");
    pageMarginRow2.orientation = "row";
    pageMarginRow2.add("statictext", undefined, "Left (in):");
    var pageLeftMarginInput = pageMarginRow2.add("edittext", undefined, "0");
    pageLeftMarginInput.characters = 5;
    pageMarginRow2.add("statictext", undefined, "Right (in):");
    var pageRightMarginInput = pageMarginRow2.add("edittext", undefined, "0");
    pageRightMarginInput.characters = 5;

    // --- Link Page Margins Logic ---
    function setAllPageMargins(val) {
      pageTopMarginInput.text = val;
      pageBottomMarginInput.text = val;
      pageLeftMarginInput.text = val;
      pageRightMarginInput.text = val;
    }
    function syncPageMargins(source) {
      if (linkPageMarginsCheckbox.value) {
        setAllPageMargins(source.text);
      }
    }
    linkPageMarginsCheckbox.onClick = function () {
      if (linkPageMarginsCheckbox.value) {
        setAllPageMargins(pageTopMarginInput.text);
      }
    };
    pageTopMarginInput.onChanging = function () { syncPageMargins(pageTopMarginInput); };
    pageBottomMarginInput.onChanging = function () { syncPageMargins(pageBottomMarginInput); };
    pageLeftMarginInput.onChanging = function () { syncPageMargins(pageLeftMarginInput); };
    pageRightMarginInput.onChanging = function () { syncPageMargins(pageRightMarginInput); };

    // Fitting Options
    var fittingGroup = dialog.add("panel", undefined, "Fitting Options");
    fittingGroup.orientation = "column";
    fittingGroup.alignChildren = "left";
    fittingGroup.margins = 10;

    var enableFittingCheckbox = fittingGroup.add("checkbox", undefined, "Enable Fitting Options");
    enableFittingCheckbox.value = false;

    var fittingRow = fittingGroup.add("group");
    fittingRow.orientation = "row";
    fittingRow.add("statictext", undefined, "Fitting Option:");
    var fittingDropdown = fittingRow.add("dropdownlist", undefined, ["Content-Aware", "Proportional"]);
    fittingDropdown.selection = 0;
    fittingDropdown.enabled = false;

    enableFittingCheckbox.onClick = function () {
      fittingDropdown.enabled = enableFittingCheckbox.value;
    };

    // Rotation Options
    var rotationGroup = dialog.add("panel", undefined, "Rotation Options");
    rotationGroup.orientation = "row";
    rotationGroup.alignChildren = "left";
    rotationGroup.margins = 10;

    rotationGroup.add("statictext", undefined, "Image Rotation:");
    var rotationDropdown = rotationGroup.add("dropdownlist", undefined, ["0", "90", "180", "270"]);
    rotationDropdown.selection = 0;
    
    // Update all group features displays when rotation changes
    rotationDropdown.onChange = function() {
      for (var i = 0; i < stepImageGroups.length; i++) {
        if (stepImageGroups[i].updateFeaturesDisplay) {
          stepImageGroups[i].updateFeaturesDisplay();
        }
      }
    };

    // Step Images Panel (multi-group)
    var stepImagesPanel = dialog.add("panel", undefined, "Step Image Groups");
    stepImagesPanel.orientation = "column";
    stepImagesPanel.alignChildren = "left";
    stepImagesPanel.margins = 10;

    var enableStepGroupsCheckbox = stepImagesPanel.add("checkbox", undefined, "Enable Step Image Groups");
    enableStepGroupsCheckbox.value = false;

    // --- Add Keep Groups Separate checkbox below Enable Step Image Groups ---
    var keepGroupsSeparateCheckbox = stepImagesPanel.add("checkbox", undefined, "Keep Groups Separate");
    keepGroupsSeparateCheckbox.value = false;
    keepGroupsSeparateCheckbox.visible = false;

    // Container for dynamic group panels
    var groupPanelContainer = stepImagesPanel.add("group");
    groupPanelContainer.orientation = "column";
    groupPanelContainer.alignChildren = "fill"; // Changed from "left" to "fill"
    groupPanelContainer.alignment = "fill"; // Add alignment property
    groupPanelContainer.spacing = 5; // Add spacing
    groupPanelContainer.visible = false; // Hide initially

    var addGroupBtn = stepImagesPanel.add("button", undefined, "Add Group");
    addGroupBtn.visible = false;

    var stepImageGroups = [];
    var maxStepGroups = 9;
    var activeGroupIndex = -1; // Track which group is currently active/selected

    // Show/hide group controls dynamically
    enableStepGroupsCheckbox.onClick = function () {
      var enabled = enableStepGroupsCheckbox.value;
      groupPanelContainer.visible = enabled;
      addGroupBtn.visible = enabled;
      keepGroupsSeparateCheckbox.visible = enabled;
      // Hide all group panels if disabling
      if (!enabled) {
        for (var i = 0; i < stepImageGroups.length; i++) {
          stepImageGroups[i].groupPanel.visible = false;
          stepImageGroups[i].enableCheckbox.value = false;
          stepImageGroups[i].enableCheckbox.enabled = false;
          stepImageGroups[i].countInput.enabled = false;
        }
      }
      // Enable/disable global rotation dropdown based on auto rotate only (not group mode)
      rotationDropdown.enabled = !autoRotateBestFitCheckbox.value;
      dialog.layout.layout(true);
      dialog.layout.resize();
      updateGlobalAutoTabEnabled();
    };

    // --- Helper: Grey out global auto tab checkboxes if any group uses auto features ---
    function updateGlobalAutoTabEnabled() {
      // Keep global checkboxes always enabled so they can be used to configure new groups
      // The checkboxes should always be available for setting up each group's configuration
      
      // Ensure all global checkboxes remain enabled
      if (typeof globalAutoGroup !== 'undefined' && globalAutoGroup.children) {
        globalAutoGroup.children[0].enabled = true; // Auto Place & Center Images
        globalAutoGroup.children[1].enabled = true; // Auto Rotate for Best Fit
      }
      if (typeof autoPlaceCheckbox !== 'undefined') {
        autoPlaceCheckbox.enabled = true;
      }
      if (typeof autoRotateBestFitCheckbox !== 'undefined') {
        autoRotateBestFitCheckbox.enabled = true;
      }
      if (typeof globalKeepSeparateCheckbox !== 'undefined') {
        globalKeepSeparateCheckbox.enabled = true;
      }
    }

    // --- Helper: Update features display in all groups when global settings change ---
    function updateAllGroupFeaturesDisplays() {
      for (var i = 0; i < stepImageGroups.length; i++) {
        var group = stepImageGroups[i];
        if (group.updateFeaturesDisplay) {
          group.updateFeaturesDisplay();
        }
      }
    }

    // --- Active Group Management (without graphics API) ---
    function setActiveGroup(groupIndex) {
      // Set active group without changing visual appearance
      // Groups will show their configuration status through the [SET] indicator instead
      activeGroupIndex = groupIndex;
      loadGroupSettingsToGlobalFields(groupIndex);
    }
    
    function loadGroupSettingsToGlobalFields(groupIndex) {
      if (groupIndex < 0 || groupIndex >= stepImageGroups.length) return;
      
      var group = stepImageGroups[groupIndex];
      var settings = group.frameSettings;
      
      // Check if this is a brand new group with all default settings
      var isNewGroup = settings.frameWidth === 0 && 
                      settings.frameHeight === 0 && 
                      settings.rows === 0 && 
                      settings.columns === 0 && 
                      settings.horizontalGutter === 0 && 
                      settings.verticalGutter === 0 && 
                      !settings.appliedAutoPlace && 
                      !settings.appliedAutoRotate && 
                      !settings.appliedKeepSeparate &&
                      settings.appliedRotation === 0;
      
      // Don't load settings for new groups - keep the reset values
      if (isNewGroup) return;
      
      // Load saved settings into global fields
      frameWidthInput.text = settings.frameWidth || "0";
      frameHeightInput.text = settings.frameHeight || "0";
      rowsInput.text = settings.rows || "0";
      columnsInput.text = settings.columns || "0";
      horizontalGutterInput.text = settings.horizontalGutter || "0";
      verticalGutterInput.text = settings.verticalGutter || "0";
      
      // Load applied settings
      autoPlaceCheckbox.value = settings.appliedAutoPlace || false;
      autoRotateBestFitCheckbox.value = settings.appliedAutoRotate || false;
      globalKeepSeparateCheckbox.value = settings.appliedKeepSeparate || false;
      
      // Set rotation dropdown
      var rotation = settings.appliedRotation || 0;
      for (var i = 0; i < rotationDropdown.items.length; i++) {
        if (parseInt(rotationDropdown.items[i].text, 10) === rotation) {
          rotationDropdown.selection = i;
          break;
        }
      }
      
      // Update UI states based on loaded settings
      setManualFieldsEnabled(!settings.appliedAutoPlace);
      setAllRotationControlsEnabled(!settings.appliedAutoRotate);
    }

    // Add Group button logic
    addGroupBtn.onClick = function () {
      if (stepImageGroups.length >= maxStepGroups) return;
      
      // Auto-enable Step Groups feature when first group is added
      if (stepImageGroups.length === 0 && !enableStepGroupsCheckbox.value) {
        enableStepGroupsCheckbox.value = true;
        // Make container and related elements visible immediately
        addGroupBtn.visible = true;
        keepGroupsSeparateCheckbox.visible = true;
        // Keep rotation available even when groups are enabled (only auto rotate should disable it)
        rotationDropdown.enabled = !autoRotateBestFitCheckbox.value;
      }
      
      // Always ensure container is visible when adding groups
      groupPanelContainer.visible = true;
      
      // Force immediate layout update for container visibility
      stepImagesPanel.layout.layout(true);
      dialog.layout.layout(true);
      
      // Re-enable global fields so user can enter new values for this group
      // This ensures fields are available for configuring each new group
      setManualFieldsEnabled(true);
      setAllRotationControlsEnabled(!autoRotateBestFitCheckbox.value);
      
      var g = stepImageGroups.length;
      // Use a generic panel title, not 'Group N'
      var groupPanel = groupPanelContainer.add("panel", undefined, "Step Group");
      groupPanel.orientation = "column";
      groupPanel.alignChildren = "fill"; // Changed from "left" to "fill"
      groupPanel.alignment = "fill"; // Add alignment
      groupPanel.margins = 6;
      
      // Make group panel clickable to select it as active
      groupPanel.onClick = function() {
        setActiveGroup(g);
      };

      // --- Group header: label and frame settings summary, in separate left/right groups ---
      var groupHeaderRow = groupPanel.add("group");
      groupHeaderRow.orientation = "row";
      groupHeaderRow.alignment = ["fill", "top"];
      groupHeaderRow.alignChildren = ["fill", "center"];
      groupHeaderRow.margins = 0;
      groupHeaderRow.spacing = 0;

      var groupLabelGroup = groupHeaderRow.add("group");
      groupLabelGroup.orientation = "row";
      groupLabelGroup.alignment = ["left", "center"];
      var groupLabel = groupLabelGroup.add("statictext", undefined, "Group " + (g + 1));
      groupLabel.alignment = ["left", "center"];
      groupLabel.preferredSize.width = 70;
      
      // Add status indicator for manual settings configured
      var statusIndicator = groupLabelGroup.add("statictext", undefined, "");
      statusIndicator.alignment = ["left", "center"];
      statusIndicator.preferredSize.width = 25;
      // Make status indicator more visually distinct
      try {
        statusIndicator.graphics.foregroundColor = statusIndicator.graphics.newPen(statusIndicator.graphics.PenType.SOLID_COLOR, [0, 0.6, 0, 1], 1); // Green color
      } catch (e) {
        // Fallback if color setting fails
      }

      var summaryGroup = groupHeaderRow.add("group");
      summaryGroup.orientation = "row";
      summaryGroup.alignment = ["right", "center"];
      summaryGroup.alignChildren = ["right", "center"];
      summaryGroup.margins = 0;
      summaryGroup.spacing = 0;
      summaryGroup.minimumSize.width = 220;
      summaryGroup.maximumSize.width = 9999;
      var frameSettingsText = summaryGroup.add("statictext", undefined, "");
      frameSettingsText.alignment = ["right", "center"];
      frameSettingsText.preferredSize.width = 220;
      frameSettingsText.maximumSize.width = 9999;
      frameSettingsText.visible = true;
      frameSettingsText.enabled = true;

      // --- Controls for group settings ---
      var controlsRow = groupPanel.add("group");
      controlsRow.orientation = "row";
      controlsRow.alignChildren = "center";
      controlsRow.margins = 0;

      var enableGroupCheckbox = controlsRow.add("checkbox", undefined, "Enable");
      enableGroupCheckbox.value = true;

      controlsRow.add("statictext", undefined, "Times per Image:");
      var countInput = controlsRow.add("edittext", undefined, "1");
      countInput.characters = 3;

      // --- Add features display text to show what global settings will be applied ---
      var featuresGroup = controlsRow.add("group");
      featuresGroup.orientation = "column";
      featuresGroup.alignment = ["fill", "center"];
      var featuresText = featuresGroup.add("statictext", undefined, "Features: None applied");
      featuresText.alignment = ["left", "center"];
      featuresText.preferredSize.width = 200;
      
      // Function to update features display based on current global settings
      function updateFeaturesDisplay() {
        var features = [];
        if (autoPlaceCheckbox.value) features.push("Auto Place");
        if (autoRotateBestFitCheckbox.value) features.push("Auto Rotate");
        if (globalKeepSeparateCheckbox.value) features.push("Keep Separate");
        
        // Add current rotation
        var currentRotation = parseInt(rotationDropdown.selection.text, 10) || 0;
        if (currentRotation === 0) {
          features.push("port");
        } else if (currentRotation === 90) {
          features.push("landsc");
        } else if (currentRotation === 180) {
          features.push("port180");
        } else if (currentRotation === 270) {
          features.push("landsc270");
        }
        
        if (features.length > 0) {
          featuresText.text = "Features: " + features.join(", ");
        } else {
          featuresText.text = "Features: Manual layout";
        }
      }
      
      // Initial update
      updateFeaturesDisplay();

      // --- Enter/Remove buttons for manual settings ---
      var enterBtn = controlsRow.add("button", undefined, "Enter");
      var removeBtn = controlsRow.add("button", undefined, "Remove");

      // --- Enter/Remove logic ---
      function doEnter() {
        // Save to the currently active group instead of this specific group
        if (activeGroupIndex >= 0 && activeGroupIndex < stepImageGroups.length) {
          var activeGroup = stepImageGroups[activeGroupIndex];
          var activeSettings = activeGroup.frameSettings;
          
          // Capture manual settings to active group
          activeSettings.frameWidth = parseFloat(frameWidthInput.text) || 0;
          activeSettings.frameHeight = parseFloat(frameHeightInput.text) || 0;
          activeSettings.rows = parseInt(rowsInput.text, 10) || 0;
          activeSettings.columns = parseInt(columnsInput.text, 10) || 0;
          activeSettings.horizontalGutter = parseFloat(horizontalGutterInput.text) || 0;
          activeSettings.verticalGutter = parseFloat(verticalGutterInput.text) || 0;
          
          // Capture current global settings when Enter is pressed
          activeSettings.appliedAutoPlace = autoPlaceCheckbox.value;
          activeSettings.appliedAutoRotate = autoRotateBestFitCheckbox.value;
          activeSettings.appliedKeepSeparate = globalKeepSeparateCheckbox.value;
          activeSettings.appliedRotation = parseInt(rotationDropdown.selection.text, 10) || 0;
          
          // Update the active group's display
          activeGroup.groupFrameSettings = activeSettings; // Keep both references for compatibility
          updateFrameSettingsSummary();
          
          // Update status indicator for active group
          if (activeSettings.statusIndicator) {
            activeSettings.statusIndicator.text = "[SET]"; // Clear text indicator
            activeSettings.statusIndicator.helpTip = "Settings configured and applied";
            try {
              activeSettings.statusIndicator.graphics.foregroundColor = activeSettings.statusIndicator.graphics.newPen(activeSettings.statusIndicator.graphics.PenType.SOLID_COLOR, [0, 0.8, 0, 1], 1); // Bright green
            } catch (e) {
              // Fallback if color setting fails
            }
          }
          
          // Update features display to show what was applied to this group
          if (activeGroup.updateFeaturesDisplay) {
            activeGroup.updateFeaturesDisplay();
          }
          
          enterBtn.text = "Reset";
          enterBtn.onClick = doReset;
        }
      }
      function doReset() {
        // Reset settings for the currently active group
        if (activeGroupIndex >= 0 && activeGroupIndex < stepImageGroups.length) {
          var activeGroup = stepImageGroups[activeGroupIndex];
          var activeSettings = activeGroup.frameSettings;
          
          // Reset manual settings in active group storage
          activeSettings.frameWidth = 0;
          activeSettings.frameHeight = 0;
          activeSettings.rows = 0;
          activeSettings.columns = 0;
          activeSettings.horizontalGutter = 0;
          activeSettings.verticalGutter = 0;
          
          // Clear applied global settings
          activeSettings.appliedAutoPlace = false;
          activeSettings.appliedAutoRotate = false;
          activeSettings.appliedKeepSeparate = false;
          activeSettings.appliedRotation = 0;
          
          // Reset the global manual input fields to defaults
          frameWidthInput.text = "0";
          frameHeightInput.text = "0";
          rowsInput.text = "0";
          columnsInput.text = "0";
          horizontalGutterInput.text = "0";
          verticalGutterInput.text = "0";
          
          // Reset global checkboxes and settings (except group-related ones)
          autoPlaceCheckbox.value = false;
          autoRotateBestFitCheckbox.value = false;
          globalKeepSeparateCheckbox.value = false;
          rotationDropdown.selection = 0; // Reset to 0 degrees
          
          // Reset manual field related checkboxes
          linkFrameMarginsCheckbox.value = false;
          linkPageMarginsCheckbox.value = false;
          enableFittingCheckbox.value = false;
          fittingDropdown.selection = 0; // Reset to Content-Aware
          fittingDropdown.enabled = false; // Disable since enableFittingCheckbox is false
          
          // Update UI states after resetting global checkboxes
          setManualFieldsEnabled(true); // Enable manual fields since auto place is now false
          setAllRotationControlsEnabled(true); // Enable rotation since auto rotate is now false
          rotationGroup.enabled = true; // Ensure rotation group stays enabled
          
          // Clear status indicator for active group
          if (activeSettings.statusIndicator) {
            activeSettings.statusIndicator.text = "";
            activeSettings.statusIndicator.helpTip = "";
            try {
              activeSettings.statusIndicator.graphics.foregroundColor = activeSettings.statusIndicator.graphics.newPen(activeSettings.statusIndicator.graphics.PenType.SOLID_COLOR, [0, 0, 0, 1], 1); // Reset to black
            } catch (e) {
              // Fallback if color setting fails
            }
          }
          
          // Update active group's features display
          if (activeGroup.updateFeaturesDisplay) {
            activeGroup.updateFeaturesDisplay();
          }
          
          updateFrameSettingsSummary();
          
          enterBtn.text = "Enter";
          enterBtn.onClick = doEnter;
        }
      }

      // --- Remove group logic ---
      removeBtn.onClick = function () {
        groupPanel.parent.remove(groupPanel);
        for (var i = 0; i < stepImageGroups.length; i++) {
          if (stepImageGroups[i].groupPanel === groupPanel) {
            stepImageGroups.splice(i, 1);
            break;
          }
        }
        dialog.layout.layout(true);
        dialog.layout.resize();
        updateGlobalAutoTabEnabled();
      };

      // --- Prevent Enter button highlight after pressing Enter in manual fields ---
      var manualEditFields = [frameWidthInput, frameHeightInput, rowsInput, columnsInput, horizontalGutterInput, verticalGutterInput];
      for (var i = 0; i < manualEditFields.length; i++) {
        (function(field) {
          field.onEnterKey = function() {
            if (groupLabel && groupLabel.active !== undefined) groupLabel.active = true;
          };
        })(manualEditFields[i]);
      }

      // Store manual settings for this group (these are the group's own values, not the global/manual fields)
      var groupFrameSettings = {
        frameWidth: 0,
        frameHeight: 0,
        rows: 0,
        columns: 0,
        horizontalGutter: 0,
        verticalGutter: 0,
        // Store applied global settings when Enter is pressed
        appliedAutoPlace: false,
        appliedAutoRotate: false,
        appliedKeepSeparate: false,
        appliedRotation: 0,
        // Store UI element references
        statusIndicator: statusIndicator,
        groupPanel: groupPanel
      };
      function formatFrameSettings(fs) {
        var manualSettings = " | W: " + fs.frameWidth + " H: " + fs.frameHeight +
          " R: " + fs.rows + " C: " + fs.columns +
          " HG: " + fs.horizontalGutter + " VG: " + fs.verticalGutter;
        
        // Also show applied global settings if any
        var appliedFeatures = [];
        if (fs.appliedAutoPlace) appliedFeatures.push("AutoPlace");
        if (fs.appliedAutoRotate) appliedFeatures.push("AutoRotate");
        if (fs.appliedKeepSeparate) appliedFeatures.push("KeepSep");
        
        // Add rotation indicator for manual settings
        var rotationIndicator = "";
        if (fs.appliedRotation === 0) {
          rotationIndicator = "port"; // Portrait (0 degrees)
        } else if (fs.appliedRotation === 90) {
          rotationIndicator = "landsc"; // Landscape (90 degrees)
        } else if (fs.appliedRotation === 180) {
          rotationIndicator = "port180"; // Portrait upside down
        } else if (fs.appliedRotation === 270) {
          rotationIndicator = "landsc270"; // Landscape reversed
        }
        
        if (appliedFeatures.length > 0) {
          return "[" + appliedFeatures.join(", ") + ", " + rotationIndicator + "]" + manualSettings;
        } else {
          return "[Manual, " + rotationIndicator + "]" + manualSettings;
        }
      }
      // Always show summary, only update on Enter/Reset
      function updateFrameSettingsSummary() {
        frameSettingsText.text = formatFrameSettings(groupFrameSettings);
        try { frameSettingsText.justify = 'right'; } catch (e) {}
        summaryGroup.layout.layout(true);
        groupHeaderRow.layout.layout(true);
        if (groupPanel && groupPanel.layout) groupPanel.layout.layout(true);
      }
      updateFrameSettingsSummary(); // Show zeros on creation
      
      // Set up Enter button handler (defined in the functions below)
      enterBtn.onClick = doEnter;

      countInput.enabled = enableGroupCheckbox.value;
      enterBtn.enabled = enableGroupCheckbox.value; // Enable if group checkbox is checked

      // --- Manual fields enablement (no auto place/rotate per group) ---
      var manualEditFields = [frameWidthInput, frameHeightInput, rowsInput, columnsInput, horizontalGutterInput, verticalGutterInput];
      
      // Update manual fields based on enable checkbox
      function updateManualFieldsEnabled() {
        var manualEnabled = enableGroupCheckbox.value;
        for (var i = 0; i < manualEditFields.length; i++) {
          manualEditFields[i].enabled = manualEnabled;
        }
      }
      updateManualFieldsEnabled(); // Initial setup

      // --- Helper to keep global controls always enabled for configuring new groups ---
      function updateGlobalAutoPlaceDisable() {
        // Keep global checkboxes always enabled so they can be used to configure new groups
        // Each group stores its own applied settings, but global fields should remain available
        if (typeof autoPlaceCheckbox !== 'undefined') {
          autoPlaceCheckbox.enabled = true;
        }
        if (typeof autoRotateBestFitCheckbox !== 'undefined') {
          autoRotateBestFitCheckbox.enabled = true;
        }
        // Rotation controls should only be disabled by the auto rotate checkbox itself
        setAllRotationControlsEnabled(!autoRotateBestFitCheckbox.value);
      }

      stepImageGroups.push({
        panel: groupPanel, // Store panel reference for text updates
        groupPanel: groupPanel,
        enableCheckbox: enableGroupCheckbox,
        countInput: countInput,
        frameSettingsText: frameSettingsText,
        enterBtn: enterBtn,
        frameSettings: groupFrameSettings, // Use consistent naming
        groupFrameSettings: groupFrameSettings,
        updateFeaturesDisplay: updateFeaturesDisplay // Add reference to update function
      });
      
      // Reset global checkboxes to default state for new group configuration
      // We need to temporarily store the onClick handlers to prevent them from firing during reset
      var autoPlaceHandler = autoPlaceCheckbox.onClick;
      var autoRotateHandler = autoRotateBestFitCheckbox.onClick;
      var keepSeparateHandler = globalKeepSeparateCheckbox.onClick;
      var linkFrameHandler = linkFrameMarginsCheckbox.onClick;
      var linkPageHandler = linkPageMarginsCheckbox.onClick;
      var enableFittingHandler = enableFittingCheckbox.onClick;
      
      // Temporarily remove onClick handlers
      autoPlaceCheckbox.onClick = null;
      autoRotateBestFitCheckbox.onClick = null;
      globalKeepSeparateCheckbox.onClick = null;
      linkFrameMarginsCheckbox.onClick = null;
      linkPageMarginsCheckbox.onClick = null;
      enableFittingCheckbox.onClick = null;
      
      // Reset checkbox values
      autoPlaceCheckbox.value = false;
      autoRotateBestFitCheckbox.value = false;
      globalKeepSeparateCheckbox.value = false;
      linkFrameMarginsCheckbox.value = false;
      linkPageMarginsCheckbox.value = false;
      enableFittingCheckbox.value = false;
      
      // Reset rotation to default
      rotationDropdown.selection = 0; // Reset to 0 degrees
      
      // Reset fitting dropdown
      fittingDropdown.selection = 0; // Reset to Content-Aware
      fittingDropdown.enabled = false; // Disable since enableFittingCheckbox is false
      
      // Restore onClick handlers
      autoPlaceCheckbox.onClick = autoPlaceHandler;
      autoRotateBestFitCheckbox.onClick = autoRotateHandler;
      globalKeepSeparateCheckbox.onClick = keepSeparateHandler;
      linkFrameMarginsCheckbox.onClick = linkFrameHandler;
      linkPageMarginsCheckbox.onClick = linkPageHandler;
      enableFittingCheckbox.onClick = enableFittingHandler;
      
      // Update UI states after resetting checkboxes
      setManualFieldsEnabled(true); // Ensure manual fields are enabled
      setAllRotationControlsEnabled(true); // Enable rotation since auto rotate is now false
      
      // Ensure all global checkboxes remain enabled for configuring this new group
      autoPlaceCheckbox.enabled = true;
      autoRotateBestFitCheckbox.enabled = true;
      globalKeepSeparateCheckbox.enabled = true;
      linkFrameMarginsCheckbox.enabled = true;
      linkPageMarginsCheckbox.enabled = true;
      enableFittingCheckbox.enabled = true;
      
      // Make this new group active
      setActiveGroup(g);
      
      enableGroupCheckbox.onClick = function () {
        var enabled = this.value;
        countInput.enabled = enabled;
        enterBtn.enabled = enabled;
        
        if (enabled) {
          updateManualFieldsEnabled();
        } else {
          // Disable manual fields when group is disabled
          for (var i = 0; i < manualEditFields.length; i++) {
            manualEditFields[i].enabled = false;
          }
        }
        updateGlobalAutoTabEnabled();
      };
      groupPanel.visible = true;
      
      // Simple layout refresh (matching the working remove function approach)
      dialog.layout.layout(true);
      dialog.layout.resize();
      updateGlobalAutoTabEnabled();
    };

    // Crop Marks Panel
    var cropMarksPanel = dialog.add("panel", undefined, "Crop Marks");
    cropMarksPanel.orientation = "column";
    cropMarksPanel.alignChildren = "left";
    cropMarksPanel.margins = 10;

    var enableCropMarks = cropMarksPanel.add("checkbox", undefined, "Enable Crop Marks (in points)");
    enableCropMarks.value = false;

    var cropMarksOptionsGroup = cropMarksPanel.add("group");
    cropMarksOptionsGroup.orientation = "row";
    cropMarksOptionsGroup.enabled = false;

    cropMarksPanel.add("statictext", undefined, "Options:");

    var cropLengthLabel = cropMarksOptionsGroup.add("statictext", undefined, "Length:");
    var cropLengthInput = cropMarksOptionsGroup.add("edittext", undefined, "7");
    cropLengthInput.characters = 4;

    var cropOffsetLabel = cropMarksOptionsGroup.add("statictext", undefined, "Offset:");
    var cropOffsetInput = cropMarksOptionsGroup.add("edittext", undefined, "3");
    cropOffsetInput.characters = 4;

    var cropWeightLabel = cropMarksOptionsGroup.add("statictext", undefined, "Stroke Weight:");
    var cropWeightInput = cropMarksOptionsGroup.add("edittext", undefined, "0.25");
    cropWeightInput.characters = 4;

    var bleedOffsetLabel = cropMarksOptionsGroup.add("statictext", undefined, "Bleed Offset (9pt=0.125):");
    var bleedOffsetInput = cropMarksOptionsGroup.add("edittext", undefined, "0");
    bleedOffsetInput.characters = 4;

    enableCropMarks.onClick = function () {
      cropMarksOptionsGroup.enabled = enableCropMarks.value;
      cropMarksRangeGroup.enabled = enableCropMarks.value;
    };

    var cropMarksRangeGroup = cropMarksPanel.add("group");
    cropMarksRangeGroup.orientation = "row";
    cropMarksRangeGroup.enabled = false;
    cropMarksRangeGroup.add("statictext", undefined, "Draw Marks Around:");
    var cropMarksRangeEach = cropMarksRangeGroup.add("radiobutton", undefined, "Each Object");
    var cropMarksRangeAll = cropMarksRangeGroup.add("radiobutton", undefined, "Entire Selection");
    cropMarksRangeEach.value = true;

    function setManualFieldsEnabled(enabled) {
      frameGroup.enabled = enabled;
      frameMarginGroup.enabled = enabled;
      pageMarginGroup.enabled = enabled;
      fittingGroup.enabled = enabled;
      // Rotation group is controlled separately by auto rotate, not auto place
    }

    // --- Helper: Enable/disable all rotation controls ---
    function setAllRotationControlsEnabled(enabled) {
      rotationDropdown.enabled = enabled;
      // Groups no longer have portrait/landscape controls
    }

    // --- Auto Rotate disables all rotation controls ---
    autoRotateBestFitCheckbox.onClick = function () {
      setAllRotationControlsEnabled(!this.value);
      this.value = !!this.value;
      updateAllGroupFeaturesDisplays(); // Update features display in all groups
    };
    // Initial state - rotation controls should only be disabled by auto rotate
    setAllRotationControlsEnabled(!autoRotateBestFitCheckbox.value);
    rotationGroup.enabled = true; // Always enable rotation group (only dropdown is controlled by auto rotate)
    
    // Initial state - manual fields should be enabled/disabled based on auto place checkbox
    setManualFieldsEnabled(!autoPlaceCheckbox.value);

    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "center";
    buttonGroup.spacing = 20;

    var okButton = buttonGroup.add("button", undefined, "OK");
    var cancelButton = buttonGroup.add("button", undefined, "Cancel");

    cancelButton.onClick = function () {
      dialog.close(0);
    };

    okButton.onClick = function () {
      // --- Determine if any group is in auto place mode ---
      var anyGroupAutoPlace = false;
      for (var i = 0; i < stepImageGroups.length; i++) {
        if (stepImageGroups[i].enableCheckbox.value && stepImageGroups[i].groupFrameSettings.appliedAutoPlace) {
          anyGroupAutoPlace = true;
          break;
        }
      }
      var isGlobalAutoPlace = autoPlaceCheckbox.value;
      var isAnyAutoPlace = isGlobalAutoPlace || anyGroupAutoPlace;

      if (
        (!isAnyAutoPlace && (
          isNaN(frameWidthInput.text) ||
          isNaN(frameHeightInput.text) ||
          isNaN(rowsInput.text) ||
          isNaN(columnsInput.text) ||
          frameWidthInput.text <= 0 ||
          frameHeightInput.text <= 0 ||
          rowsInput.text <= 0 ||
          columnsInput.text <= 0
        )) ||
        isNaN(horizontalGutterInput.text) ||
        isNaN(verticalGutterInput.text) ||
        isNaN(frameTopMarginInput.text) ||
        isNaN(frameBottomMarginInput.text) ||
        isNaN(frameLeftMarginInput.text) ||
        isNaN(frameRightMarginInput.text) ||
        isNaN(pageTopMarginInput.text) ||
        isNaN(pageBottomMarginInput.text) ||
        isNaN(pageLeftMarginInput.text) ||
        isNaN(pageRightMarginInput.text) ||
        (enableCropMarks.value && (
          isNaN(cropLengthInput.text) ||
          isNaN(cropOffsetInput.text) ||
          isNaN(cropWeightInput.text) ||
          isNaN(bleedOffsetInput.text)
        ))
      ) {
        alert("Please enter valid numeric values.");
        return;
      }
      dialog.close(1);
    };

    if (dialog.show() !== 1) {
      alert("User cancelled the operation.");
      return;
    }

    // --- Retrieve Keep Groups Separate value ---
    var keepGroupsSeparate = keepGroupsSeparateCheckbox.value;
    // --- Groups now use global auto settings, no individual group auto settings ---

    // --- Retrieve user input ---
    var frameWidth = parseFloat(frameWidthInput.text);
    var frameHeight = parseFloat(frameHeightInput.text);
    var rows = parseInt(rowsInput.text, 10);
    var columns = parseInt(columnsInput.text, 10);
    var horizontalGutter = parseFloat(horizontalGutterInput.text);
    var verticalGutter = parseFloat(verticalGutterInput.text);

    var frameTopMargin = parseFloat(frameTopMarginInput.text);
    var frameBottomMargin = parseFloat(frameBottomMarginInput.text);
    var frameLeftMargin = parseFloat(frameLeftMarginInput.text);
    var frameRightMargin = parseFloat(frameRightMarginInput.text);

    var pageTopMargin = parseFloat(pageTopMarginInput.text);
    var pageBottomMargin = parseFloat(pageBottomMarginInput.text);
    var pageLeftMargin = parseFloat(pageLeftMarginInput.text);
    var pageRightMargin = parseFloat(pageRightMarginInput.text);

    var enableFitting = enableFittingCheckbox.value;
    var fittingOption = fittingDropdown.selection.text;
    var rotationAngle = parseInt(rotationDropdown.selection.text, 10);

    var doCropMarks = enableCropMarks.value;
    var cropMarkLength = parseFloat(cropLengthInput.text);
    var cropMarkOffset = parseFloat(cropOffsetInput.text);
    var cropMarkWidth = parseFloat(cropWeightInput.text);
    var cropMarkRange = cropMarksRangeEach.value ? 0 : 1;
    var bleedOffset = parseFloat(bleedOffsetInput.text);

    // --- Use crop mark values as entered (in points) ---
    var cropMarkLengthPt = cropMarkLength;
    var cropMarkOffsetPt = cropMarkOffset;
    var cropMarkWidthPt = cropMarkWidth;
    var bleedOffsetPt = bleedOffset;

    var autoPlaceAndCenter = autoPlaceCheckbox.value;
    var globalKeepImagesSeparate = globalKeepSeparateCheckbox.value;

    var stepGroups = [];
    var groupImageFilesFiltered = [];
    // --- Prompt for images per group after dialog ---
    if (enableStepGroupsCheckbox.value) {
      var anyGroupHasImages = false;
      var anyGroupEnabled = false;
      for (var g = 0; g < stepImageGroups.length; g++) {
        var group = stepImageGroups[g];
        if (group.enableCheckbox.value) {
          anyGroupEnabled = true;
          // Prompt for images for this group
          var files = File.openDialog("Select image files for Group " + (g+1), "*.jpg;*.png;*.tif", true);
          if (files && files.length > 0) {
            groupImageFiles[g] = files;
            stepGroups.push({
              count: parseInt(group.countInput.text, 10) || 1,
              rotation: group.groupFrameSettings.appliedAutoRotate ? 0 : group.groupFrameSettings.appliedRotation,
              keepSeparate: group.groupFrameSettings.appliedKeepSeparate,
              autoRotate: group.groupFrameSettings.appliedAutoRotate
            });
            groupImageFilesFiltered.push(files);
            anyGroupHasImages = true;
          } else {
            groupImageFiles[g] = [];
          }
        }
      }
      // If Step Groups are enabled but no groups are actually enabled, fall back to single image selection
      if (!anyGroupEnabled) {
        stepGroups.push({
          count: 1,
          rotation: parseInt(rotationDropdown.selection.text, 10) || 0,
          keepSeparate: globalKeepImagesSeparate
        });
        var files = File.openDialog("Select image files to place", "*.jpg;*.png;*.tif", true);
        if (!files || files.length === 0) {
          alert("No images selected. Operation canceled.");
          return;
        }
        groupImageFilesFiltered.push(files);
      } else if (!anyGroupHasImages) {
        alert("No images selected for any enabled group. Please select images for each enabled group.");
        return;
      }
    } else {
      stepGroups.push({
        count: 1,
        rotation: parseInt(rotationDropdown.selection.text, 10) || 0,
        keepSeparate: globalKeepImagesSeparate // Use global keep separate in non-group mode
      });
      // Use a single image file dialog for non-group mode
      var files = File.openDialog("Select image files to place", "*.jpg;*.png;*.tif", true);
      if (!files || files.length === 0) {
        alert("No images selected. Operation canceled.");
        return;
      }
      groupImageFilesFiltered.push(files);
    }

    // --- Placement plan for groups ---
    var placementPlan = [];
    if (enableStepGroupsCheckbox.value && keepGroupsSeparate) {
      // Each group will be placed on a separate page
      for (var g = 0; g < stepGroups.length; g++) {
        for (var i = 0; i < groupImageFilesFiltered[g].length; i++) {
          for (var c = 0; c < stepGroups[g].count; c++) {
            placementPlan.push({
              file: groupImageFilesFiltered[g][i],
              rotation: stepGroups[g].rotation,
              groupIdx: g,
              keepSeparate: true, // force separate for all
              forceNewPage: true  // custom flag for this mode
            });
          }
        }
      }
    } else {
      for (var g = 0; g < stepGroups.length; g++) {
        for (var i = 0; i < groupImageFilesFiltered[g].length; i++) {
          for (var c = 0; c < stepGroups[g].count; c++) {
            placementPlan.push({
              file: groupImageFilesFiltered[g][i],
              rotation: stepGroups[g].rotation,
              groupIdx: g,
              keepSeparate: stepGroups[g].keepSeparate
            });
          }
        }
      }
    }

    // --- SHARED HELPER FUNCTIONS ---
    function inchesToPoints(val) { return parseFloat(val) * 72; }
    
    function convertMarginsAndGuttersToPoints() {
      frameTopMargin = inchesToPoints(frameTopMargin);
      frameBottomMargin = inchesToPoints(frameBottomMargin);
      frameLeftMargin = inchesToPoints(frameLeftMargin);
      frameRightMargin = inchesToPoints(frameRightMargin);
      pageTopMargin = inchesToPoints(pageTopMargin);
      pageBottomMargin = inchesToPoints(pageBottomMargin);
      pageLeftMargin = inchesToPoints(pageLeftMargin);
      pageRightMargin = inchesToPoints(pageRightMargin);
      horizontalGutter = inchesToPoints(horizontalGutter);
      verticalGutter = inchesToPoints(verticalGutter);
    }
    
    function setPageMargins(page) {
      page.marginPreferences.top = pageTopMargin;
      page.marginPreferences.bottom = pageBottomMargin;
      page.marginPreferences.left = pageLeftMargin;
      page.marginPreferences.right = pageRightMargin;
    }
    
    function createFrame(page, x1, y1, x2, y2) {
      return page.rectangles.add({
        geometricBounds: [y1, x1, y2, x2],
        strokeWeight: 1,
        strokeColor: doc.swatches.itemByName("Black"),
        fillColor: doc.swatches.itemByName("None")
      });
    }
    
    function placeImageInFrame(frame, imageFile, rotation) {
      var placedImage = frame.place(imageFile)[0];
      
      // Apply fitting options
      if (enableFitting) {
        if (fittingOption === "Content-Aware") {
          frame.fit(FitOptions.CONTENT_AWARE_FIT);
        } else if (fittingOption === "Proportional") {
          frame.fit(FitOptions.PROPORTIONALLY);
        }
      } else {
        frame.fit(FitOptions.CENTER_CONTENT);
      }
      
      // Apply rotation if specified
      if (rotation && rotation !== 0) {
        placedImage.absoluteRotationAngle = rotation;
        frame.fit(FitOptions.CENTER_CONTENT);
      }
      
      return placedImage;
    }
    
    function removeEmptyFrames(frames, placedCount) {
      for (var j = placedCount; j < frames.length; j++) {
        try {
          frames[j].remove();
        } catch (e) {
          // Continue if frame removal fails
        }
      }
    }

    // --- AUTO PLACE & CENTER LOGIC (Global settings) ---
    if (autoPlaceAndCenter) {
      convertMarginsAndGuttersToPoints();

      // --- Special logic for global auto place + auto rotate + keep separate ---
      if (autoPlaceAndCenter && autoRotateBestFitCheckbox.value && globalKeepImagesSeparate) {
        // 1. Probe all image sizes and annotate placementPlan with best-fit (rotation, width, height)
        var sizeGroups = {};
        for (var i = 0; i < placementPlan.length; i++) {
          var probeFile = File(placementPlan[i].file);
          var imgW = 0, imgH = 0, bestRotation = 0, bestW = 0, bestH = 0;
          if (probeFile.exists) {
            var probeRect = doc.pages[0].rectangles.add({geometricBounds: [0,0,72,72]});
            var probePlaced = probeRect.place(probeFile)[0];
            var probeBounds = probePlaced.visibleBounds;
            imgW = probeBounds[3] - probeBounds[1];
            imgH = probeBounds[2] - probeBounds[0];
            probePlaced.remove();
            probeRect.remove();
            // Try both rotations for best fit (classic grid logic)
            var page = doc.pages[0];
            var pageBounds = page.bounds;
            var pageTop = pageBounds[0];
            var pageLeft = pageBounds[1];
            var pageWidth = pageBounds[3] - pageBounds[1];
            var pageHeight = pageBounds[2] - pageBounds[0];
            var usableWidth = pageWidth - pageLeftMargin - pageRightMargin;
            var usableHeight = pageHeight - pageTopMargin - pageBottomMargin;
            var cols0 = Math.max(1, Math.floor((usableWidth + horizontalGutter) / (imgW + horizontalGutter)));
            var rows0 = Math.max(1, Math.floor((usableHeight + verticalGutter) / (imgH + verticalGutter)));
            var count0 = cols0 * rows0;
            var cols90 = Math.max(1, Math.floor((usableWidth + horizontalGutter) / (imgH + horizontalGutter)));
            var rows90 = Math.max(1, Math.floor((usableHeight + verticalGutter) / (imgW + verticalGutter)));
            var count90 = cols90 * rows90;
            if (count90 > count0) {
              bestRotation = 90;
              bestW = imgH;
              bestH = imgW;
            } else {
              bestRotation = 0;
              bestW = imgW;
              bestH = imgH;
            }
          }
          placementPlan[i]._bestW = bestW;
          placementPlan[i]._bestH = bestH;
          placementPlan[i]._bestRotation = bestRotation;
          // Group key: width|height|rotation (rounded to 2 decimals for safety)
          var key = bestW.toFixed(2) + '|' + bestH.toFixed(2) + '|' + bestRotation;
          if (!sizeGroups[key]) sizeGroups[key] = [];
          sizeGroups[key].push(placementPlan[i]);
        }
        // 2. For each group, place all images of that size in grid(s) on consecutive pages
        var isFirstGroup = true;
        for (var key in sizeGroups) {
          var group = sizeGroups[key];
          if (!group.length) continue;
          // Use the first image in group to probe size again (for grid calc)
          var imgW = group[0]._bestW, imgH = group[0]._bestH, rotation = group[0]._bestRotation;
          var groupIdx = 0;
          var firstPageOfGroup = true;
          while (groupIdx < group.length) {
            // Always start a new page for each new size group (except very first group)
            var page;
            if (isFirstGroup && firstPageOfGroup) {
              page = doc.pages[0];
              isFirstGroup = false;
            } else {
              page = doc.pages.add();
            }
            firstPageOfGroup = false;
            setPageMargins(page);
            var pageBounds = page.bounds;
            var pageTop = pageBounds[0];
            var pageLeft = pageBounds[1];
            var pageWidth = pageBounds[3] - pageBounds[1];
            var pageHeight = pageBounds[2] - pageBounds[0];
            // Use margin box for usable area and centering (classic logic)
            var usableWidth = pageWidth - pageLeftMargin - pageRightMargin;
            var usableHeight = pageHeight - pageTopMargin - pageBottomMargin;
            // Calculate grid for this page (classic logic, always at least 1 frame)
            var cols = Math.max(1, Math.floor((usableWidth + horizontalGutter) / (imgW + horizontalGutter)));
            var rows = Math.max(1, Math.floor((usableHeight + verticalGutter) / (imgH + verticalGutter)));
            var count = cols * rows;
            if (cols < 1 || rows < 1 || count < 1) {
              cols = 1;
              rows = 1;
              count = 1;
            }
            var gridWidth = cols * imgW + (cols - 1) * horizontalGutter;
            var gridHeight = rows * imgH + (rows - 1) * verticalGutter;
            // For partial pages, always justify to top left (not centered)
            var startX = pageLeft + pageLeftMargin;
            var startY = pageTop + pageTopMargin;
            // Always create a full grid of frames (even if not all will be filled)
            var frames = [];
            for (var i = 0; i < count; i++) {
              var row = Math.floor(i / cols);
              var col = i % cols;
              var x1 = startX + col * (imgW + horizontalGutter);
              var y1 = startY + row * (imgH + verticalGutter);
              var x2 = x1 + imgW;
              var y2 = y1 + imgH;
              // Epsilon for tight fits (classic logic)
              if (x2 > pageLeft + pageWidth - pageRightMargin + 0.01 || y2 > pageTop + pageHeight - pageBottomMargin + 0.01) {
                continue;
              }
              // Use shared function to create frame with consistent properties
              var frame = createImageFrame(page, [y1, x1, y2, x2], doc);
              frames.push(frame);
            }
            // Place images in frames (only for this group, never carry over)
            var imageLayer = doc.layers.itemByName("Layer 1");
            if (!imageLayer.isValid) {
              imageLayer = doc.layers.add({ name: "Layer 1" });
            }
            for (var f = 0; f < frames.length; f++) {
              frames[f].itemLayer = imageLayer;
              frames[f].fillColor = doc.swatches.itemByName("None"); // Ensure no fill for all frames
            }
            var frameIdx = 0;
            // Place up to 'count' images from group[groupIdx] ... group[groupIdx+count-1]
            while (groupIdx < group.length && frameIdx < frames.length) {
              var plan = group[groupIdx];
              var imageFile = File(plan.file);
              var frame = frames[frameIdx];
              // Defensive checks before placing
              if (
                frame && frame.isValid &&
                imageFile.exists &&
                frame.geometricBounds &&
                (frame.geometricBounds[2] - frame.geometricBounds[0] > 0.1) &&
                (frame.geometricBounds[3] - frame.geometricBounds[1] > 0.1)
              ) {
                try {
                  var placedImage = frame.place(imageFile)[0];
                  
                  // Use shared function for fitting and rotation
                  applyImageFitting(frame, placedImage, enableFitting, fittingOption, rotation);
                  
                  frame.locked = false;
                } catch (e) {
                  // Remove the frame if placement fails
                  try { frame.remove(); } catch (e2) {}
                }
              }
              groupIdx++;
              frameIdx++;
            }
            // Do NOT remove unused frames before centering/crop marks!
            // All frames (including empty) are kept for centering and crop marks
            // Filter out any invalid or removed frames before centering/crop marks
            var validFrames = [];
            for (var i = 0; i < frames.length; i++) {
              var f = frames[i];
              if (
                f && f.isValid &&
                f.constructor.name === "Rectangle" &&
                !f.locked &&
                f.visible &&
                typeof f.geometricBounds !== "undefined"
              ) {
                validFrames.push(f);
              }
            }
            // Center group of frames (classic logic)
            if (validFrames.length > 0) {
              try {
                var groupObj = page.groups.add(validFrames);
                var groupBounds = groupObj.visibleBounds;
                var groupWidth = groupBounds[3] - groupBounds[1];
                var groupHeight = groupBounds[2] - groupBounds[0];
                var groupCenterX = groupBounds[1] + groupWidth / 2;
                var groupCenterY = groupBounds[0] + groupHeight / 2;
                var usableCenterX = pageLeft + pageLeftMargin + usableWidth / 2;
                var usableCenterY = pageTop + pageTopMargin + usableHeight / 2;
                var dx = usableCenterX - groupCenterX;
                var dy = usableCenterY - groupCenterY;
                groupObj.move([groupBounds[1] + dx, groupBounds[0] + dy]);
                groupObj.ungroup();
              } catch (e) {
                // Bypass invalid parameter error for grouping
              }
            }
            // --- Defensive: Always set fillColor, strokeColor, and contentType for all frames ---
            for (var i = 0; i < frames.length; i++) {
              var frame = frames[i];
              if (frame && frame.isValid) {
                try {
                  frame.fillColor = doc.swatches.itemByName("None");
                  frame.strokeColor = doc.swatches.itemByName("None");
                  if (frame.hasOwnProperty('contentType')) {
                    frame.contentType = ContentType.GRAPHIC;
                  }
                } catch (e) {}
              }
            }
            // --- Robust crop mark drawing: only outermost edges/corners, no marks between adjacent frames, PLUS bleed marks ---
            // NOTE: Crop marks for this mode (Keep Separate = TRUE) are disabled to prevent conflicts
            // Crop marks will be handled by the main crop marks section for consistency
            // if (doCropMarks && validFrames.length > 0) {
            //   // Use shared crop marks function
            //   drawCropMarks(doc, validFrames, cropMarkWidthPt, cropMarkOffsetPt, cropMarkLengthPt, bleedOffsetPt);
            // }
          }
        }
        
        // Apply crop marks if enabled
        if (doCropMarks) {
          applyCropMarksToDocument(doc, cropMarkLengthPt, cropMarkOffsetPt, cropMarkWidthPt, bleedOffsetPt);
        }
        
        alert("All images placed successfully!");
        // Continue to crop marks section below
      } else if (autoPlaceAndCenter && autoRotateBestFitCheckbox.value && !globalKeepImagesSeparate) {
        // Place all images continuously using position-based tracking for mixed sizes
        var currentPage = doc.pages[0];
        var currentPageFrames = [];
        var placedCount = 0;
        
        // Track page reference point
        var pageTopLeft = { x: 0, y: 0 }; // Will be set when page is initialized

        for (var i = 0; i < placementPlan.length; i++) {
          var plan = placementPlan[i];
          var probeFile = File(plan.file);
          var imgW = 0, imgH = 0, bestRotation = 0;

          // Probe image size and determine best rotation
          if (probeFile.exists) {
            var probeRect = currentPage.rectangles.add({geometricBounds: [0,0,72,72]});
            var probePlaced = probeRect.place(probeFile)[0];
            var probeBounds = probePlaced.visibleBounds;
            var originalW = probeBounds[3] - probeBounds[1];
            var originalH = probeBounds[2] - probeBounds[0];
            probePlaced.remove();
            probeRect.remove();

            // Calculate best fit rotation for this image
            var pageBounds = currentPage.bounds;
            var pageWidth = pageBounds[3] - pageBounds[1];
            var pageHeight = pageBounds[2] - pageBounds[0];
            var usableWidth = pageWidth - pageLeftMargin - pageRightMargin;
            var usableHeight = pageHeight - pageTopMargin - pageBottomMargin;

            // Try both rotations and pick the one that fits better
            var cols0 = Math.max(1, Math.floor((usableWidth + horizontalGutter) / (originalW + horizontalGutter)));
            var rows0 = Math.max(1, Math.floor((usableHeight + verticalGutter) / (originalH + verticalGutter)));
            var count0 = cols0 * rows0;

            var cols90 = Math.max(1, Math.floor((usableWidth + horizontalGutter) / (originalH + horizontalGutter)));
            var rows90 = Math.max(1, Math.floor((usableHeight + verticalGutter) / (originalW + verticalGutter)));
            var count90 = cols90 * rows90;

            if (count90 > count0) {
              bestRotation = 90;
              imgW = originalH;
              imgH = originalW;
            } else {
              bestRotation = 0;
              imgW = originalW;
              imgH = originalH;
            }
          }

          // Calculate grid for THIS image on current page (preserve original behavior)
          var pageBounds = currentPage.bounds;
          var pageTop = pageBounds[0];
          var pageLeft = pageBounds[1];
          var usableWidth = pageBounds[3] - pageBounds[1] - pageLeftMargin - pageRightMargin;
          var usableHeight = pageBounds[2] - pageBounds[0] - pageTopMargin - pageBottomMargin;
          
          var cols = Math.max(1, Math.floor((usableWidth + horizontalGutter) / (imgW + horizontalGutter)));
          var rows = Math.max(1, Math.floor((usableHeight + verticalGutter) / (imgH + verticalGutter)));
          var totalSlotsOnPage = cols * rows;
          
          // Set page top-left reference if not set
          if (pageTopLeft.x === 0 && pageTopLeft.y === 0) {
            pageTopLeft.x = pageLeft + pageLeftMargin;
            pageTopLeft.y = pageTop + pageTopMargin;
          }

          // Find the next available position for this image size
          var framePosition = null;
          var positionFound = false;
          
          // Try positions in left-to-right, top-to-bottom order
          for (var row = 0; row < rows && !positionFound; row++) {
            for (var col = 0; col < cols && !positionFound; col++) {
              var x1 = pageTopLeft.x + col * (imgW + horizontalGutter);
              var y1 = pageTopLeft.y + row * (imgH + verticalGutter);
              
              // Check if this position would fit within the usable area
              if (x1 + imgW <= pageLeft + pageLeftMargin + usableWidth && 
                  y1 + imgH <= pageTop + pageTopMargin + usableHeight) {
                
                // Check if this entire frame area is available (not just the top-left corner)
                var areaIsClear = true;
                
                // Check against all existing frames on this page
                for (var existingIdx = 0; existingIdx < currentPageFrames.length && areaIsClear; existingIdx++) {
                  var existingFrame = currentPageFrames[existingIdx];
                  if (existingFrame && existingFrame.isValid) {
                    var existingBounds = existingFrame.geometricBounds;
                    var existingTop = existingBounds[0];
                    var existingLeft = existingBounds[1];
                    var existingBottom = existingBounds[2];
                    var existingRight = existingBounds[3];
                    
                    // Check if the new frame would overlap with this existing frame
                    // Two rectangles overlap if they intersect in both X and Y dimensions
                    var newRight = x1 + imgW;
                    var newBottom = y1 + imgH;
                    
                    // Allow touching edges but not overlapping
                    var xOverlap = (x1 < existingRight) && (newRight > existingLeft);
                    var yOverlap = (y1 < existingBottom) && (newBottom > existingTop);
                    
                    if (xOverlap && yOverlap) {
                      areaIsClear = false;
                    }
                  }
                }
                
                if (areaIsClear) {
                  framePosition = { x: x1, y: y1, row: row, col: col };
                  positionFound = true;
                }
              }
            }
          }

          // If no position found on current page, start a new page
          if (!positionFound) {
            // Start new page
            currentPage = doc.pages.add();
            currentPage.marginPreferences.top = pageTopMargin;
            currentPage.marginPreferences.bottom = pageBottomMargin;
            currentPage.marginPreferences.left = pageLeftMargin;
            currentPage.marginPreferences.right = pageRightMargin;
            currentPageFrames = [];
            
            // Recalculate for new page
            pageBounds = currentPage.bounds;
            pageTop = pageBounds[0];
            pageLeft = pageBounds[1];
            pageTopLeft.x = pageLeft + pageLeftMargin;
            pageTopLeft.y = pageTop + pageTopMargin;
            
            // Use first position on new page
            framePosition = { x: pageTopLeft.x, y: pageTopLeft.y, row: 0, col: 0 };
          }

          // Create frame using the found position with shared function
          var frame = createImageFrame(currentPage, [framePosition.y, framePosition.x, framePosition.y + imgH, framePosition.x + imgW], doc);

          // Place the image
          try {
            var placedImage = frame.place(probeFile)[0];
            // Use shared function for fitting
            applyImageFitting(frame, placedImage, enableFitting, fittingOption, bestRotation);
            
            currentPageFrames.push(frame);
            placedCount++;
          } catch (e) {
            try { frame.remove(); } catch (e2) {}
          }
        }

        // --- Center all frames on each page after placement ---
        var pages = doc.pages;
        for (var pageIdx = 0; pageIdx < pages.length; pageIdx++) {
          var page = pages[pageIdx];
          var pageFrames = [];
          
          // Collect all frames on this page
          for (var rectIdx = 0; rectIdx < page.rectangles.length; rectIdx++) {
            var rect = page.rectangles[rectIdx];
            if (rect && rect.isValid && rect.images && rect.images.length > 0) {
              pageFrames.push(rect);
            }
          }
          
          // Center the group of frames on this page
          if (pageFrames.length > 0) {
            try {
              var pageBounds = page.bounds;
              var pageTop = pageBounds[0];
              var pageLeft = pageBounds[1];
              var pageWidth = pageBounds[3] - pageBounds[1];
              var pageHeight = pageBounds[2] - pageBounds[0];
              var usableWidth = pageWidth - pageLeftMargin - pageRightMargin;
              var usableHeight = pageHeight - pageTopMargin - pageBottomMargin;
              
              // Determine the grid parameters from the actual frames placed on this page
              // Find the most common frame size to establish the grid
              var frameSizes = {};
              var maxCount = 0;
              var mostCommonW = 0, mostCommonH = 0;
              
              for (var f = 0; f < pageFrames.length; f++) {
                var bounds = pageFrames[f].geometricBounds;
                var w = Math.round((bounds[3] - bounds[1]) * 100) / 100; // Round to avoid floating point issues
                var h = Math.round((bounds[2] - bounds[0]) * 100) / 100;
                var sizeKey = w + "x" + h;
                
                if (!frameSizes[sizeKey]) {
                  frameSizes[sizeKey] = { count: 0, w: w, h: h };
                }
                frameSizes[sizeKey].count++;
                
                if (frameSizes[sizeKey].count > maxCount) {
                  maxCount = frameSizes[sizeKey].count;
                  mostCommonW = w;
                  mostCommonH = h;
                }
              }
              
              // Use the most common frame size to determine grid
              var frameW = mostCommonW;
              var frameH = mostCommonH;
              
              // Calculate potential maximum grid for this frame size
              var maxCols = Math.floor((usableWidth + horizontalGutter) / (frameW + horizontalGutter));
              var maxRows = Math.floor((usableHeight + verticalGutter) / (frameH + verticalGutter));
              var maxFramesPerPage = maxCols * maxRows;
              
              // If this is a full page (80% or more of capacity), just center it
              if (pageFrames.length >= maxFramesPerPage * 0.8) {
                // Group all frames temporarily to get overall bounds
                var groupObj = page.groups.add(pageFrames);
                var groupBounds = groupObj.visibleBounds;
                var groupWidth = groupBounds[3] - groupBounds[1];
                var groupHeight = groupBounds[2] - groupBounds[0];
                var groupCenterX = groupBounds[1] + groupWidth / 2;
                var groupCenterY = groupBounds[0] + groupHeight / 2;
                
                // Calculate center of usable page area
                var usableCenterX = pageLeft + pageLeftMargin + usableWidth / 2;
                var usableCenterY = pageTop + pageTopMargin + usableHeight / 2;
                
                // Calculate offset needed to center the group
                var dx = usableCenterX - groupCenterX;
                var dy = usableCenterY - groupCenterY;
                
                // Move the group to center
                groupObj.move([groupBounds[1] + dx, groupBounds[0] + dy]);
                
                // Ungroup to return individual frames
                groupObj.ungroup();
              } else {
                // Partial page: Fill remaining space with empty frames, then center all frames
                
                // Find all occupied positions in the grid
                var occupiedPositions = {};
                
                for (var f = 0; f < pageFrames.length; f++) {
                  var bounds = pageFrames[f].geometricBounds;
                  var frameX = bounds[1];
                  var frameY = bounds[0];
                  
                  // Calculate grid position based on the established grid parameters
                  var col = Math.round((frameX - pageLeft - pageLeftMargin) / (frameW + horizontalGutter));
                  var row = Math.round((frameY - pageTop - pageTopMargin) / (frameH + verticalGutter));
                  
                  var key = row + "_" + col;
                  occupiedPositions[key] = true;
                }
                
                // Fill remaining spaces following left-to-right, top-to-bottom pattern
                var emptyFrames = [];
                
                for (var row = 0; row < maxRows; row++) {
                  for (var col = 0; col < maxCols; col++) {
                    var key = row + "_" + col;
                    
                    // If this position is not occupied, add an empty frame
                    if (!occupiedPositions[key]) {
                      var x = pageLeft + pageLeftMargin + col * (frameW + horizontalGutter);
                      var y = pageTop + pageTopMargin + row * (frameH + verticalGutter);
                      
                      // Check if this position would fit within the usable area
                      if (x + frameW <= pageLeft + pageLeftMargin + usableWidth && 
                          y + frameH <= pageTop + pageTopMargin + usableHeight) {
                        
                        // Use shared function to create empty frame
                        var emptyFrame = createImageFrame(page, [y, x, y + frameH, x + frameW], doc);
                        
                        emptyFrames.push(emptyFrame);
                      }
                    }
                  }
                }
                
                // Center the complete grid (images + empty frames) - do NOT remove empty frames
                var allFrames = pageFrames.concat(emptyFrames);
                if (allFrames.length > 0) {
                  var groupObj = page.groups.add(allFrames);
                  var groupBounds = groupObj.visibleBounds;
                  var groupWidth = groupBounds[3] - groupBounds[1];
                  var groupHeight = groupBounds[2] - groupBounds[0];
                  var groupCenterX = groupBounds[1] + groupWidth / 2;
                  var groupCenterY = groupBounds[0] + groupHeight / 2;
                  
                  var usableCenterX = pageLeft + pageLeftMargin + usableWidth / 2;
                  var usableCenterY = pageTop + pageTopMargin + usableHeight / 2;
                  
                  var dx = usableCenterX - groupCenterX;
                  var dy = usableCenterY - groupCenterY;
                  
                  groupObj.move([groupBounds[1] + dx, groupBounds[0] + dy]);
                  groupObj.ungroup();
                }
                
                // Keep empty frames visible - do NOT remove them yet
              }
            } catch (e) {
              // If grouping fails, skip centering for this page
            }
          }
        }

        // --- Now remove all empty frames (those without images) ---
        var pages = doc.pages;
        for (var pageIdx = 0; pageIdx < pages.length; pageIdx++) {
          var page = pages[pageIdx];
          var framesToRemove = [];
          
          // Collect empty frames for removal
          for (var rectIdx = 0; rectIdx < page.rectangles.length; rectIdx++) {
            var rect = page.rectangles[rectIdx];
            if (rect && rect.isValid && (!rect.images || rect.images.length === 0)) {
              framesToRemove.push(rect);
            }
          }
          
          // Remove empty frames
          for (var removeIdx = 0; removeIdx < framesToRemove.length; removeIdx++) {
            try {
              framesToRemove[removeIdx].remove();
            } catch (e) {
              // Continue if frame removal fails
            }
          }
        }

        // Apply crop marks if enabled
        if (doCropMarks) {
          applyCropMarksToDocument(doc, cropMarkLengthPt, cropMarkOffsetPt, cropMarkWidthPt, bleedOffsetPt);
        }

        alert("Placed " + placedCount + " images successfully!");
        // Continue to crop marks section below
      }
    } else { // End of if (autoPlaceAndCenter), start of manual placement
      // --- MANUAL PLACEMENT LOGIC (using values in inches as entered) ---
      // NO conversion to points - use values as is in inches
      
      var placedCount = 0;
      var page = doc.pages[0];
      
      // Set page margins (in inches)
      page.marginPreferences.top = pageTopMargin;
      page.marginPreferences.bottom = pageBottomMargin;
      page.marginPreferences.left = pageLeftMargin;
      page.marginPreferences.right = pageRightMargin;
      
      var pageBounds = page.bounds;
      var pageTop = pageBounds[0];
      var pageLeft = pageBounds[1];
      var pageWidth = pageBounds[3] - pageBounds[1];
      var pageHeight = pageBounds[2] - pageBounds[0];
      
      // Use frame margins for positioning (in inches)
      var startX = pageLeft + frameLeftMargin;
      var startY = pageTop + frameTopMargin;
      
      // Place images using same logic as auto mode
      var imageIndex = 0;
      var currentPage = page;
      
      while (imageIndex < placementPlan.length) {
        // Create frames for current page
        var pageFrames = [];
        var count = rows * columns;
        
        for (var i = 0; i < count; i++) {
          var row = Math.floor(i / columns);
          var col = i % columns;
          var x1 = startX + col * (frameWidth + horizontalGutter);
          var y1 = startY + row * (frameHeight + verticalGutter);
          var x2 = x1 + frameWidth;
          var y2 = y1 + frameHeight;
          
          // Check bounds (in inches)
          if (x2 > pageLeft + pageWidth - frameRightMargin + 0.01 || 
              y2 > pageTop + pageHeight - frameBottomMargin + 0.01) {
            continue;
          }
          
          // Use shared function to create frame
          var frame = createImageFrame(currentPage, [y1, x1, y2, x2], doc);
          pageFrames.push(frame);
        }
        
        // Place images in current page frames
        var framesUsed = 0;
        for (var j = 0; j < pageFrames.length && imageIndex < placementPlan.length; j++) {
          try {
            var imageFile = File(placementPlan[imageIndex].file);
            if (imageFile.exists) {
              var placedImage = pageFrames[j].place(imageFile)[0];
              
              // Use shared function for fitting and rotation
              applyImageFitting(pageFrames[j], placedImage, enableFitting, fittingOption, placementPlan[imageIndex].rotation);
              
              placedCount++;
              framesUsed++;
            }
            imageIndex++;
          } catch (e) {
            imageIndex++;
            // Continue on error (same as auto mode)
          }
        }
        
        // Remove empty frames on current page
        for (var k = framesUsed; k < pageFrames.length; k++) {
          try {
            pageFrames[k].remove();
          } catch (e) {
            // Continue on error
          }
        }
        
        // If there are more images to place, create a new page
        if (imageIndex < placementPlan.length) {
          currentPage = doc.pages.add();
          
          // Set page margins for new page
          currentPage.marginPreferences.top = pageTopMargin;
          currentPage.marginPreferences.bottom = pageBottomMargin;
          currentPage.marginPreferences.left = pageLeftMargin;
          currentPage.marginPreferences.right = pageRightMargin;
          
          // Recalculate page bounds for new page
          var pageBounds = currentPage.bounds;
          pageTop = pageBounds[0];
          pageLeft = pageBounds[1];
          pageWidth = pageBounds[3] - pageBounds[1];
          pageHeight = pageBounds[2] - pageBounds[0];
          
          // Recalculate start position for new page
          startX = pageLeft + frameLeftMargin;
          startY = pageTop + frameTopMargin;
        }
      }
      
      // Apply crop marks if enabled
      if (doCropMarks) {
        applyCropMarksToDocument(doc, cropMarkLengthPt, cropMarkOffsetPt, cropMarkWidthPt, bleedOffsetPt);
      }
      
      alert("Placed " + placedCount + " images in manual layout!");
      // Continue to crop marks section below
    } // End of else block (manual placement)
  } catch (e) {
    alert("An error occurred while placing images: " + e.message);
    // Error displayed for debugging
  }
}
// Run the script
placeImagesOnDocumentPages();