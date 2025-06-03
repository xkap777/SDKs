#target indesign

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

    // --- Add global Keep Image Sizes Separate checkbox (only enabled when Auto Place is checked) ---
    var globalKeepSeparateCheckbox = globalAutoGroup.add('checkbox', undefined, 'Keep Image Sizes Separate');
    globalKeepSeparateCheckbox.value = false;
    globalKeepSeparateCheckbox.enabled = false;

    // --- Only enable globalKeepSeparateCheckbox when Auto Place is checked ---
    autoPlaceCheckbox.onClick = function () {
      setManualFieldsEnabled(!this.value);
      globalKeepSeparateCheckbox.enabled = !!this.value;
      if (!this.value) globalKeepSeparateCheckbox.value = false;
    };
    // --- Disable globalKeepSeparateCheckbox if Auto Place is unchecked at dialog start ---
    globalKeepSeparateCheckbox.enabled = !!autoPlaceCheckbox.value;

    // --- When globalKeepSeparateCheckbox is enabled, allow user to toggle it, otherwise always set to false ---
    globalKeepSeparateCheckbox.onClick = function () {
      if (!globalKeepSeparateCheckbox.enabled) {
        globalKeepSeparateCheckbox.value = false;
      }
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

    // Step Images Panel (multi-group)
    var stepImagesPanel = dialog.add("panel", undefined, "Step Image Groups");
    stepImagesPanel.orientation = "column";
    stepImagesPanel.alignChildren = "left";
    stepImagesPanel.margins = 10;

    var enableStepGroupsCheckbox = stepImagesPanel.add("checkbox", undefined, "Enable Step Image Groups");
    enableStepGroupsCheckbox.value = false;

    // Container for dynamic group panels
    var groupPanelContainer = stepImagesPanel.add("group");
    groupPanelContainer.orientation = "column";
    groupPanelContainer.alignChildren = "left";
    groupPanelContainer.visible = false; // Hide initially

    var addGroupBtn = stepImagesPanel.add("button", undefined, "Add Group");
    addGroupBtn.visible = false;

    var stepImageGroups = [];
    var maxStepGroups = 9;

    // Show/hide group controls dynamically
    enableStepGroupsCheckbox.onClick = function () {
      groupPanelContainer.visible = this.value;
      addGroupBtn.visible = this.value;
      // Hide all group panels if disabling
      if (!this.value) {
        for (var i = 0; i < stepImageGroups.length; i++) {
          stepImageGroups[i].groupPanel.visible = false;
          stepImageGroups[i].enableCheckbox.value = false;
          stepImageGroups[i].enableCheckbox.enabled = false;
          stepImageGroups[i].countInput.enabled = false;
          stepImageGroups[i].portraitRadio.enabled = false;
          stepImageGroups[i].landscapeRadio.enabled = false;
          stepImageGroups[i].keepSeparateCheckbox.enabled = false;
        }
      }
      // Enable/disable global rotation dropdown based on group mode
      rotationDropdown.enabled = !this.value;
      dialog.layout.layout(true);
      dialog.layout.resize();
      updateGlobalAutoTabEnabled();
    };

    // --- Helper: Grey out global auto tab checkboxes if any group uses auto features ---
    function updateGlobalAutoTabEnabled() {
      var anyGroupAutoPlace = false;
      var anyGroupAutoRotate = false;
      for (var i = 0; i < stepImageGroups.length; i++) {
        var group = stepImageGroups[i];
        if (group.enableCheckbox.value && group.autoPlaceCheckbox.value) {
          anyGroupAutoPlace = true;
        }
        if (group.enableCheckbox.value && group.autoRotateCheckbox.value) {
          anyGroupAutoRotate = true;
        }
      }
      // Grey out each global control independently
      globalAutoGroup.children[0].enabled = !anyGroupAutoPlace; // Auto Place & Center Images
      globalAutoGroup.children[1].enabled = !anyGroupAutoRotate; // Auto Rotate for Best Fit
      // --- NEW: Disable manual fields if any group auto place is active ---
      setManualFieldsEnabled(!anyGroupAutoPlace && !autoPlaceCheckbox.value);
    }

    // Add Group button logic
    addGroupBtn.onClick = function () {
      if (stepImageGroups.length >= maxStepGroups) return;
      var g = stepImageGroups.length;
      // Use a generic panel title, not 'Group N'
      var groupPanel = groupPanelContainer.add("panel", undefined, "Step Group");
      groupPanel.orientation = "column";
      groupPanel.alignChildren = "left";
      groupPanel.margins = 6;

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
      enableGroupCheckbox.value = false;

      controlsRow.add("statictext", undefined, "Times per Image:");
      var countInput = controlsRow.add("edittext", undefined, "1");
      countInput.characters = 3;

      controlsRow.add("statictext", undefined, "Orientation:");
      var portraitRadio = controlsRow.add("radiobutton", undefined, "Portrait");
      var landscapeRadio = controlsRow.add("radiobutton", undefined, "Landscape");
      portraitRadio.value = true;
      landscapeRadio.value = false;

      var keepSeparateCheckbox = controlsRow.add("checkbox", undefined, "Keep Image Sizes Separate");
      keepSeparateCheckbox.value = false;

      // --- Add checkboxes for Auto Place and Auto Rotate (independent) ---
      var autoModeGroup = controlsRow.add("group");
      autoModeGroup.orientation = "row";
      var autoPlaceCheckbox = autoModeGroup.add("checkbox", undefined, "Auto Place");
      var autoRotateCheckbox = autoModeGroup.add("checkbox", undefined, "Auto Rotate");
      autoPlaceCheckbox.value = false;
      autoRotateCheckbox.value = false;

      // --- Enter/Reset/Remove buttons for manual settings ---
      var enterBtn = controlsRow.add("button", undefined, "Enter");
      var resetBtn = controlsRow.add("button", undefined, "Reset");
      var removeBtn = controlsRow.add("button", undefined, "Remove");

      // --- Enter/Reset toggle logic ---
      function doEnter() {
        groupFrameSettings.frameWidth = parseFloat(frameWidthInput.text) || 0;
        groupFrameSettings.frameHeight = parseFloat(frameHeightInput.text) || 0;
        groupFrameSettings.rows = parseInt(rowsInput.text, 10) || 0;
        groupFrameSettings.columns = parseInt(columnsInput.text, 10) || 0;
        groupFrameSettings.horizontalGutter = parseFloat(horizontalGutterInput.text) || 0;
        groupFrameSettings.verticalGutter = parseFloat(verticalGutterInput.text) || 0;
        updateFrameSettingsSummary();
        if (groupLabel && groupLabel.active !== undefined) groupLabel.active = true;
        enterBtn.text = "Reset";
        enterBtn.onClick = doReset;
      }
      function doReset() {
        groupFrameSettings.frameWidth = 0;
        groupFrameSettings.frameHeight = 0;
        groupFrameSettings.rows = 0;
        groupFrameSettings.columns = 0;
        groupFrameSettings.horizontalGutter = 0;
        groupFrameSettings.verticalGutter = 0;
        updateFrameSettingsSummary();
        keepSeparateCheckbox.value = false;
        autoPlaceCheckbox.value = false;
        autoRotateCheckbox.value = false;
        portraitRadio.value = true;
        landscapeRadio.value = false;
        if (groupLabel && groupLabel.active !== undefined) groupLabel.active = true;
        enterBtn.text = "Enter";
        enterBtn.onClick = doEnter;
      }
      enterBtn.onClick = doEnter;
      resetBtn.onClick = doReset;

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
        verticalGutter: 0
      };
      function formatFrameSettings(fs) {
        return " | W: " + fs.frameWidth + " H: " + fs.frameHeight +
          " R: " + fs.rows + " C: " + fs.columns +
          " HG: " + fs.horizontalGutter + " VG: " + fs.verticalGutter;
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
      // Only update group values and summary on Enter
      enterBtn.onClick = function () {
        groupFrameSettings.frameWidth = parseFloat(frameWidthInput.text) || 0;
        groupFrameSettings.frameHeight = parseFloat(frameHeightInput.text) || 0;
        groupFrameSettings.rows = parseInt(rowsInput.text, 10) || 0;
        groupFrameSettings.columns = parseInt(columnsInput.text, 10) || 0;
        groupFrameSettings.horizontalGutter = parseFloat(horizontalGutterInput.text) || 0;
        groupFrameSettings.verticalGutter = parseFloat(verticalGutterInput.text) || 0;
        updateFrameSettingsSummary();
        // Remove focus from button to prevent highlight by setting focus to group label
        if (groupLabel && groupLabel.active !== undefined) groupLabel.active = true;
      };
      // Only reset group values and summary on Reset
      resetBtn.onClick = function () {
        groupFrameSettings.frameWidth = 0;
        groupFrameSettings.frameHeight = 0;
        groupFrameSettings.rows = 0;
        groupFrameSettings.columns = 0;
        groupFrameSettings.horizontalGutter = 0;
        groupFrameSettings.verticalGutter = 0;
        updateFrameSettingsSummary();
        // Reset all checkboxes and radios in this group to default, except Enable
        // enableGroupCheckbox.value = false; // Do NOT reset Enable
        keepSeparateCheckbox.value = false;
        autoPlaceCheckbox.value = false;
        autoRotateCheckbox.value = false;
        portraitRadio.value = true;
        landscapeRadio.value = false;
        // Remove focus from button to prevent highlight by setting focus to group label
        if (groupLabel && groupLabel.active !== undefined) groupLabel.active = true;
      };

      countInput.enabled = false;
      portraitRadio.enabled = false;
      landscapeRadio.enabled = false;
      keepSeparateCheckbox.enabled = false;
      enterBtn.enabled = false;
      resetBtn.enabled = false;
      autoPlaceCheckbox.enabled = false;
      autoRotateCheckbox.enabled = false;

      // --- Checkbox logic for auto place/auto rotate ---
      function updateAutoModeUI() {
        // If auto rotate is checked, disable portrait/landscape
        if (autoRotateCheckbox.value) {
          portraitRadio.enabled = false;
          landscapeRadio.enabled = false;
        } else {
          portraitRadio.enabled = enableGroupCheckbox.value && !autoRotateBestFitCheckbox.value;
          landscapeRadio.enabled = enableGroupCheckbox.value && !autoRotateBestFitCheckbox.value;
        }
        // Always enable group auto checkboxes and Enter/Reset if group is enabled
        autoPlaceCheckbox.enabled = enableGroupCheckbox.value;
        autoRotateCheckbox.enabled = enableGroupCheckbox.value;
        enterBtn.enabled = enableGroupCheckbox.value; // Always enabled if group is enabled
        resetBtn.enabled = enableGroupCheckbox.value; // Always enabled if group is enabled
      }
      autoPlaceCheckbox.onClick = function () {
        updateAutoModeUI();
        updateGlobalAutoTabEnabled();
      };
      autoRotateCheckbox.onClick = function () {
        updateAutoModeUI();
        updateGlobalAutoTabEnabled();
      };

      enableGroupCheckbox.onClick = function () {
        var enabled = this.value;
        var rotationEnabled = enabled && !autoRotateBestFitCheckbox.value;
        countInput.enabled = enabled;
        keepSeparateCheckbox.enabled = enabled;
        // Always enable group auto checkboxes and Enter/Reset if group is enabled
        autoPlaceCheckbox.enabled = enabled;
        autoRotateCheckbox.enabled = enabled;
        enterBtn.enabled = enabled; // Always enabled if group is enabled
        resetBtn.enabled = enabled; // Always enabled if group is enabled
        // If disabling, also clear checkboxes and re-enable global
        if (!enabled) {
          autoPlaceCheckbox.value = false;
          autoRotateCheckbox.value = false;
          portraitRadio.enabled = false;
          landscapeRadio.enabled = false;
        } else {
          updateAutoModeUI();
        }
        updateGlobalAutoTabEnabled();
      };

      // --- Helper to update global auto place controls based on group settings ---
      function updateGlobalAutoPlaceDisable() {
        var anyGroupAutoPlace = false;
        for (var i = 0; i < stepImageGroups.length; i++) {
          if (stepImageGroups[i].enableCheckbox.value && stepImageGroups[i].autoPlaceCheckbox.value) {
            anyGroupAutoPlace = true;
            break;
          }
        }
        // Disable global auto place and manual fields if any group is auto place
        autoPlaceCheckbox.enabled = !anyGroupAutoPlace;
        setManualFieldsEnabled(!autoPlaceCheckbox.value && !anyGroupAutoPlace);
        setAllRotationControlsEnabled(!autoRotateBestFitCheckbox.value && !anyGroupAutoPlace);
      }

      stepImageGroups.push({
        groupPanel: groupPanel,
        enableCheckbox: enableGroupCheckbox,
        countInput: countInput,
        portraitRadio: portraitRadio,
        landscapeRadio: landscapeRadio,
        keepSeparateCheckbox: keepSeparateCheckbox,
        frameSettingsText: frameSettingsText,
        enterBtn: enterBtn,
        resetBtn: resetBtn,
        groupFrameSettings: groupFrameSettings,
        autoPlaceCheckbox: autoPlaceCheckbox,
        autoRotateCheckbox: autoRotateCheckbox
      });
      // Attach updateGlobalAutoTabEnabled to all relevant group controls:
      autoPlaceCheckbox.onClick = function () {
        updateAutoModeUI();
        updateGlobalAutoTabEnabled();
      };
      autoRotateCheckbox.onClick = function () {
        updateAutoModeUI();
        updateGlobalAutoTabEnabled();
      };
      enableGroupCheckbox.onClick = function () {
        var enabled = this.value;
        var rotationEnabled = enabled && !autoRotateBestFitCheckbox.value;
        countInput.enabled = enabled;
        keepSeparateCheckbox.enabled = enabled;
        var manualEnabled = enabled && !autoPlaceCheckbox.value;
        enterBtn.enabled = manualEnabled;
        resetBtn.enabled = manualEnabled;
        autoPlaceCheckbox.enabled = enabled;
        autoRotateCheckbox.enabled = enabled;
        if (!enabled) {
          autoPlaceCheckbox.value = false;
          autoRotateCheckbox.value = false;
          portraitRadio.enabled = false;
          landscapeRadio.enabled = false;
        } else {
          updateAutoModeUI();
        }
        updateGlobalAutoTabEnabled();
      };
      groupPanel.visible = true;
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
      rotationGroup.enabled = enabled;
    }

    // --- Helper: Enable/disable all rotation controls ---
    function setAllRotationControlsEnabled(enabled) {
      rotationDropdown.enabled = enabled;
      for (var g = 0; g < stepImageGroups.length; g++) {
        // Only enable portrait/landscape if group is enabled and not auto rotate
        var group = stepImageGroups[g];
        var groupEnabled = group.enableCheckbox.value;
        var groupAutoRotate = group.autoRotateCheckbox && group.autoRotateCheckbox.value;
        group.portraitRadio.enabled = enabled && groupEnabled && !groupAutoRotate;
        group.landscapeRadio.enabled = enabled && groupEnabled && !groupAutoRotate;
      }
    }

    // --- Auto Rotate disables all rotation controls ---
    autoRotateBestFitCheckbox.onClick = function () {
      setAllRotationControlsEnabled(!this.value);
      this.value = !!this.value;
    };
    // Initial state
    setAllRotationControlsEnabled(!autoRotateBestFitCheckbox.value);

    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "center";
    buttonGroup.spacing = 20;

    var okButton = buttonGroup.add("button", undefined, "OK");
    var cancelButton = buttonGroup.add("button", undefined, "Cancel");

    cancelButton.onClick = function () {
      dialog.close(0);
    };

    okButton.onClick = function () {
      if (
        (!autoPlaceCheckbox.value && (
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

    var autoPlaceAndCenter = autoPlaceCheckbox.value;
    var globalKeepImagesSeparate = globalKeepSeparateCheckbox.value;

    var stepGroups = [];
    var groupImageFilesFiltered = [];
    if (enableStepGroupsCheckbox.value) {
      for (var g = 0; g < stepImageGroups.length; g++) {
        var group = stepImageGroups[g];
        if (group.enableCheckbox.value && groupImageFiles[g] && groupImageFiles[g].length > 0) {
          stepGroups.push({
            count: parseInt(group.countInput.text, 10) || 1,
            rotation: group.portraitRadio.value ? 0 : 90,
            keepSeparate: group.keepSeparateCheckbox.value
          });
          groupImageFilesFiltered.push(groupImageFiles[g]);
        }
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

    // --- AUTO PLACE & CENTER LOGIC ---
    if (autoPlaceAndCenter) {
      function inchesToPoints(val) { return parseFloat(val) * 72; }
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

      var totalImages = placementPlan.length;
      var imgCounter = 0;
      var pageIndex = 0;
      var lastGroupIdx = -1;

      while (imgCounter < totalImages) {
        var page = pageIndex === 0 ? doc.pages[0] : doc.pages.add();

        page.marginPreferences.top = pageTopMargin;
        page.marginPreferences.bottom = pageBottomMargin;
        page.marginPreferences.left = pageLeftMargin;
        page.marginPreferences.right = pageRightMargin;

        var pageBounds = page.bounds;
        var pageTop = pageBounds[0];
        var pageLeft = pageBounds[1];
        var pageWidth = pageBounds[3] - pageBounds[1];
        var pageHeight = pageBounds[2] - pageBounds[0];

        var usableWidth = pageWidth - pageLeftMargin - pageRightMargin;
        var usableHeight = pageHeight - pageTopMargin - pageBottomMargin;

        var currentGroupIdx = placementPlan[imgCounter].groupIdx;
        if (imgCounter > 0 && placementPlan[imgCounter].keepSeparate && currentGroupIdx !== lastGroupIdx) {
          page = doc.pages.add();
          pageIndex++;
        }
        lastGroupIdx = currentGroupIdx;

        var plan = placementPlan[imgCounter];
        var probeFile = File(plan.file);
        var imgW = 0, imgH = 0;
        if (probeFile.exists) {
          var probeRect = page.rectangles.add({geometricBounds: [0,0,72,72]});
          var probePlaced = probeRect.place(probeFile)[0];
          var probeBounds = probePlaced.visibleBounds;
          imgW = probeBounds[3] - probeBounds[1];
          imgH = probeBounds[2] - probeBounds[0];
          probePlaced.remove();
          probeRect.remove();
        }
        if (imgW === 0 || imgH === 0) {
          alert("Could not determine image size for grid calculation. Please check your images.");
          return;
        }
        var best = {
          rotation: plan.rotation,
          frameW: imgW,
          frameH: imgH,
          columns: Math.floor((usableWidth + horizontalGutter) / (imgW + horizontalGutter)),
          rows: Math.floor((usableHeight + verticalGutter) / (imgH + verticalGutter)),
          count: Math.floor((usableWidth + horizontalGutter) / (imgW + horizontalGutter)) *
                 Math.floor((usableHeight + verticalGutter) / (imgH + verticalGutter))
        };

        if (autoRotateBestFitCheckbox.value) {
          var cols0 = Math.floor((usableWidth + horizontalGutter) / (imgW + horizontalGutter));
          var rows0 = Math.floor((usableHeight + verticalGutter) / (imgH + verticalGutter));
          var count0 = cols0 * rows0;
          var cols90 = Math.floor((usableWidth + horizontalGutter) / (imgH + horizontalGutter));
          var rows90 = Math.floor((usableHeight + verticalGutter) / (imgW + verticalGutter));
          var count90 = cols90 * rows90;
          if (count90 > count0) {
            best = {
              rotation: 90,
              frameW: imgH,
              frameH: imgW,
              columns: cols90,
              rows: rows90,
              count: count90
            };
          } else {
            best = {
              rotation: 0,
              frameW: imgW,
              frameH: imgH,
              columns: cols0,
              rows: rows0,
              count: count0
            };
          }
        }

        if (best.columns < 1 || best.rows < 1 || best.count < 1) {
          alert("Could not fit any images on the page with the current margins. Please reduce margins or use smaller images.");
          return;
        }

        var gridWidth = best.columns * best.frameW + (best.columns - 1) * horizontalGutter;
        var gridHeight = best.rows * best.frameH + (best.rows - 1) * verticalGutter;

        var startX = pageLeft + pageLeftMargin + (usableWidth - gridWidth) / 2;
        var startY = pageTop + pageTopMargin + (usableHeight - gridHeight) / 2;

        var frames = [];
var placedOnThisPage = 0;
for (var i = 0; i < best.count && imgCounter < totalImages; i++) {
  var row = Math.floor(i / best.columns);
  var col = i % best.columns;
  var x1 = startX + col * (best.frameW + horizontalGutter);
  var y1 = startY + row * (best.frameH + verticalGutter);
  var x2 = x1 + best.frameW;
  var y2 = y1 + best.frameH;
  if (x2 > pageLeft + pageWidth - pageRightMargin + 0.01 || y2 > pageTop + pageHeight - pageBottomMargin + 0.01) {
    continue;
  }
  var plan = placementPlan[imgCounter];
  var imageFile = File(plan.file);
  if (imageFile.exists) {
    // Place the frame at the intended position and size (unrotated)
    var frame = page.rectangles.add({
      geometricBounds: [y1, x1, y2, x2],
      strokeWeight: 1,
      strokeColor: doc.swatches.itemByName("None")
    });

    // Place the image and auto-rotate for best fit
    var placedImage = frame.place(imageFile)[0];
    if (enableFitting) {
      if (fittingOption === "Content-Aware") {
        frame.fit(FitOptions.CONTENT_AWARE_FIT);
      } else if (fittingOption === "Proportional") {
        frame.fit(FitOptions.FILL_PROPORTIONALLY);
      }
    }
    // If best.rotation is 90, rotate the image inside the frame
    if (best.rotation === 90) {
      placedImage.absoluteRotationAngle = 90;
      frame.fit(FitOptions.CENTER_CONTENT);
    } else {
      frame.fit(FitOptions.CENTER_CONTENT);
    }
    frame.locked = false;
    frames.push(frame);
  }
  imgCounter++;
  placedOnThisPage++;
  if (plan.keepSeparate && placedOnThisPage === best.count) {
    break;
  }
}

// Group, center, ungroup
if (frames.length > 0) {
  var group = page.groups.add(frames);
  var groupBounds = group.visibleBounds;
  var groupWidth = groupBounds[3] - groupBounds[1];
  var groupHeight = groupBounds[2] - groupBounds[0];
  var groupCenterX = groupBounds[1] + groupWidth / 2;
  var groupCenterY = groupBounds[0] + groupHeight / 2;
  var usableCenterX = pageLeft + pageLeftMargin + usableWidth / 2;
  var usableCenterY = pageTop + pageTopMargin + usableHeight / 2;
  var dx = usableCenterX - groupCenterX;
  var dy = usableCenterY - groupCenterY;
  group.move([groupBounds[1] + dx, groupBounds[0] + dy]);
  group.ungroup();
}

        if (doCropMarks && frames.length > 0) {
          app.select(frames);
          drawCropMarks_OuterOnly(
            cropMarkRange,
            true,
            cropMarkLength,
            cropMarkOffset,
            cropMarkWidth,
            bleedOffset
          );
        }
        pageIndex++;
      }
      alert("All images placed successfully!");
      return;
    }

    // --- MANUAL LOGIC (unchanged) ---
    // (You can add your manual grid logic here if needed.)

  } catch (e) {
    alert("An error occurred while placing images: " + e.message);
  
  }
}

// --- Crop marks logic with bleed support ---
function drawCropMarks_OuterOnly(myRange, myDoCropMarks, myCropMarkLength, myCropMarkOffset, myCropMarkWidth, bleedOffset) {
  var myDocument = app.activeDocument;
  var myOldRulerOrigin = myDocument.viewPreferences.rulerOrigin;
  myDocument.viewPreferences.rulerOrigin = RulerOrigin.spreadOrigin;
  var myOldXUnits = myDocument.viewPreferences.horizontalMeasurementUnits;
  var myOldYUnits = myDocument.viewPreferences.verticalMeasurementUnits;
  myDocument.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.points;
  myDocument.viewPreferences.verticalMeasurementUnits = MeasurementUnits.points;
  var myLayer = myDocument.layers.item("myCropMarks");
  try { var myLayerName = myLayer.name; }
  catch (myError) { myLayer = myDocument.layers.add({ name: "myCropMarks" }); }
  var myRegistrationColor = myDocument.colors.item("Registration");
  var myNoneSwatch = myDocument.swatches.item("None");

  var allBounds = [];
  for (var i = 0; i < app.selection.length; i++) {
    allBounds.push(app.selection[i].visibleBounds);
  }
  var tolerance = 0.5; // points

  for (var myCounter = 0; myCounter < app.selection.length; myCounter++) {
    var myObject = app.selection[myCounter];
    var b = myObject.visibleBounds; // [y1, x1, y2, x2]
    var y1 = b[0], x1 = b[1], y2 = b[2], x2 = b[3];

    if (myRange == 0) { // Each Object
      if (myDoCropMarks == true) {
        var topOuter = isEdgeOuter(b, allBounds, "top", tolerance);
        var bottomOuter = isEdgeOuter(b, allBounds, "bottom", tolerance);
        var leftOuter = isEdgeOuter(b, allBounds, "left", tolerance);
        var rightOuter = isEdgeOuter(b, allBounds, "right", tolerance);

        // --- EDGES: Draw original and two parallel bleed marks (offset perpendicular to mark) ---
        if (topOuter) {
          myDrawLine([y1 - myCropMarkOffset, x1, y1 - (myCropMarkOffset + myCropMarkLength), x1], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
          myDrawLine([y1 - myCropMarkOffset, x2, y1 - (myCropMarkOffset + myCropMarkLength), x2], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
          myDrawLine([y1 - myCropMarkOffset, x1 - bleedOffset, y1 - (myCropMarkOffset + myCropMarkLength), x1 - bleedOffset], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
          myDrawLine([y1 - myCropMarkOffset, x2 + bleedOffset, y1 - (myCropMarkOffset + myCropMarkLength), x2 + bleedOffset], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
        }
        if (bottomOuter) {
          myDrawLine([y2 + myCropMarkOffset, x1, y2 + myCropMarkOffset + myCropMarkLength, x1], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
          myDrawLine([y2 + myCropMarkOffset, x2, y2 + myCropMarkOffset + myCropMarkLength, x2], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
          myDrawLine([y2 + myCropMarkOffset, x1 - bleedOffset, y2 + myCropMarkOffset + myCropMarkLength, x1 - bleedOffset], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
          myDrawLine([y2 + myCropMarkOffset, x2 + bleedOffset, y2 + myCropMarkOffset + myCropMarkLength, x2 + bleedOffset], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
        }
        if (leftOuter) {
          myDrawLine([y1, x1 - myCropMarkOffset, y1, x1 - (myCropMarkOffset + myCropMarkLength)], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
          myDrawLine([y2, x1 - myCropMarkOffset, y2, x1 - (myCropMarkOffset + myCropMarkLength)], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
          myDrawLine([y1 - bleedOffset, x1 - myCropMarkOffset, y1 - bleedOffset, x1 - (myCropMarkOffset + myCropMarkLength)], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
          myDrawLine([y2 + bleedOffset, x1 - myCropMarkOffset, y2 + bleedOffset, x1 - (myCropMarkOffset + myCropMarkLength)], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
        }
        if (rightOuter) {
          myDrawLine([y1, x2 + myCropMarkOffset, y1, x2 + myCropMarkOffset + myCropMarkLength], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
          myDrawLine([y2, x2 + myCropMarkOffset, y2, x2 + myCropMarkOffset + myCropMarkLength], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
          myDrawLine([y1 - bleedOffset, x2 + myCropMarkOffset, y1 - bleedOffset, x2 + myCropMarkOffset + myCropMarkLength], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
          myDrawLine([y2 + bleedOffset, x2 + myCropMarkOffset, y2 + bleedOffset, x2 + myCropMarkOffset + myCropMarkLength], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
        }

        // --- CORNERS: Draw original and bleed mark (offset only y for left/right sides, only x for top/bottom sides) ---
        if (topOuter && leftOuter) {
          myDrawLine([y1, x1 - myCropMarkOffset, y1, x1 - (myCropMarkOffset + myCropMarkLength)], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
          myDrawLine([y1 + bleedOffset, x1 - myCropMarkOffset, y1 + bleedOffset, x1 - (myCropMarkOffset + myCropMarkLength)], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
          myDrawLine([y1 - myCropMarkOffset, x1, y1 - (myCropMarkOffset + myCropMarkLength), x1], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
          myDrawLine([y1 - myCropMarkOffset, x1 + bleedOffset, y1 - (myCropMarkOffset + myCropMarkLength), x1 + bleedOffset], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
        }
        if (topOuter && rightOuter) {
          myDrawLine([y1, x2 + myCropMarkOffset, y1, x2 + myCropMarkOffset + myCropMarkLength], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
          myDrawLine([y1 + bleedOffset, x2 + myCropMarkOffset, y1 + bleedOffset, x2 + myCropMarkOffset + myCropMarkLength], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
          myDrawLine([y1 - myCropMarkOffset, x2, y1 - (myCropMarkOffset + myCropMarkLength), x2], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
          myDrawLine([y1 - myCropMarkOffset, x2 - bleedOffset, y1 - (myCropMarkOffset + myCropMarkLength), x2 - bleedOffset], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
        }
        if (bottomOuter && leftOuter) {
          myDrawLine([y2, x1 - myCropMarkOffset, y2, x1 - (myCropMarkOffset + myCropMarkLength)], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
          myDrawLine([y2 - bleedOffset, x1 - myCropMarkOffset, y2 - bleedOffset, x1 - (myCropMarkOffset + myCropMarkLength)], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
          myDrawLine([y2 + myCropMarkOffset, x1, y2 + myCropMarkOffset + myCropMarkLength, x1], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
          myDrawLine([y2 + myCropMarkOffset, x1 + bleedOffset, y2 + myCropMarkOffset + myCropMarkLength, x1 + bleedOffset], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
        }
        if (bottomOuter && rightOuter) {
          myDrawLine([y2, x2 + myCropMarkOffset, y2, x2 + myCropMarkOffset + myCropMarkLength], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
          myDrawLine([y2 - bleedOffset, x2 + myCropMarkOffset, y2 - bleedOffset, x2 + myCropMarkOffset + myCropMarkLength], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
          myDrawLine([y2 + myCropMarkOffset, x2, y2 + myCropMarkOffset + myCropMarkLength, x2], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
          myDrawLine([y2 + myCropMarkOffset, x2 - bleedOffset, y2 + myCropMarkOffset + myCropMarkLength, x2 - bleedOffset], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
        }
      }
    } else { // Entire Selection (bounding box)
      if (myCounter == 0) {
        var myX1 = x1, myY1 = y1, myX2 = x2, myY2 = y2;
      } else {
        if (y1 < myY1) myY1 = y1;
        if (x1 < myX1) myX1 = x1;
        if (y2 > myY2) myY2 = y2;
        if (x2 > myX2) myX2 = x2;
      }
    }
  }
  if (myRange != 0) {
    if (myDoCropMarks == true) {
      myDrawCropMarks(myX1, myY1, myX2, myY2, myCropMarkLength, myCropMarkOffset, myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
      myDrawCropMarks(myX1 - bleedOffset, myY1 - bleedOffset, myX2 + bleedOffset, myY2 + bleedOffset, myCropMarkLength, myCropMarkOffset, myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
      myDrawCropMarks(myX1 + bleedOffset, myY1 + bleedOffset, myX2 - bleedOffset, myY2 - bleedOffset, myCropMarkLength, myCropMarkOffset, myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
    }
  }
  myDocument.viewPreferences.rulerOrigin = myOldRulerOrigin;
  myDocument.viewPreferences.horizontalMeasurementUnits = myOldXUnits;
  myDocument.viewPreferences.verticalMeasurementUnits = myOldYUnits;
}

function isEdgeOuter(myBounds, allBounds, edge, tolerance) {
  for (var i = 0; i < allBounds.length; i++) {
    var other = allBounds[i];
    if (other === myBounds) continue;
    if (edge === "top" && Math.abs(myBounds[0] - other[2]) < tolerance &&
      myBounds[1] < other[3] && myBounds[3] > other[1]) {
      return false;
    }
    if (edge === "bottom" && Math.abs(myBounds[2] - other[0]) < tolerance &&
      myBounds[1] < other[3] && myBounds[3] > other[1]) {
      return false;
    }
    if (edge === "left" && Math.abs(myBounds[1] - other[3]) < tolerance &&
      myBounds[0] < other[2] && myBounds[2] > other[0]) {
      return false;
    }
    if (edge === "right" && Math.abs(myBounds[3] - other[1]) < tolerance &&
      myBounds[0] < other[2] && myBounds[2] > other[0]) {
      return false;
    }
  }
  return true;
}

function myDrawCropMarks(myX1, myY1, myX2, myY2, myCropMarkLength, myCropMarkOffset, myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer) {
  myDrawLine([myY1 - myCropMarkOffset, myX1, myY1 - (myCropMarkOffset + myCropMarkLength), myX1], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
  myDrawLine([myY2 + myCropMarkOffset, myX1, myY2 + myCropMarkOffset + myCropMarkLength, myX1], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
  myDrawLine([myY1 - myCropMarkOffset, myX2, myY1 - (myCropMarkOffset + myCropMarkLength), myX2], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
  myDrawLine([myY2 + myCropMarkOffset, myX2, myY2 + myCropMarkOffset + myCropMarkLength, myX2], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
}

function myDrawLine(myBounds, myStrokeWeight, myRegistrationColor, myNoneSwatch, myLayer) {
  app.activeWindow.activeSpread.graphicLines.add(myLayer, undefined, undefined, {
    strokeWeight: myStrokeWeight,
    fillColor: myNoneSwatch,
    strokeColor: myRegistrationColor,
    geometricBounds: myBounds
  });
}

// Run the script
placeImagesOnDocumentPages();