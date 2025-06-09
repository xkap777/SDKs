#target indesign

function placeImagesOnDocumentPages() {
  // Ensure a document is open and get the active document
  if (!app.documents.length) {
    alert("No document is open. Please open a document and try again.");
    return;
  }
  var doc = app.activeDocument;

  try {
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
    var globalAutoPlaceCheckbox = globalAutoGroup.add('checkbox', undefined, 'Auto Place & Center Images');
    globalAutoPlaceCheckbox.value = false;

    // --- Add global Keep Image Sizes Separate checkbox (only enabled when Auto Place is checked) ---
    var globalKeepSeparateCheckbox = globalAutoGroup.add('checkbox', undefined, 'Keep Image Sizes Separate');
    globalKeepSeparateCheckbox.value = false;
    globalKeepSeparateCheckbox.enabled = false;

    // --- Only enable globalKeepSeparateCheckbox when Auto Place is checked ---
    globalAutoPlaceCheckbox.onClick = function () {
      // If checked, disable all manual fields
      setManualFieldsEnabled(!this.value);
      globalKeepSeparateCheckbox.enabled = !!this.value;
      if (!this.value) globalKeepSeparateCheckbox.value = false;
      updateGlobalAutoTabEnabled();
    };
    // --- Disable globalKeepSeparateCheckbox if Auto Place is unchecked at dialog start ---
    globalKeepSeparateCheckbox.enabled = !!globalAutoPlaceCheckbox.value;

    // --- When globalKeepSeparateCheckbox is enabled, allow user to toggle it, otherwise always set to false ---
    globalKeepSeparateCheckbox.onClick = function () {
      if (!globalKeepSeparateCheckbox.enabled) {
        globalKeepSeparateCheckbox.value = false;
      }
    };

    var autoRotateBestFitCheckbox = globalAutoGroup.add('checkbox', undefined, 'Auto Rotate for Best Fit (Auto Place only)');
    autoRotateBestFitCheckbox.value = false;
    autoRotateBestFitCheckbox.enabled = true; // Always enabled

    // Frame Dimensions (GLOBAL)
    var frameGroup = dialog.add("panel", undefined, "Frame Dimensions");
    frameGroup.orientation = "column";
    frameGroup.alignChildren = "left";
    frameGroup.margins = 10;

    var frameRow1 = frameGroup.add("group");
    frameRow1.orientation = "row";
    frameRow1.add("statictext", undefined, "Frame Width (in):");
    var globalFrameWidthInput = frameRow1.add("edittext", undefined, "0");
    globalFrameWidthInput.characters = 5;
    frameRow1.add("statictext", undefined, "Frame Height (in):");
    var globalFrameHeightInput = frameRow1.add("edittext", undefined, "0");
    globalFrameHeightInput.characters = 5;

    var frameRow2 = frameGroup.add("group");
    frameRow2.orientation = "row";
    frameRow2.add("statictext", undefined, "Rows:");
    var globalRowsInput = frameRow2.add("edittext", undefined, "0");
    globalRowsInput.characters = 3;
    frameRow2.add("statictext", undefined, "Columns:");
    var globalColumnsInput = frameRow2.add("edittext", undefined, "0");
    globalColumnsInput.characters = 3;

    var frameRow3 = frameGroup.add("group");
    frameRow3.orientation = "row";
    frameRow3.add("statictext", undefined, "Horizontal Gutter (in):");
    var globalHorizontalGutterInput = frameRow3.add("edittext", undefined, "0");
    globalHorizontalGutterInput.characters = 4;
    frameRow3.add("statictext", undefined, "Vertical Gutter (in):");
    var globalVerticalGutterInput = frameRow3.add("edittext", undefined, "0");
    globalVerticalGutterInput.characters = 4;

    // Frame Margins (GLOBAL)
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
    var globalFrameTopMarginInput = frameMarginRow1.add("edittext", undefined, "0");
    globalFrameTopMarginInput.characters = 5;
    frameMarginRow1.add("statictext", undefined, "Bottom (in):");
    var globalFrameBottomMarginInput = frameMarginRow1.add("edittext", undefined, "0");
    globalFrameBottomMarginInput.characters = 5;

    var frameMarginRow2 = frameMarginGroup.add("group");
    frameMarginRow2.orientation = "row";
    frameMarginRow2.add("statictext", undefined, "Left (in):");
    var globalFrameLeftMarginInput = frameMarginRow2.add("edittext", undefined, "0");
    globalFrameLeftMarginInput.characters = 5;
    frameMarginRow2.add("statictext", undefined, "Right (in):");
    var globalFrameRightMarginInput = frameMarginRow2.add("edittext", undefined, "0");
    globalFrameRightMarginInput.characters = 5;

    // Page Margins (GLOBAL)
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
    var globalPageTopMarginInput = pageMarginRow1.add("edittext", undefined, "0");
    globalPageTopMarginInput.characters = 5;
    pageMarginRow1.add("statictext", undefined, "Bottom (in):");
    var globalPageBottomMarginInput = pageMarginRow1.add("edittext", undefined, "0");
    globalPageBottomMarginInput.characters = 5;

    var pageMarginRow2 = pageMarginGroup.add("group");
    pageMarginRow2.orientation = "row";
    pageMarginRow2.add("statictext", undefined, "Left (in):");
    var globalPageLeftMarginInput = pageMarginRow2.add("edittext", undefined, "0");
    globalPageLeftMarginInput.characters = 5;
    pageMarginRow2.add("statictext", undefined, "Right (in):");
    var globalPageRightMarginInput = pageMarginRow2.add("edittext", undefined, "0");
    globalPageRightMarginInput.characters = 5;

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

    // Add global Keep Groups Separate checkbox
    var globalKeepGroupsSeparateCheckbox = stepImagesPanel.add("checkbox", undefined, "Keep Groups Separate");
    globalKeepGroupsSeparateCheckbox.value = false;
    globalKeepGroupsSeparateCheckbox.enabled = true;
    // Hide unless step groups are enabled
    globalKeepGroupsSeparateCheckbox.visible = false;

    // Container for dynamic group panels
    var groupPanelContainer = stepImagesPanel.add("group");
    groupPanelContainer.orientation = "column";
    groupPanelContainer.alignChildren = "left";
    groupPanelContainer.visible = false; // Hide initially

    var addGroupBtn = stepImagesPanel.add("button", undefined, "Add Group");
    addGroupBtn.visible = false;

    var stepImageGroups = [];
    var maxStepGroups = 9;
    var groupImageFiles = [];

    // Show/hide group controls dynamically
    enableStepGroupsCheckbox.onClick = function () {
      groupPanelContainer.visible = this.value;
      addGroupBtn.visible = this.value;
      globalKeepGroupsSeparateCheckbox.visible = this.value;
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
      updateManualFieldsState();
    };

    // --- Helper: Disable global auto tab checkboxes if any group exists ---
    function updateGlobalAutoTabEnabled() {
      var anyGroupExists = stepImageGroups.length > 0;
      if (anyGroupExists) {
        globalAutoPlaceCheckbox.enabled = false;
        globalAutoPlaceCheckbox.value = false;
        globalKeepSeparateCheckbox.enabled = false;
        globalKeepSeparateCheckbox.value = false;
      } else {
        globalAutoPlaceCheckbox.enabled = true;
        globalKeepSeparateCheckbox.enabled = globalAutoPlaceCheckbox.value;
      }
    }

    // --- Ensure manual fields are updated on globalAutoPlaceCheckbox change even if groups are not enabled ---
    globalAutoPlaceCheckbox.onClick = function () {
      updateGlobalAutoTabEnabled();
      globalKeepSeparateCheckbox.enabled = !!this.value && globalAutoPlaceCheckbox.enabled;
      if (!this.value) globalKeepSeparateCheckbox.value = false;
      updateManualFieldsState(); // Ensure manual fields are updated only via centralized logic
    };

    // Add Group button logic
    addGroupBtn.onClick = function () {
      if (stepImageGroups.length >= maxStepGroups) return;
      var g = stepImageGroups.length;
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
      enableGroupCheckbox.enabled = true; // Ensure checkbox is enabled for new group

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

      // --- Apply button for manual entry ---
      var applyBtn = controlsRow.add("button", undefined, "Apply");
      applyBtn.enabled = true;

      // --- Reset/Remove buttons for manual settings ---
      var resetBtn = controlsRow.add("button", undefined, "Reset");
      var removeBtn = controlsRow.add("button", undefined, "Remove");

      // --- Disable all group controls except Enable by default ---
      countInput.enabled = false;
      keepSeparateCheckbox.enabled = false;
      autoPlaceCheckbox.enabled = false;
      autoRotateCheckbox.enabled = false;
      portraitRadio.enabled = false;
      landscapeRadio.enabled = false;
      applyBtn.enabled = false;
      resetBtn.enabled = false;

      // --- Remove group logic ---
      removeBtn.onClick = function () {
        groupPanel.parent.remove(groupPanel);
        for (var i = 0; i < stepImageGroups.length; i++) {
          if (stepImageGroups[i].groupPanel === groupPanel) {
            stepImageGroups.splice(i, 1);
            break;
          }
        }
        groupImageFiles[g] = [];
        dialog.layout.layout(true);
        dialog.layout.resize();
        updateGlobalAutoTabEnabled();
        updateManualFieldsState();
      };

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
      // Always show summary, update on any field change
      function updateFrameSettingsSummary() {
        frameSettingsText.text = formatFrameSettings(groupFrameSettings);
        try { frameSettingsText.justify = 'right'; } catch (e) {}
        summaryGroup.layout.layout(true);
        groupHeaderRow.layout.layout(true);
        if (groupPanel && groupPanel.layout) groupPanel.layout.layout(true);
      }
      // Update groupFrameSettings live as user edits fields
      function updateGroupFrameSettings() {
        groupFrameSettings.frameWidth = parseFloat(globalFrameWidthInput.text) || 0;
        groupFrameSettings.frameHeight = parseFloat(globalFrameHeightInput.text) || 0;
        groupFrameSettings.rows = parseInt(globalRowsInput.text, 10) || 0;
        groupFrameSettings.columns = parseInt(globalColumnsInput.text, 10) || 0;
        groupFrameSettings.horizontalGutter = parseFloat(globalHorizontalGutterInput.text) || 0;
        groupFrameSettings.verticalGutter = parseFloat(globalVerticalGutterInput.text) || 0;
        updateFrameSettingsSummary();
      }
      // Attach live update to relevant fields
      globalFrameWidthInput.onChanging = updateGroupFrameSettings;
      globalFrameHeightInput.onChanging = updateGroupFrameSettings;
      globalRowsInput.onChanging = updateGroupFrameSettings;
      globalColumnsInput.onChanging = updateGroupFrameSettings;
      globalHorizontalGutterInput.onChanging = updateGroupFrameSettings;
      globalVerticalGutterInput.onChanging = updateGroupFrameSettings;

      // --- Apply button logic ---
      applyBtn.onClick = function () {
        groupFrameSettings.frameWidth = parseFloat(globalFrameWidthInput.text) || 0;
        groupFrameSettings.frameHeight = parseFloat(globalFrameHeightInput.text) || 0;
        groupFrameSettings.rows = parseInt(globalRowsInput.text, 10) || 0;
        groupFrameSettings.columns = parseInt(globalColumnsInput.text, 10) || 0;
        groupFrameSettings.horizontalGutter = parseFloat(globalHorizontalGutterInput.text) || 0;
        groupFrameSettings.verticalGutter = parseFloat(globalVerticalGutterInput.text) || 0;
        updateFrameSettingsSummary();
        applyBtn.text = "Applied";
        $.sleep(300);
        applyBtn.text = "Apply";
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
        keepSeparateCheckbox.value = false;
        autoPlaceCheckbox.value = false;
        autoRotateCheckbox.value = false;
        portraitRadio.value = true;
        landscapeRadio.value = false;
        applyBtn.enabled = true;
        countInput.enabled = true;
        portraitRadio.enabled = true;
        landscapeRadio.enabled = true;
        globalHorizontalGutterInput.enabled = true;
        globalVerticalGutterInput.enabled = true;
        globalFrameWidthInput.enabled = true;
        globalFrameHeightInput.enabled = true;
        globalRowsInput.enabled = true;
        globalColumnsInput.enabled = true;
      };

      // --- Checkbox logic for auto place/auto rotate ---
      function updateAutoModeUI() {
        // If auto rotate is checked, disable portrait/landscape
        if (autoRotateCheckbox.value) {
          portraitRadio.enabled = false;
          landscapeRadio.enabled = false;
          // Disable global auto rotate if any group auto rotate is checked
          autoRotateBestFitCheckbox.value = false;
          autoRotateBestFitCheckbox.enabled = false;
        } else {
          portraitRadio.enabled = enableGroupCheckbox.value && !autoRotateBestFitCheckbox.value && !autoPlaceCheckbox.value;
          landscapeRadio.enabled = enableGroupCheckbox.value && !autoRotateBestFitCheckbox.value && !autoPlaceCheckbox.value;
          // Re-enable global auto rotate if no group auto rotate is checked
          var anyGroupAutoRotate = false;
          for (var i = 0; i < stepImageGroups.length; i++) {
            if (stepImageGroups[i].enableCheckbox.value && stepImageGroups[i].autoRotateCheckbox.value) {
              anyGroupAutoRotate = true;
              break;
            }
          }
          autoRotateBestFitCheckbox.enabled = !anyGroupAutoRotate;
        }
        // Always enable group auto checkboxes and Reset if group is enabled
        autoPlaceCheckbox.enabled = enableGroupCheckbox.value;
        autoRotateCheckbox.enabled = enableGroupCheckbox.value;
        resetBtn.enabled = enableGroupCheckbox.value; // Always enabled if group is enabled
        // Manual fields enable/disable is now always handled by updateManualFieldsState()
      }

      // Ensure all group autoPlace/autoRotate checkboxes call updateManualFieldsState
      autoPlaceCheckbox.onClick = function () {
        updateAutoModeUI();
        updateGlobalAutoTabEnabled();
        updateManualFieldsState();
      };
      autoRotateCheckbox.onClick = function () {
        updateAutoModeUI();
        updateGlobalAutoTabEnabled();
        updateManualFieldsState();
      };
      enableGroupCheckbox.onClick = function () {
        var enabled = this.value;
        // Dynamically determine the group index for this groupPanel
        var groupIdx = -1;
        for (var i = 0; i < stepImageGroups.length; i++) {
          if (stepImageGroups[i].groupPanel === groupPanel) {
            groupIdx = i;
            break;
          }
        }
        // Enable/disable all group controls except Enable
        countInput.enabled = enabled;
        keepSeparateCheckbox.enabled = enabled;
        autoPlaceCheckbox.enabled = enabled;
        autoRotateCheckbox.enabled = enabled;
        portraitRadio.enabled = enabled;
        landscapeRadio.enabled = enabled;
        applyBtn.enabled = enabled;
        resetBtn.enabled = enabled;
        var rotationEnabled = enabled && !autoRotateBestFitCheckbox.value;
        if (enabled) {
          groupPanel.text = "Step Group";
        } else {
          groupImageFiles[groupIdx] = [];
          groupPanel.text = "Step Group";
          autoPlaceCheckbox.value = false;
          autoRotateCheckbox.value = false;
        }
        updateAutoModeUI();
        updateGlobalAutoTabEnabled();
        updateManualFieldsState(); // Always call centralized logic
      };
      // Add the group to the stepImageGroups array so it is tracked and interactive
      stepImageGroups.push({
        groupPanel: groupPanel,
        enableCheckbox: enableGroupCheckbox,
        countInput: countInput,
        portraitRadio: portraitRadio,
        landscapeRadio: landscapeRadio,
        keepSeparateCheckbox: keepSeparateCheckbox,
        autoPlaceCheckbox: autoPlaceCheckbox,
        autoRotateCheckbox: autoRotateCheckbox
      });
      groupPanel.visible = true;
      dialog.layout.layout(true);
      dialog.layout.resize();
      updateGlobalAutoTabEnabled();
      updateManualFieldsState();
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

    // --- Centralized manual fields enable/disable logic ---
    function updateManualFieldsState() {
      var anyGroupExists = stepImageGroups && stepImageGroups.length > 0;
      var manualEnabled = true;
      if (anyGroupExists) {
        // Disable global auto options
        globalAutoPlaceCheckbox.enabled = false;
        globalAutoPlaceCheckbox.value = false;
        globalKeepSeparateCheckbox.enabled = false;
        globalKeepSeparateCheckbox.value = false;
        // Manual fields are enabled unless any group has autoPlace or autoRotate checked
        for (var i = 0; i < stepImageGroups.length; i++) {
          var group = stepImageGroups[i];
          if (group.enableCheckbox.value && (group.autoPlaceCheckbox.value || group.autoRotateCheckbox.value)) {
            manualEnabled = false;
            break;
          }
        }
      } else {
        // No groups: global auto options enabled
        globalAutoPlaceCheckbox.enabled = true;
        globalKeepSeparateCheckbox.enabled = globalAutoPlaceCheckbox.value;
        // Manual fields disabled if global autoPlace or autoRotate is checked
        manualEnabled = !(globalAutoPlaceCheckbox.value || autoRotateBestFitCheckbox.value);
      }
      setManualFieldsEnabled(manualEnabled);
    }

    function setManualFieldsEnabled(enabled) {
      // Only affect global/manual fields, not group fields
      frameGroup.enabled = enabled;
      frameMarginGroup.enabled = enabled;
      pageMarginGroup.enabled = enabled;
      fittingGroup.enabled = enabled;
      enableFittingCheckbox.enabled = enabled;
      fittingDropdown.enabled = enabled && enableFittingCheckbox.value;
      rotationGroup.enabled = enabled;
      rotationDropdown.enabled = enabled;
      // All global input fields
      globalFrameWidthInput.enabled = enabled;
      globalFrameHeightInput.enabled = enabled;
      globalRowsInput.enabled = enabled;
      globalColumnsInput.enabled = enabled;
      globalHorizontalGutterInput.enabled = enabled;
      globalVerticalGutterInput.enabled = enabled;
      globalFrameTopMarginInput.enabled = enabled;
      globalFrameBottomMarginInput.enabled = enabled;
      globalFrameLeftMarginInput.enabled = enabled;
      globalFrameRightMarginInput.enabled = enabled;
      globalPageTopMarginInput.enabled = enabled;
      globalPageBottomMarginInput.enabled = enabled;
      globalPageLeftMarginInput.enabled = enabled;
      globalPageRightMarginInput.enabled = enabled;
      // Force UI update
      if (dialog && dialog.layout) {
        dialog.layout.layout(true);
        dialog.layout.resize();
      }
    }

    // --- Ensure manual fields and global auto options are updated on relevant changes ---
    globalAutoPlaceCheckbox.onClick = function () { updateManualFieldsState(); };
    autoRotateBestFitCheckbox.onClick = function () { updateManualFieldsState(); };
    // Removed duplicate/empty addGroupBtn.onClick assignment here
    // In each group, after autoPlaceCheckbox or autoRotateCheckbox is toggled:
    // autoPlaceCheckbox.onClick = function () { updateManualFieldsState(); };
    // autoRotateCheckbox.onClick = function () { updateManualFieldsState(); };
    // ...existing code...

    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "center";
    buttonGroup.spacing = 20;

    var okButton = buttonGroup.add("button", undefined, "OK");
    var cancelButton = buttonGroup.add("button", undefined, "Cancel");

    cancelButton.onClick = function () {
      dialog.close(0);
    };

    okButton.onClick = function () {
      // --- Validation: skip manual field validation if either Auto Place or Auto Rotate is checked ---
      var skipManualValidation = false;
      if (enableStepGroupsCheckbox.value) {
        for (var g = 0; g < stepImageGroups.length; g++) {
          var group = stepImageGroups[g];
          if (group.enableCheckbox.value && (group.autoPlaceCheckbox.value || group.autoRotateCheckbox.value)) {
            skipManualValidation = true;
            break;
          }
        }
      } else {
        skipManualValidation = globalAutoPlaceCheckbox.value || autoRotateBestFitCheckbox.value;
      }
      if (
        (!skipManualValidation && (
          isNaN(globalFrameWidthInput.text) ||
          isNaN(globalFrameHeightInput.text) ||
          isNaN(globalRowsInput.text) ||
          isNaN(globalColumnsInput.text) ||
          globalFrameWidthInput.text <= 0 ||
          globalFrameHeightInput.text <= 0 ||
          globalRowsInput.text <= 0 ||
          globalColumnsInput.text <= 0
        )) ||
        isNaN(globalHorizontalGutterInput.text) ||
        isNaN(globalVerticalGutterInput.text) ||
        isNaN(globalFrameTopMarginInput.text) ||
        isNaN(globalFrameBottomMarginInput.text) ||
        isNaN(globalFrameLeftMarginInput.text) ||
        isNaN(globalFrameRightMarginInput.text) ||
        isNaN(globalPageTopMarginInput.text) ||
        isNaN(globalPageBottomMarginInput.text) ||
        isNaN(globalPageLeftMarginInput.text) ||
        isNaN(globalPageRightMarginInput.text) ||
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
      // --- After validation, prompt for image files for each enabled group ---
      if (enableStepGroupsCheckbox.value) {
        for (var g = 0; g < stepImageGroups.length; g++) {
          var group = stepImageGroups[g];
          if (group.enableCheckbox.value) {
            var files = File.openDialog("Select image files for Group " + (g+1), "*.jpg;*.png;*.tif", true);
            if (!files || files.length === 0) {
              alert("No images selected for Group " + (g+1) + ". Operation canceled.");
              return;
            }
            groupImageFiles[g] = files;
            group.groupPanel.text = "Step Group (" + files.length + " images)";
          } else {
            groupImageFiles[g] = [];
            group.groupPanel.text = "Step Group";
          }
        }
      }
      dialog.close(1);
    };

    if (dialog.show() !== 1) {
      alert("User cancelled the operation.");
      return;
    }

    // --- Retrieve user input ---
    var frameWidth = parseFloat(globalFrameWidthInput.text);
    var frameHeight = parseFloat(globalFrameHeightInput.text);
    var rows = parseInt(globalRowsInput.text, 10);
    var columns = parseInt(globalColumnsInput.text, 10);
    var horizontalGutter = parseFloat(globalHorizontalGutterInput.text);
    var verticalGutter = parseFloat(globalVerticalGutterInput.text);

    var frameTopMargin = parseFloat(globalFrameTopMarginInput.text);
    var frameBottomMargin = parseFloat(globalFrameBottomMarginInput.text);
    var frameLeftMargin = parseFloat(globalFrameLeftMarginInput.text);
    var frameRightMargin = parseFloat(globalFrameRightMarginInput.text);

    var pageTopMargin = parseFloat(globalPageTopMarginInput.text);
    var pageBottomMargin = parseFloat(globalPageBottomMarginInput.text);
    var pageLeftMargin = parseFloat(globalPageLeftMarginInput.text);
    var pageRightMargin = parseFloat(globalPageRightMarginInput.text);

    var enableFitting = enableFittingCheckbox.value;
    var fittingOption = fittingDropdown.selection.text;
    var rotationAngle = parseInt(rotationDropdown.selection.text, 10);

    var doCropMarks = enableCropMarks.value;
    var cropMarkLength = parseFloat(cropLengthInput.text);
    var cropMarkOffset = parseFloat(cropOffsetInput.text);
    var cropMarkWidth = parseFloat(cropWeightInput.text);
    var cropMarkRange = cropMarksRangeEach.value ? 0 : 1;
    var bleedOffset = parseFloat(bleedOffsetInput.text);

    var autoPlaceAndCenter = globalAutoPlaceCheckbox.value;
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
    if (globalKeepGroupsSeparateCheckbox.value) {
      // Each group is processed independently
      for (var g = 0; g < stepGroups.length; g++) {
        for (var i = 0; i < groupImageFilesFiltered[g].length; i++) {
          for (var c = 0; c < stepGroups[g].count; c++) {
            placementPlan.push({
              file: groupImageFilesFiltered[g][i],
              rotation: stepGroups[g].rotation,
              groupIdx: g,
              keepSeparate: true, // force group separation
              keepImageSizes: stepGroups[g].keepSeparate
            });
          }
        }
      }
    } else {
      // Pool images from groups with keepSeparate unchecked
      var pooledImages = [];
      var pooledRotations = [];
      var pooledGroupIdxs = [];
      for (var g = 0; g < stepGroups.length; g++) {
        if (stepGroups[g].keepSeparate) {
          // Process this group independently
          for (var i = 0; i < groupImageFilesFiltered[g].length; i++) {
            for (var c = 0; c < stepGroups[g].count; c++) {
              placementPlan.push({
                file: groupImageFilesFiltered[g][i],
                rotation: stepGroups[g].rotation,
                groupIdx: g,
                keepSeparate: true,
                keepImageSizes: true
              });
            }
          }
        } else {
          // Pool these images for joint placement
          for (var i = 0; i < groupImageFilesFiltered[g].length; i++) {
            for (var c = 0; c < stepGroups[g].count; c++) {
              pooledImages.push(groupImageFilesFiltered[g][i]);
              pooledRotations.push(stepGroups[g].rotation);
              pooledGroupIdxs.push(g);
            }
          }
        }
      }
      // Add pooled images as a single pseudo-group
      if (pooledImages.length > 0) {
        for (var i = 0; i < pooledImages.length; i++) {
          placementPlan.push({
            file: pooledImages[i],
            rotation: pooledRotations[i],
            groupIdx: pooledGroupIdxs[i],
            keepSeparate: false,
            keepImageSizes: false
          });
        }
      }
    }

    // --- DEBUG: Check for empty placement plan ---
    if (placementPlan.length === 0) {
      alert('No images to place. Please check your group settings and image selections.');
      return;
    }

    // --- AUTO PLACE & CENTER LOGIC ---
    var shouldAutoPlace = false;
    if (enableStepGroupsCheckbox.value) {
      // If any enabled group has Auto Place checked, we should auto place
      for (var g = 0; g < stepImageGroups.length; g++) {
        var group = stepImageGroups[g];
        if (group.enableCheckbox.value && group.autoPlaceCheckbox.value) {
          shouldAutoPlace = true;
          break;
        }
      }
    } else {
      shouldAutoPlace = autoPlaceAndCenter;
    }
    if (shouldAutoPlace) {
      // --- Progress Palette UI ---
      var progressWin = new Window('palette', 'Placing Images...');
      progressWin.orientation = 'column';
      progressWin.alignChildren = 'center';
      progressWin.spacing = 10;
      var msgGroup = progressWin.add('group');
      msgGroup.alignment = ['center', 'top'];
      var msgText = msgGroup.add('statictext', undefined, 'Please be patient. This might take a while.');
      msgText.justify = 'center';
      msgText.graphics.font = ScriptUI.newFont(msgText.graphics.font.family, msgText.graphics.font.style, msgText.graphics.font.size + 1);
      var iconGroup = progressWin.add('group');
      iconGroup.alignment = ['center', 'top'];
      var iconText = iconGroup.add('statictext', undefined, '\\');
      iconText.characters = 2;
      var progressText = progressWin.add('statictext', undefined, 'Processing...');
      progressText.alignment = ['center', 'top'];
      var cancelBtn = progressWin.add('button', undefined, 'Cancel');
      var userCancelled = false;
      cancelBtn.onClick = function() {
        userCancelled = true;
        progressText.text = 'Cancelling...';
      };
      progressWin.show();
      // Animate for 1 second before starting (optional)
      var delayStart = new Date().getTime();
      var icons = ['|', '/', '-', '\\'];
      var animIdx = 0;
      while (new Date().getTime() - delayStart < 1000 && !userCancelled) {
        iconText.text = icons[animIdx % 4];
        progressWin.update();
        animIdx++;
        $.sleep(100);
      }
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

      // --- GROUP PLACEMENT LOGIC ---
      if (enableStepGroupsCheckbox.value) {
        for (var g = 0; g < stepGroups.length; g++) {
          var groupFiles = groupImageFilesFiltered[g];
          if (!groupFiles || groupFiles.length === 0) continue;
          // Probe all images for their dimensions
          var sizeMap = {};
          var sizeOrder = [];
          for (var i = 0; i < groupFiles.length; i++) {
            var f = File(groupFiles[i]);
            if (!f.exists) continue;
            var probeRect = doc.pages[0].rectangles.add({geometricBounds: [0,0,72,72]});
            var probePlaced = probeRect.place(f)[0];
            var bounds = probePlaced.visibleBounds;
            var w = bounds[3] - bounds[1];
            var h = bounds[2] - bounds[0];
            probePlaced.remove();
            probeRect.remove();
            var key = w.toFixed(2) + 'x' + h.toFixed(2);
            if (!sizeMap[key]) {
              sizeMap[key] = [];
              sizeOrder.push(key);
            }
            sizeMap[key].push({file: f, width: w, height: h});
          }
          // For each unique size group, create pages as needed
          for (var s = 0; s < sizeOrder.length; s++) {
            var key = sizeOrder[s];
            var images = sizeMap[key];
            var imgW = images[0].width;
            var imgH = images[0].height;
            // Calculate how many fit per page
            var usableWidth = doc.documentPreferences.pageWidth - pageLeftMargin - pageRightMargin;
            var usableHeight = doc.documentPreferences.pageHeight - pageTopMargin - pageBottomMargin;
            var cols = Math.max(1, Math.floor((usableWidth + horizontalGutter) / (imgW + horizontalGutter)));
            var rows = Math.max(1, Math.floor((usableHeight + verticalGutter) / (imgH + verticalGutter)));
            var perPage = rows * cols;
            var placed = 0;
            var lastAnimTime = new Date().getTime();
            var animIdxLocal = 0;
            while (placed < images.length) {
              var page = (g === 0 && s === 0 && placed === 0) ? doc.pages[0] : doc.pages.add();
              page.marginPreferences.top = pageTopMargin;
              page.marginPreferences.bottom = pageBottomMargin;
              page.marginPreferences.left = pageLeftMargin;
              page.marginPreferences.right = pageRightMargin;
              var start = placed;
              var end = Math.min(placed + perPage, images.length);
              var numThisPage = end - start;
              // Always use full grid for centering
              var rowsThisPage = rows;
              var colsThisPage = cols;
              var gridWidth = colsThisPage * imgW + (colsThisPage - 1) * horizontalGutter;
              var gridHeight = rowsThisPage * imgH + (rowsThisPage - 1) * verticalGutter;
              var usableWidth = doc.documentPreferences.pageWidth - pageLeftMargin - pageRightMargin;
              var usableHeight = doc.documentPreferences.pageHeight - pageTopMargin - pageBottomMargin;
              var xOffset = pageLeftMargin + (usableWidth - gridWidth) / 2;
              var yOffset = pageTopMargin + (usableHeight - gridHeight) / 2;
              var frames = [];
              for (var relIdx = 0; relIdx < perPage; relIdx++) {
                var row = Math.floor(relIdx / cols);
                var col = relIdx % cols;
                var x1 = xOffset + col * (imgW + horizontalGutter);
                var y1 = yOffset + row * (imgH + verticalGutter);
                var x2 = x1 + imgW;
                var y2 = y1 + imgH;
                var frame = null;
                try {
                  frame = page.rectangles.add({geometricBounds: [y1, x1, y2, x2]});
                } catch (e) {
                  frame = null;
                }
                if (frame && frame.isValid) {
                  // Place image if available
                  if (relIdx < numThisPage) {
                    try {
                      frame.place(images[start + relIdx].file);
                    } catch (e) {}
                  }
                  frames.push(frame);
                }
                // --- Progress indicator update every 3 seconds ---
                var now = new Date().getTime();
                if (now - lastAnimTime >= 3000) {
                  iconText.text = icons[animIdxLocal % 4];
                  animIdxLocal++;
                  progressText.text = 'Placing images... (' + (placed + relIdx + 1) + ' of ' + images.length + ')';
                  progressWin.update();
                  lastAnimTime = now;
                }
                if (userCancelled) {
                  alert('Operation cancelled by user.');
                  progressWin.close();
                  throw new Error('User cancelled');
                }
              }
              // --- Only select valid, placed frames for crop marks ---
              if (doCropMarks && frames.length > 0) {
                var placedFrames = frames.slice(0, numThisPage);
                var validFrames = [];
                for (var i = 0; i < placedFrames.length; i++) {
                  try {
                    if (placedFrames[i] && placedFrames[i].isValid) validFrames.push(placedFrames[i]);
                  } catch (e) {}
                }
                if (validFrames.length > 0) {
                  app.select(validFrames);
                  drawCropMarks_OuterOnly(
                    cropMarkRange, // myRange
                    doCropMarks,   // myDoCropMarks
                    cropMarkLength, // myCropMarkLength
                    cropMarkOffset, // myCropMarkOffset
                    cropMarkWidth,  // myCropMarkWidth
                    bleedOffset     // bleedOffset
                  );
                }
              }
              // Delete empty frames (those without images)
              for (var f = numThisPage; f < frames.length; f++) {
                try {
                  if (frames[f] && frames[f].isValid) frames[f].remove();
                } catch (e) {}
              }
              placed += numThisPage;
            } // <-- Close while (placed < images.length)
          } // <-- Close for (var s = 0; s < sizeOrder.length; s++)
        } // <-- Close for (var g = 0; g < stepGroups.length; g++)
        alert('All group images placed successfully!');
        return;
      } // closes if (enableStepGroupsCheckbox.value)
    } // closes if (shouldAutoPlace)

    // --- MANUAL LOGIC (unchanged) ---
    // (You can add your manual grid logic here if needed.)
  } catch (e) {
    alert("An error occurred while placing images: " + e.message);
  } finally {
    // Always restore measurement units to inches (points) after script runs
    try {
      app.activeDocument.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.INCHES;
      app.activeDocument.viewPreferences.verticalMeasurementUnits = MeasurementUnits.INCHES;
    } catch (e) {
      // If no document is open, do nothing
    }
  }

  // --- Crop marks logic: MyCropMarks (integrated) ---
  function drawCropMarks_OuterOnly(myRange, myDoCropMarks, myCropMarkLength, myCropMarkOffset, myCropMarkWidth, bleedOffset) {
    if (!myDoCropMarks) return;
    var doc = app.activeDocument;
    var oldRulerOrigin = doc.viewPreferences.rulerOrigin;
    var oldXUnits = doc.viewPreferences.horizontalMeasurementUnits;
    var oldYUnits = doc.viewPreferences.verticalMeasurementUnits;
    doc.viewPreferences.rulerOrigin = RulerOrigin.spreadOrigin;
    doc.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.points;
    doc.viewPreferences.verticalMeasurementUnits = MeasurementUnits.points;
    // Ensure crop marks layer exists and is unlocked/visible
    var cropLayer;
    try {
      cropLayer = doc.layers.item("myCropMarks");
      var _ = cropLayer.name; // force error if doesn't exist
    } catch (e) {
      cropLayer = doc.layers.add({ name: "myCropMarks" });
    }
    cropLayer.locked = false;
    cropLayer.visible = true;
    var regColor = doc.colors.item("Registration");
    var noneSwatch = doc.swatches.item("None");
    var selection = app.selection;
    if (!selection || selection.length === 0) return;
    var allBounds = [];
    for (var i = 0; i < selection.length; i++) {
      if (selection[i] && selection[i].isValid) allBounds.push(selection[i].visibleBounds);
    }
    var tolerance = 0.5;
    for (var i = 0; i < selection.length; i++) {
      var obj = selection[i];
      if (!obj || !obj.isValid) continue;
      var b = obj.visibleBounds;
      var y1 = b[0], x1 = b[1], y2 = b[2], x2 = b[3];
      if (myRange == 0) { // Each Object
        var topOuter = isEdgeOuter(b, allBounds, "top", tolerance);
        var bottomOuter = isEdgeOuter(b, allBounds, "bottom", tolerance);
        var leftOuter = isEdgeOuter(b, allBounds, "left", tolerance);
        var rightOuter = isEdgeOuter(b, allBounds, "right", tolerance);
        // Top
        if (topOuter) {
          myDrawLine([y1 - myCropMarkOffset, x1, y1 - (myCropMarkOffset + myCropMarkLength), x1], myCropMarkWidth, regColor, noneSwatch, cropLayer);
          myDrawLine([y1 - myCropMarkOffset, x2, y1 - (myCropMarkOffset + myCropMarkLength), x2], myCropMarkWidth, regColor, noneSwatch, cropLayer);
          myDrawLine([y1 - myCropMarkOffset, x1 - bleedOffset, y1 - (myCropMarkOffset + myCropMarkLength), x1 - bleedOffset], myCropMarkWidth, regColor, noneSwatch, cropLayer);
          myDrawLine([y1 - myCropMarkOffset, x2 + bleedOffset, y1 - (myCropMarkOffset + myCropMarkLength), x2 + bleedOffset], myCropMarkWidth, regColor, noneSwatch, cropLayer);
        }
        // Bottom
        if (bottomOuter) {
          myDrawLine([y2 + myCropMarkOffset, x1, y2 + myCropMarkOffset + myCropMarkLength, x1], myCropMarkWidth, regColor, noneSwatch, cropLayer);
          myDrawLine([y2 + myCropMarkOffset, x2, y2 + myCropMarkOffset + myCropMarkLength, x2], myCropMarkWidth, regColor, noneSwatch, cropLayer);
          myDrawLine([y2 + myCropMarkOffset, x1 - bleedOffset, y2 + myCropMarkOffset + myCropMarkLength, x1 - bleedOffset], myCropMarkWidth, regColor, noneSwatch, cropLayer);
          myDrawLine([y2 + myCropMarkOffset, x2 + bleedOffset, y2 + myCropMarkOffset + myCropMarkLength, x2 + bleedOffset], myCropMarkWidth, regColor, noneSwatch, cropLayer);
        }
        // Left
        if (leftOuter) {
          myDrawLine([y1, x1 - myCropMarkOffset, y1, x1 - (myCropMarkOffset + myCropMarkLength)], myCropMarkWidth, regColor, noneSwatch, cropLayer);
          myDrawLine([y2, x1 - myCropMarkOffset, y2, x1 - (myCropMarkOffset + myCropMarkLength)], myCropMarkWidth, regColor, noneSwatch, cropLayer);
          myDrawLine([y1 - bleedOffset, x1 - myCropMarkOffset, y1 - bleedOffset, x1 - (myCropMarkOffset + myCropMarkLength)], myCropMarkWidth, regColor, noneSwatch, cropLayer);
          myDrawLine([y2 + bleedOffset, x1 - myCropMarkOffset, y2 + bleedOffset, x1 - (myCropMarkOffset + myCropMarkLength)], myCropMarkWidth, regColor, noneSwatch, cropLayer);
        }
        // Right
        if (rightOuter) {
          myDrawLine([y1, x2 + myCropMarkOffset, y1, x2 + myCropMarkOffset + myCropMarkLength], myCropMarkWidth, regColor, noneSwatch, cropLayer);
          myDrawLine([y2, x2 + myCropMarkOffset, y2, x2 + myCropMarkOffset + myCropMarkLength], myCropMarkWidth, regColor, noneSwatch, cropLayer);
          myDrawLine([y1 - bleedOffset, x2 + myCropMarkOffset, y1 - bleedOffset, x2 + myCropMarkOffset + myCropMarkLength], myCropMarkWidth, regColor, noneSwatch, cropLayer);
          myDrawLine([y2 + bleedOffset, x2 + myCropMarkOffset, y2 + bleedOffset, x2 + myCropMarkOffset + myCropMarkLength], myCropMarkWidth, regColor, noneSwatch, cropLayer);
        }
      } else if (myRange != 0 && i == 0) { // Entire selection bounding box (only once)
        var myX1 = x1, myY1 = y1, myX2 = x2, myY2 = y2;
        for (var j = 1; j < selection.length; j++) {
          var bb = selection[j].visibleBounds;
          if (bb[0] < myY1) myY1 = bb[0];
          if (bb[1] < myX1) myX1 = bb[1];
          if (bb[2] > myY2) myY2 = bb[2];
          if (bb[3] > myX2) myX2 = bb[3];
        }
        myDrawCropMarks(myX1, myY1, myX2, myY2, myCropMarkLength, myCropMarkOffset, myCropMarkWidth, regColor, noneSwatch, cropLayer);
        myDrawCropMarks(myX1 - bleedOffset, myY1 - bleedOffset, myX2 + bleedOffset, myY2 + bleedOffset, myCropMarkLength, myCropMarkOffset, myCropMarkWidth, regColor, noneSwatch, cropLayer);
        myDrawCropMarks(myX1 + bleedOffset, myY1 + bleedOffset, myX2 - bleedOffset, myY2 - bleedOffset, myCropMarkLength, myCropMarkOffset, myCropMarkWidth, regColor, noneSwatch, cropLayer);
      }
    }
    doc.viewPreferences.rulerOrigin = oldRulerOrigin;
    doc.viewPreferences.horizontalMeasurementUnits = oldXUnits;
    doc.viewPreferences.verticalMeasurementUnits = oldYUnits;
  }
  function myDrawLine(bounds, strokeWeight, color, swatch, layer) {
    var line = app.activeWindow.activePage.graphicLines.add({
      geometricBounds: bounds,
      strokeWeight: strokeWeight,
      strokeColor: color,
      itemLayer: layer
    });
    return line;
  }
  function myDrawCropMarks(x1, y1, x2, y2, len, offset, width, color, swatch, layer) {
    // Top
    myDrawLine([y1 - offset, x1, y1 - (offset + len), x1], width, color, swatch, layer);
    myDrawLine([y1 - offset, x2, y1 - (offset + len), x2], width, color, swatch, layer);
    // Bottom
    myDrawLine([y2 + offset, x1, y2 + offset + len, x1], width, color, swatch, layer);
    myDrawLine([y2 + offset, x2, y2 + offset + len, x2], width, color, swatch, layer);
    // Left
    myDrawLine([y1, x1 - offset, y1, x1 - (offset + len)], width, color, swatch, layer);
    myDrawLine([y2, x1 - offset, y2, x1 - (offset + len)], width, color, swatch, layer);
    // Right
    myDrawLine([y1, x2 + offset, y1, x2 + offset + len], width, color, swatch, layer);
    myDrawLine([y2, x2 + offset, y2, x2 + offset + len], width, color, swatch, layer);
  }
  function isEdgeOuter(bounds, allBounds, edge, tolerance) {
    var y1 = bounds[0], x1 = bounds[1], y2 = bounds[2], x2 = bounds[3];
    for (var i = 0; i < allBounds.length; i++) {
      var b = allBounds[i];
      if (edge === "top" && Math.abs(y1 - b[2]) < tolerance && x1 < b[3] && x2 > b[1]) return false;
      if (edge === "bottom" && Math.abs(y2 - b[0]) < tolerance && x1 < b[3] && x2 > b[1]) return false;
      if (edge === "left" && Math.abs(x1 - b[3]) < tolerance && y1 < b[2] && y2 > b[0]) return false;
      if (edge === "right" && Math.abs(x2 - b[1]) < tolerance && y1 < b[2] && y2 > b[0]) return false;
    }
    return true;
  }

} // <-- Properly close placeImagesOnDocumentPages

placeImagesOnDocumentPages();