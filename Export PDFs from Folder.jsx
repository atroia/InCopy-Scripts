/* --------------------------------------
Export PDFs from Folder
by Aaron Troia (@atroia)
Modified Date: 05/06/25

Description: 
Allow Export of both Layout and Story views in InCopy 
from folder.

Change Log:
v1.0.1 - added iteration into myTotalDocs so that 
exportGallery() and exportLayout() would only run on one
open file and once run, close the files. Added try block
to main()
v1.0.2 - added userInteration preferences to supress 
"Missing Font" dialog
-------------------------------------- */

var scptName = "Export PDFs from Folder";
var scptVersion = "v1.0.1";
var thePath;
var FULL;

function main() {
  try{
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT; // surpress "Missing Fonts" dialog
    openFiles();
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL; // Allow all dialogs again.
  } catch (e) {
    alert(e.line);
  }
}

function openFiles() {
  var myDocs = Folder.selectDialog("Select a folder with InCopy assignments");
  if (!myDocs) exit(0);
  var myTotalDocs = myDocs.getFiles("*.icma");
  for (var i = myTotalDocs.length - 1; i >= 0; i--) {
    app.open(myTotalDocs[i]);
    exportGalley();
    exportLayout();
    app.activeDocument.close();
  }
}

/* =============================== */
/* ====  PDF Export - Layout  ==== */
/* =============================== */

function exportLayout() {
  layoutPrefs();
  exportPath();

  // Here you can set the suffix at the end of the name
  FULL = thePath + "_LAYOUT.pdf"; // Print PDF

  app.activeDocument.exportFile(ExportFormat.PDF_TYPE, new File(FULL));
}


/* =============================== */
/* ====  PDF Export - Galley  ==== */
/* =============================== */

function exportGalley() {
  galleryPrefs();
  exportPath();

  // Here you can set the suffix at the end of the name
  FULL = thePath + "_STORY.pdf"; // Print PDF

  app.activeDocument.exportFile(ExportFormat.PDF_TYPE, new File(FULL));
}

/* ======================= */
/* ====  Export Path  ==== */
/* ======================= */

function exportPath() {
  thePath = String(app.activeDocument.fullName).replace(/\..+$/, "") + ".pdf";
  thePath = String(new File(thePath)); //.saveDlg()
  thePath = thePath.replace(/\.pdf$/, "");
}


/* ======================= */
/* ====  PREFERENCES  ==== */
/* ======================= */

function layoutPrefs(){
  // Switch to view to export
  app.activeWindow.viewTab = ViewTabs.LAYOUT_VIEW;

  // Acrobat Compatability
  app.layoutPDFExportPreferences.acrobatCompatibility =
    AcrobatCompatibility.ACROBAT_5;
  // pageRange = ; // page spread
  app.layoutPDFExportPreferences.exportReaderSpreads = false;

  //fonts
  app.layoutPDFExportPreferences.subsetFontsBelow = 100; // 0-100

  //options
  app.layoutPDFExportPreferences.includeNotes = true;
  app.layoutPDFExportPreferences.pageInformationMarks = true;
  app.layoutPDFExportPreferences.optimizePDF = false;
  app.layoutPDFExportPreferences.generateThumbnails = false;
  app.layoutPDFExportPreferences.interactiveElementsOption =
    InteractiveElementsOptions.DO_NOT_INCLUDE; // InteractiveElementsOptions.APPEARANCE_ONLY

  app.layoutPDFExportPreferences.viewPDF = false;
}

function galleryPrefs() {
  // Switch to Tab to export
  app.activeWindow.viewTab = ViewTabs.STORY_VIEW; // ViewTabs.GALLEY_VIEW

  // Acrobat Compatability
  (app.galleyPDFExportPreferences.acrobatCompatibility =
    AcrobatCompatibility.ACROBAT_5),
    //fonts
    (app.galleyPDFExportPreferences.subsetFontsBelow = 100); // 0-100
  app.galleyPDFExportPreferences.appliedFont = "Adobe Garamond Pro";
  app.galleyPDFExportPreferences.fontStyle = "Regular";
  app.galleyPDFExportPreferences.pointSize = 10;
  app.galleyPDFExportPreferences.leading = 100;

  //options
  app.galleyPDFExportPreferences.includePageInfo = true;
  app.galleyPDFExportPreferences.includeStoryInfo = true;
  app.galleyPDFExportPreferences.includeAllStories = true; //If true; exports all stories. If false; exports only the current story.
  app.galleyPDFExportPreferences.includeParagraphStyles = false;
  app.galleyPDFExportPreferences.includeNotes = true;
  app.galleyPDFExportPreferences.includeAllNotes = true; // If true; exports both invisible and visible notes. If false; exports only visible notes.
  app.galleyPDFExportPreferences.showNotesBackgroundsInColor = true;
  app.galleyPDFExportPreferences.includeTrackedChanges = true;
  app.galleyPDFExportPreferences.showTrackedChangesBackgroundsInColor = true;
  app.galleyPDFExportPreferences.includeAccurateLineEndings = true;
  // lineRange = ; // The lines to export. . Can return = LineRange enumerator or String.
  app.galleyPDFExportPreferences.includeLineNumbers = false;
  // storyColumns = ; // (Fill Page Option) The width of columns to export. Valid only when include accurate line endings is true.

  app.galleyPDFExportPreferences.viewPDF = false;
}

main();
