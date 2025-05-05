/* --------------------------------------
Export InCopy PDFs
by Aaron Troia (@atroia)
Modified Date: 05/05/25

Description: 
Allow Export of both Layout and Story views in InCopy

-------------------------------------- */

var scptName = "Export InCopy PDFs"
var scptVersion = "v1.0.0"
var d = app.activeDocument; 


main();


function main () {
     try {
    if (app.documents.length == 0) {
      alert("No documents are open.");
    } else {
         exportGalley();
         exportLayout();
      }
    } catch (e) {
    alert(e.line);
  }
}


/* =============================== */
/* ====  PDF Export - Layout  ==== */
/* =============================== */

function exportLayout() {
   // Switch to view to export
   app.activeWindow.viewTab = ViewTabs.LAYOUT_VIEW;

   // Acrobat Compatability
    app.layoutPDFExportPreferences.acrobatCompatibility = AcrobatCompatibility.ACROBAT_5;
         // pageRange = ; // page spread
    app.layoutPDFExportPreferences.exportReaderSpreads = false;
         
         //fonts
    app.layoutPDFExportPreferences.subsetFontsBelow = 100; // 0-100

         //options
    app.layoutPDFExportPreferences.includeNotes = true;
    app.layoutPDFExportPreferences.pageInformationMarks = true;
    app.layoutPDFExportPreferences.optimizePDF = false;
    app.layoutPDFExportPreferences.generateThumbnails = false;
    app.layoutPDFExportPreferences.interactiveElementsOption = InteractiveElementsOptions.DO_NOT_INCLUDE; // InteractiveElementsOptions.APPEARANCE_ONLY

    app.layoutPDFExportPreferences.viewPDF = false;


  if (d.saved) {
    var thePath = String(d.fullName).replace(/\..+$/, "") + ".pdf";
    thePath = String(new File(thePath)); //.saveDlg()
    if (thePath == null) {
      alert ("You pressed Cancel!");
      exit();
    }
  } else {
    thePath = String(new File()); //.saveDlg()
    if (thePath == null){
      alert ("You pressed Cancel!");
      exit();
    }
  }


  thePath = thePath.replace(/\.pdf$/, "");
  thePath2 = thePath.replace(/(\d+b|\.pdf$)/, "");
  // Here you can set the suffix at the end of the name
  FULL = thePath + "_LAYOUT.pdf"; // Print PDF



d.exportFile(ExportFormat.PDF_TYPE, new File(FULL));
}


/* =============================== */
/* ====  PDF Export - Galley  ==== */
/* =============================== */

function exportGalley() {
   // Switch to Tab to export
   app.activeWindow.viewTab = ViewTabs.STORY_VIEW; // ViewTabs.GALLEY_VIEW

   // Acrobat Compatability
   app.galleyPDFExportPreferences.acrobatCompatibility = AcrobatCompatibility.ACROBAT_5,

   //fonts
   app.galleyPDFExportPreferences.subsetFontsBelow = 100;     // 0-100
   app.galleyPDFExportPreferences.appliedFont = "Adobe Garamond Pro";
   app.galleyPDFExportPreferences.fontStyle = "Regular";
   app.galleyPDFExportPreferences.pointSize = 10;
   app.galleyPDFExportPreferences.leading = 100;

   //options
   app.galleyPDFExportPreferences.includePageInfo = true;
   app.galleyPDFExportPreferences.includeStoryInfo = true;
      app.galleyPDFExportPreferences.includeAllStories = true;   //If true; exports all stories. If false; exports only the current story.
   app.galleyPDFExportPreferences.includeParagraphStyles = false;
   app.galleyPDFExportPreferences.includeNotes = true;
   app.galleyPDFExportPreferences.includeAllNotes = true;  // If true; exports both invisible and visible notes. If false; exports only visible notes.
   app.galleyPDFExportPreferences.showNotesBackgroundsInColor = true;
   app.galleyPDFExportPreferences.includeTrackedChanges = true;
      app.galleyPDFExportPreferences.showTrackedChangesBackgroundsInColor = true;
   app.galleyPDFExportPreferences.includeAccurateLineEndings = true;
      // lineRange = ; // The lines to export. . Can return = LineRange enumerator or String.
      app.galleyPDFExportPreferences.includeLineNumbers = false;
      // storyColumns = ; // (Fill Page Option) The width of columns to export. Valid only when include accurate line endings is true.

   app.galleyPDFExportPreferences.viewPDF = false;



   if (d.saved) {
      var thePath = String(d.fullName).replace(/\..+$/, "") + ".pdf";
      thePath = String(new File(thePath)); //.saveDlg()
      if (thePath == null) {
      alert ("You pressed Cancel!");
      exit();
      }
   } else {
      thePath = String(new File()); //.saveDlg()
      if (thePath == null){
      alert ("You pressed Cancel!");
      exit();
      }
   }


   thePath = thePath.replace(/\.pdf$/, "");
   thePath2 = thePath.replace(/(\d+b|\.pdf$)/, "");
   // Here you can set the suffix at the end of the name
   FULL = thePath + "_STORY.pdf"; // Print PDF


   d.exportFile(ExportFormat.PDF_TYPE, new File(FULL));
}


