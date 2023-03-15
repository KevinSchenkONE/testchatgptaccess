function onOpen() {
    var ui = SpreadsheetApp.getUi();
    var uemail = Session.getEffectiveUser().getEmail();
    if (uemail == "david.typanski@one-line.com"
      || uemail == "john.stone@one-line.com"
      || uemail == "jamie.lambson@one-line.com"
      || uemail == "adam.matthijs@one-line.com") {
      ui.createMenu('NROF')
        .addItem("Import Data", "importRawData")
        .addItem("Distribute Data", "splitRawData")
        .addItem("Consolidate Sheets to temp", "consolidtesheetstotemp")
        .addItem("Consolidate temp to all", "consolidatetemptoall")
        .addItem('Consolidate All (was combine sheets)', 'consolidateall')
        .addItem("Archive", "archiveold")
        .addItem("Do Everything", "doitall")
        .addItem("Export All NROF Data (Parms col 22)", "exportAllNROFData")
        .addItem("Send email containing sheet records to Adam, John and Dave", "sendEmail")
        .addToUi();
    }
  };
  var gThisSSID = "1ADpEOZn33Vu1umjndiTKADJ9BZun11f6-ehjCNkvdTE";
  //var gSS = SpreadsheetApp.openById(gThisSSID);
  var gArchID = "1zeF1V_ZNpRPUGEodlgJyAmio-QxxNdD62TOtR5zx_Rc";
  var gFeedID = "1gfbxOhuQYgsaGByHesRv2h2CzYHuIcvt4Tzgpo0cKUc";
  
  //var gArchSS=SpreadsheetApp.getActiveSpreadsheet();
  
  
  function doitall() {
  
    var dstart = new Date();
    logaction("Status", "Do All Start=" + dstart, "doitall");
    logaction("Status", "consolidateall Start", "doitall");
    consolidateall();
    //return;
    logaction("Status", "ImportRawData Start", "doitall");
    importRawData();
    logaction("Status", "splitRawData Start", "doitall");
    splitRawData();
    logaction("Status", "archiveold skipped", "doitall");
    //archiveold();
    // logaction("Status", "Export All NROF Data", "doitall");
    //exportAllNROFData();
    logaction("Status", "Do All End " + (((new Date()) - dstart) / 1000) + " seconds total", "doitall");
  };
  function onEdit(e) {
    if (!e) return;
    var aName = e.range.getSheet().getName();
    if (aName == "DOC - Documentation"
      || aName == "MKTG - Marketing"
      || aName == "NROF Team - Audit"
      || aName == "SALES-CRO"
      || aName == "SALES-SRO"
      || aName == "SALES-CAN"
      || aName == "SALES-JPN"
      || aName == "SALES-WRO"
      || aName == "SALES-ERO"
      || aName == "SALES-OTH"
      || aName == "CSVC - Customer Service") {
      var erow = e.range.getRow();
      var ecol = e.range.getColumn();
      if (ecol == 14) {
        var oldinfo = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(aName).getRange(erow, 54).getValue();
        if (oldinfo == "" || oldinfo == null) {
          SpreadsheetApp.getActiveSpreadsheet().getSheetByName(aName).getRange(erow, 54).setValue(new Date());
        }
  
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(aName).getRange(erow, 55).setValue(Session.getEffectiveUser().getEmail());
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(aName).getRange(erow, 56).setValue(new Date());
      }
      else if (ecol == 18) {
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(aName).getRange(erow, 48).setValue(Session.getEffectiveUser().getEmail());
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(aName).getRange(erow, 49).setValue(new Date());
      }
      else if (ecol == 23) {
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(aName).getRange(erow, 50).setValue(Session.getEffectiveUser().getEmail());
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(aName).getRange(erow, 51).setValue(new Date());
      }
      else if (ecol == 24) {
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(aName).getRange(erow, 52).setValue(Session.getEffectiveUser().getEmail());
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(aName).getRange(erow, 53).setValue(new Date());
      }
    }
  };
  function splitRawData() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    //var ss = gSS;
    var dstart = new Date();
    var lastlogline = logaction("Status", "splitRawData Start=" + dstart, "splitRawData");
    var rawdataSheet = ss.getSheetByName("Raw Data");
    var parmsheet = ss.getSheetByName('Parameters');
    console.log("Status", "splitRawData - Sorting Raw");
    var namecol = parmsheet.getRange(3, 6).getValue();
    var nameindex = namecol - 1;
    var colsinsource = parmsheet.getRange(1, 6).getValue();
    var dataRange = rawdataSheet.getRange(2, 1, rawdataSheet.getMaxRows() - 1, colsinsource);
    dataRange.sort([
      { column: 47, ascending: true },
      { column: 3, ascending: true },
      { column: 11, ascending: true }
    ]);
    console.log("splitRawData - Clearing targets");
    clearIndividualSheets();
    console.log("splitRawData - Reading Raw");
    var rawdata = dataRange.getValues();
    var startR = 0;
    var slEnd = 0;
    var curSheet = rawdata[0][nameindex];
    var prevSheet = curSheet;
    console.log("Writing Totals");
    writetots(rawdata, 3);
    console.log("splitRawData - Looping");
    for (var rQ = 1; rQ < rawdata.length; rQ++) {
      curSheet = rawdata[rQ][nameindex];
      if (!(curSheet == prevSheet && rQ != rawdata.length - 1)) {
        var wSheet = ss.getSheetByName(prevSheet);
        slEnd = rQ - startR;
        if (rQ == rawdata.length - 1) {
          slEnd++
        }
        console.log("splitRawData - Working on " + wSheet.getName());
        bigWrite(slice2d(rawdata, startR, 0, slEnd, colsinsource), parmsheet.getRange(1, 18).getValue(), wSheet.getName(), wSheet.getRange(2, 1, slEnd, colsinsource).getA1Notation());
        // wSheet.sort(3);
        //logaction("Status", "splitRawData start fix columns for: " + wSheet.getName(), "splitRawData", lastlogline++);
        fixthesecolumns(wSheet.getName());
        //logaction("Status", "splitRawData end fix columns for: " + wSheet.getName() + " start sort", "splitRawData", lastlogline++);
        // sortThisSheet(wSheet.getName());
        startR = rQ;
        prevSheet = curSheet;
      }
    }//return Math.floor((utc2 - utc1) / _MS_PER_DAY)
    //logaction("Status", "splitRawData End " + (((new Date()) - dstart) / 1000) + " seconds total", "splitRawData", lastlogline++);
    console.log(deleteSheetData("Raw Data"));
    console.log("splitRawData End " + (((new Date()) - dstart) / (1000)) + " seconds total");
    logaction("Status", "splitRawData End " + (((new Date()) - dstart) / (1000)) + " seconds total", "splitRawData", lastlogline++);
    SpreadsheetApp.flush();
  };
  
  function GetLastRowNumber(useSheetName, colToCount, optSSID) {
    if ("" == optSSID || null == optSSID) {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
    }
    else {
      var ss = SpreadsheetApp.openById(optSSID);
    }
    var sheet = ss.getSheetByName(useSheetName);
    var lRow = sheet.getLastRow();
    lRow++;
    var range = sheet.getRange(1, 1, lRow, colToCount);
    var values = range.getValues(); // get all data in one call
    var ct = 0;
    var lastdatarow = 0;
    for (ct = 0; ct < values.length; ct++) {
      if (values[ct][colToCount - 1] != "") {
        lastdatarow = ct + 1
      }
    }
    return lastdatarow;
  };
  
  function importRawData() {
    var dstart = new Date();
    var tss = SpreadsheetApp.getActiveSpreadsheet();
    //var tss = gSS;
    var ts = tss.getSheetByName("Raw Data");
    var lastCol = tss.getSheetByName('Parameters').getRange(1, 6).getValue();
    var lastIndex = lastCol - 1;
    var splitCol = tss.getSheetByName('Parameters').getRange(3, 6).getValue();
    var rawss = SpreadsheetApp.openById(tss.getSheetByName('Parameters').getRange(2, 18).getValue());
    var rdSheet = rawss.getSheetByName('All NROF Bookings');
    var lastlogline = logaction("Status", "ImportRawData Start=" + dstart, "ImportRawData");
  
    rdSheet.sort(1);
    var datrows = rdSheet.getMaxRows() - 1;
    var SRange = rdSheet.getRange(2, 1, datrows, lastCol);
    var SData = SRange.getValues();
    logaction("Status", "importRawData got values", "importRawData", lastlogline++);
    var archss = SpreadsheetApp.openById(tss.getSheetByName('Parameters').getRange(3, 18).getValue());
    var archLR = GetLastRowNumber("Archive", 1, tss.getSheetByName('Parameters').getRange(3, 18).getValue())
    var archSheet = archss.getSheetByName('Archive');
    var archData = archSheet.getRange(2, 1, archLR - 1, lastCol);
  
  
    // t stands for target
    logaction("Status", "ImportRawData opening ALL NROF Bookings", "ImportRawData", lastlogline++);
    var oSheet = tss.getSheetByName("All NROF Bookings");
    var oLR = GetLastRowNumber("All NROF Bookings", 1);
    var oRange = oSheet.getRange(2, 1, oLR - 1, lastCol);
    var oData = oRange.getValues();
    // o stands for other
  
    var dupeloops = SData.length;
    logaction("Status", "ImportRawData remove dupes start for " + dupeloops + " rows", "ImportRawData", lastlogline++);
    var prevbkg = "me";
    var oDate = new Date();
    var dupesremoved = "NROF dupe check removed ";
    var foundadupe = 0;
    //var archMsg = "The following BKG Numbers are new to the Raw Data feed but found in the archive.  The last values for comments and comment date and commentor have been updated with the values from the archive";
    for (var x = 0; x < oData.length; x++) {
      oData[x][lastIndex] = 0;
    }
    for (var w = dupeloops - 1; w >= 0; w = w - 1)//dupeloops is sdata.length
    {
      if (SData[w][0] == prevbkg) {
        dupesremoved = "" + dupesremoved + " BKG " + SData[w + 1][0] + " line " + (w + 1) + " w=" + w;
        foundadupe++;
        logaction("Status", "dupeloops=" + dupeloops + " w=" + w + " about to splice " + (w + 1), "ImportRawData", lastlogline++);
        SData.splice(w + 1, 1)
      }
      if (foundadupe) {
        logaction("Status", dupesremoved, "ImportRawData", lastlogline++);
      }
      foundadupe = 0;
      SData[w][lastIndex] = 1;
      SData[w][lastIndex - 1] = oDate;
      found = 0;
      for (var j = 0; j < oData.length; j++) // for all the All NROF Bookings Data
      {
        if (SData[w][0] == oData[j][0]) {
          SData[w][lastIndex - 2] = oData[j][lastIndex - 2];
          oData[j][lastIndex - 1] = oDate;
          oData[j][lastIndex] = 1;
          found = 1
          for (var k = splitCol; k <= lastIndex - 3; k++) {
            SData[w][k] = oData[j][k];
          }
        }
      }
      if (!found) {
        for (var archrow = archData.length - 1; archrow >= 0; archrow = archrow - 1) {
          if (archData[archrow][0] == SData[w][0] && !found) {
            SData[w][lastIndex - 2] = archData[j][lastIndex - 2];
            SData[w][13] = archData[j][13];
            SData[w][17] = archData[j][17];
            SData[w][22] = archData[j][22];
            SData[w][23] = archData[j][23];
            found = 1;
            for (var k = splitCol; k <= lastIndex - 3; k++) {
              SData[w][k] = archData[archrow][k];
            }
          }
        }
        if (!found) {
          SData[w][lastIndex - 2] = oDate;
        }
      }
      prevbkg = SData[w][0];
    }
    //  logaction("Status","ImportRawData Starting Loop of "+ SData.length +" rows "+oData.length +" times for a total of "+SData.length*oData.length+" actions","ImportRawData",lastlogline++);
    logaction("Status", "ImportRawData End synching All NROF and Raw data, clearing Raw Data Sheet", "ImportRawData", lastlogline++);
    //range.sort({column: 2, ascending: false});
    if (ts.getMaxRows() > 3) {
      ts.deleteRows(3, ts.getMaxRows() - 2);
    }
    ts.getRange(2, 1, 1, lastCol).clearContent();
    for (var dl = 0; dl < SData.length; dl++) {
      if (SData[dl][15] != "") {
        SData[dl][15] = new Date(SData[dl][15])
      }
      else {
        var x = new Date();
        SData[dl][15] = new Date(x.setHours(0, 0, 0, 0))
      }
  
      //this is the date added to file
      // SData[dl][52]=new Date(SData[dl][52]);//Internal Marketing Comments When
      // SData[dl][53]=new Date(SData[dl][53]);//Audit First Comment Date
      // SData[dl][55]=new Date(SData[dl][55]);//Audit Last Comment Date
      SData[dl][56] = new Date(SData[dl][lastIndex - 2]); //first sighted
      SData[dl][57] = new Date(SData[dl][lastIndex - 1]);//last sighted
    }
  
    logaction("Status", "ImportRawData data cleared, NOT writiing totals to analysis", "ImportRawData", lastlogline++);
    writetots(SData, 2);
    logaction("Status", "ImportRawData totals written, start big write of Raw Data", "ImportRawData", lastlogline++);
    bigWrite(SData, tss.getSheetByName('Parameters').getRange(1, 18).getValue(), ts.getName(), ts.getRange(2, 1, SData.length, lastCol).getA1Notation());
    logaction("Status", "End bigWrite, start fix date columns", "ImportRawData", lastlogline++);
    fixthesecolumns("Raw Data");
    logaction("Status", "End fix columns for Raw Data", "ImportRawData", lastlogline++);
    var dataRange = ts.getRange(2, 1, ts.getMaxRows() - 1, lastCol);
  
    dataRange.sort([
      { column: 47, ascending: true },
      { column: 3, ascending: true },
      { column: 11, ascending: true }
    ]);
  
    logaction("Status", "ImportRawData Clearing All NROF Bookings", "ImportRawData", lastlogline++);
    oSheet.getRange(2, splitCol + 1, oData.length, 12).clearContent()
    logaction("Status", "ImportRawData All NROF Bookings cleared, Start update to All NROF Bookings", "ImportRawData", lastlogline++);
    bigWrite(slice2d(oData, 0, splitCol, oData.length, 12), tss.getSheetByName('Parameters').getRange(1, 18).getValue(), oSheet.getName(), oSheet.getRange(2, splitCol + 1, oData.length, 12).getA1Notation());
    logaction("Status", "Start fix columns for: " + oSheet.getName(), "ImportRawData", lastlogline++);
    fixthesecolumns(oSheet.getName());
    logaction("Status", "End fix columns for: " + oSheet.getName(), "ImportRawData", lastlogline++);
    logaction("Status", "End fix columns for: " + oSheet.getName() + ", ImportRawData End " + (((new Date()) - dstart) / 1000) + " seconds total", "ImportRawData", lastlogline++);
    SpreadsheetApp.flush();
  };
  function logaction(mType, message, func, optlastline) {
    var parmrow = [];
    var theparms = [];
    var line = 0;
    var date = new Date();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Automation Log");
    //var sheet = gSS.getSheetByName("Automation Log");
    if (0 == optlastline || null == optlastline) {
      line = GetLastRowNumber("Automation Log", 1);
    }
    else {
      line = optlastline;
    }
    line++;
    //theparms.push(date);
    sheet.getRange(line, 1, 1, 1).setValue(date);
    //theparms.push(mType);
    theparms.push(message);
    theparms.push(func);
    parmrow.push(theparms);
    sheet.getRange(line, 3, 1, 2).setValues(parmrow);
    return line;
  };
  function slice2d(array, rowIndexBase0, colIndex, numRows, numCols) {
    var result = [];
    for (var i = rowIndexBase0; i < (rowIndexBase0 + numRows); i++) {
      result.push(array[i].slice(colIndex, colIndex + numCols));
    }
    return result;
  };
  function clearIndividualSheets() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    //var ss = gSS;
    var parmSheet = ss.getSheetByName('Parameters');
    var splitSheetLastLine = GetLastRowNumber('Parameters', 3);
    var splitSheetNames = parmSheet.getRange(2, 3, splitSheetLastLine - 1, 1).getValues();
    var lr = 0;
    for (var i = 0; i < splitSheetNames.length; i++) {
      var curSheet = ss.getSheetByName(splitSheetNames[i]);
      lr = curSheet.getMaxRows();
      curSheet.setFrozenRows(1);
      curSheet.getRange(2, 1, 1, curSheet.getMaxColumns()).clearContent();
      if (lr > 2) {
        curSheet.deleteRows(3, lr - 2);
      }
    }
  };
  //MailApp.sendEmail(emailaddy, subject, message);
  //message="Dear " + refdata[row][1] + ", \n\n" + refdata[row][0] + " has been assigned to
  function consolidtesheetstotemp() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    //var ss = gSS;
    var dstart = new Date();
    var lastlogline = logaction("Status", "consolidtesheetstotemp Start=" + dstart, "consolidtesheetstotemp");
    var parmSheet = ss.getSheetByName('Parameters');
    var splitSheetLastLine = GetLastRowNumber('Parameters', 3);
    var splitSheetNames = parmSheet.getRange(2, 3, splitSheetLastLine - 1, 1).getValues();
    var tarSheet = ss.getSheetByName('Temp NROF Bookings');
    var colsinsource = parmSheet.getRange(1, 6).getValue();
    var ssLR = 0;
    var lr = tarSheet.getMaxRows();
    logaction("Status", "consolidtesheetstotemp deleting " + (lr - 2) + " rows from temp", "consolidtesheetstotemp", lastlogline++);
    if (lr > 2) {
      tarSheet.deleteRows(3, lr - 2);
    }
    tarSheet.getRange(2, 1, 1, tarSheet.getMaxColumns()).clearContent();
    var vals = [];
    var emailaddy = "";
    var subject = "";
    var message = "";
    logaction("Status", "consolidtesheetstotemp Starting Looping", "consolidtesheetstotemp", lastlogline++);
    for (i = 0; i < splitSheetNames.length; i++) {
      logaction("Status", "consolidtesheetstotemp working on " + splitSheetNames[i], "consolidtesheetstotemp", lastlogline++);
      var tvals = [];
      ssLR = GetLastRowNumber(splitSheetNames[i], 1);
      if (ssLR >= 2) {
        tvals = getdataarray(splitSheetNames[i], colsinsource)
        for (var j = 0; j < tvals.length; j++) {
          if (tvals[j][0] == "BKG Number") {
            logaction("Status", "consolidtesheetstotemp **Data Error** sending emails NROF Sheet " + splitSheetNames[i] + " row " + (j + 2) + " has the first column containing BKG Number", "consolidtesheetstotemp", lastlogline++);
            emailaddy = "david.typanski@one-line.com";
            subject = "NROF **ERROR** - Action Required";
            message = "NROF Sheet " + splitSheetNames[i] + " row " + (j + 2) + " has the first column containing BKG Number and it should not.\n\n"
            message = message + "As a result, there is the POSSIBILTY that one record of data was not consolidated because it is wrongly in the header row. "
            message = message + "Here are the details and how to fix if needed.\n\n"
            message = message + "Someone likely sorted the sheet incorrectly. This is checked during consolidtesheetstotemp but"
            message = message + " will not impact the upload to Cognos because the row was skipped.  Given this error it is likely that"
            message = message + " the first row on this sheet no longer contains the header and instead contains some actual data."
            message = message + "\n\nThose data in the first row are not being consolidated."
            message = message + " Please remove record " + (j + 2) + " from " + splitSheetNames[i] + " and then:"
            message = message + "\n\n1. Insert a header row in row 1 if needed."
            message = message + "\n2. Then run Consolodate All and then Export All NROF Data again."
            message = message + "\n3. Once this is done you will need to get someone to run the Cognos import again."
            message = message + "\n4. After Cognos has produced a new export file for NROF to ingest, run Do Everything."
            message = message + "\n\nThis has been logged on line " + lastlogline + " in the tab Automation Log."
            message = message + "\n\nThis message is being sent to david.typanski@one-line.com, adam.matthijs@one-line.com, "
            message = message + "john.stone@one-line.com, and soumya.joseph@one-line.com."
            MailApp.sendEmail(emailaddy, subject, message);
            emailaddy = "adam.matthijs@one-line.com";
            MailApp.sendEmail(emailaddy, subject, message);
            emailaddy = "john.stone@one-line.com";
            MailApp.sendEmail(emailaddy, subject, message);
            emailaddy = "soumya.joseph@one-line.com";
            MailApp.sendEmail(emailaddy, subject, message);
          }
          else {
            vals.push(tvals[j]);
          }
        }
      }
    }
    logaction("Status", "consolidtesheetstotemp writing tot", "consolidtesheetstotemp", lastlogline++);
    writetots(vals, 4);
    logaction("Status", "consolidtesheetstotemp Start bigWrite", "consolidtesheetstotemp", lastlogline++);
    bigWrite(vals, parmSheet.getRange(1, 18).getValue(), tarSheet.getName(), tarSheet.getRange(2, 1, vals.length, colsinsource).getA1Notation());
    SpreadsheetApp.flush();
    logaction("Status", "End big write, start fix columns for: " + tarSheet.getName(), "consolidtesheetstotemp", lastlogline++)
    fixthesecolumns(tarSheet.getName());
    logaction("Status", "Col fix ended, consolidtesheetstotemp end " + (((new Date()) - dstart) / 1000) + " seconds total", "consolidtesheetstotemp", lastlogline++)
  
  };
  function consolidateall() {
    consolidtesheetstotemp();
    consolidatetemptoall();
    exportAllNROFData();
  };
  function getdataarray(Sheetname, cols) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    //var ss = gSS;
    var s = ss.getSheetByName(Sheetname);
    var lr = GetLastRowNumber(Sheetname, 1);
    var result = s.getRange(2, 1, lr - 1, cols).getValues();
    return result;
  };
  //message="Dear " + refdata[row][1] + ", \n\n" + refdata[row][0] + " has been assigned to
  function consolidatetemptoall() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    //var ss = gSS;
    var dstart = new Date();
    var lastlogline = logaction("Status", "consolidatetemptoall Start=" + dstart, "consolidatetemptoall");
    var parmSheet = ss.getSheetByName('Parameters');
    var colsinsource = parmSheet.getRange(1, 6).getValue();
    var colsintarget = colsinsource;
    var tarSheet = ss.getSheetByName('All NROF Bookings');
    var tarSRows = GetLastRowNumber('All NROF Bookings', 1);
    var srcSheet = ss.getSheetByName('Temp NROF Bookings');
    var srcRows = GetLastRowNumber('Temp NROF Bookings', 1);
    var zips = 0;
    var ones = 0;
    var foundRow = 0;
    // var Dte = new Date();
    var tempRow = [];
    if (srcRows == 1) {
      SpreadsheetApp.getUi().alert("No data in source file (Temp)");
      return;
    }
    logaction("Status", "consolidatetemptoall getting values", "consolidatetemptoall", lastlogline++);
    var tarvals = tarSheet.getRange(2, 1, tarSRows - 1, colsintarget).getValues();
    var srcvals = srcSheet.getRange(2, 1, srcRows - 1, colsinsource).getValues();
    logaction("Status", "consolidatetemptoall loop start", "consolidatetemptoall", lastlogline++);
    for (var i = 0; i < srcvals.length; i++) {
      foundRow = 0;
      for (var j = 0; j < tarvals.length; j++) {
        if (tarvals[j][0] == srcvals[i][0]) {
          foundRow = 1;
          tarRow = j;
          for (var k = 0; k < colsinsource; k++) {
            tarvals[j][k] = srcvals[i][k];
          }
        }
      }
      if (foundRow == 0) {
        tempRow = [];
        for (var k = 0; k < colsinsource; k++) {
          tempRow.push(srcvals[i][k]);
        }
        tarvals.push(tempRow);
        tarSRows++
      }
    }
    logaction("Status", "consolidatetemptoall done looping, now clearing target", "consolidatetemptoall", lastlogline++);
    if (tarSheet.getMaxRows() >= 3) { tarSheet.deleteRows(3, tarSheet.getMaxRows() - 2) };
    tarSheet.getRange(2, 1, 1, tarSheet.getMaxColumns()).clearContent();
    logaction("Status", "consolidatetemptoall write tots", "consolidatetemptoall", lastlogline++);
    writetots(tarvals, 5);
    logaction("Status", "Start bigWrite", "consolidatetemptoall", lastlogline++);
    bigWrite(tarvals, parmSheet.getRange(1, 18).getValue(), tarSheet.getName(), tarSheet.getRange(2, 1, tarvals.length, colsintarget).getA1Notation());
    logaction("Status", "consolidatetemptoall start fix columns for: " + tarSheet.getName(), "consolidatetemptoall", lastlogline++);
    fixthesecolumns(tarSheet.getName());
    logaction("Status", "consolidatetemptoall end fix columns for: " + tarSheet.getName(), "consolidatetemptoall", lastlogline++);
    for (var c = 0; c < tarvals.length; c++) {
      if (tarvals[c][colsinsource - 1] == 0) {
        zips++
      }
      else if (tarvals[c][colsinsource - 1] == 1) {
        ones++
      }
    }
    //ss.getSheetByName('Analysis').getRange(2, 6).setValue(zips);
    //ss.getSheetByName('Analysis').getRange(2, 7).setValue(ones);
    logaction("Status", deleteSheetData("Temp NROF Bookings"), "consolidatetemptoall", lastlogline++);
    logaction("Status", "consolidatetemptoall End " + (((new Date()) - dstart) / 1000) + " seconds total", "consolidatetemptoall", lastlogline++);
    SpreadsheetApp.flush();
  };
  
  function archiveold() {
    var dstart = new Date();
    var lastlogline = logaction("Status", "Archiveold Start=" + dstart, "Archiveold");
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    //var ss = gSS;
    var parmSheet = ss.getSheetByName('Parameters');
    var colsinsource = parmSheet.getRange(1, 6).getValue();
    //var archID=parmSheet.getRange(3,18).getValue();
    var archID = gArchID;
    var tarss = SpreadsheetApp.openById(archID);
    var lr = GetLastRowNumber("Archive", 1, archID);
    var tarSheet = tarss.getSheetByName('Archive');
  
    logaction("Status", "Sorting and reading archive data:lr=" + lr, "Archiveold", lastlogline++);
    // tarSRows will have the first row with empty cell in the archive.  The target row.
    var srcSheet = ss.getSheetByName('All NROF Bookings');
    srcSheet.sort(colsinsource);
    var srcRows = GetLastRowNumber('All NROF Bookings', 1);
    var srcVals = srcSheet.getRange(2, 1, srcRows - 1, colsinsource).getValues();
    var i = 0;
    logaction("Status", "Archiveold Finding eligible data to archive colsinsource=" + colsinsource + " srcRows=" + srcRows + " srcVals.length=" + srcVals.length, "Archiveold", lastlogline++);
    while (srcVals[i][colsinsource - 1] == 0) {
      i++
    }
    i++
    if (i == 1) {
      logaction("Status", "Archiveold End - !!Nothing to Archive!!", "Archiveold", lastlogline++);
    }
    else {
      logaction("Status", "Archiveold Writing data tarSRows=" + lr + " i-1=" + (i - 1), "Archiveold", lastlogline++);
      tarSheet.getRange(lr + 1, 1, i - 1, colsinsource).setValues(slice2d(srcVals, 0, 0, i - 1, colsinsource));
      logaction("Status", "Archiveold Deleting Source Data", "Archiveold", lastlogline++);
      srcSheet.deleteRows(2, i - 1);
  
    }
  
    logaction("Status", "Archiveold start fix columns for: " + tarSheet.getName(), "archiveold", lastlogline++);
    fixthesecolumns('Archive', archID);
    logaction("Status", "Archiveold end fix columns for: " + tarSheet.getName(), "archiveold", lastlogline++);
    logaction("Status", "Archiveold End " + (((new Date()) - dstart) / 1000) + " seconds total", "Archiveold", lastlogline++);
  };
  function writetots(thearray, thecol) {
    //return;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    //var ss = gSS;
    var splitSheetNames = ss.getSheetByName('Parameters').getRange(2, 3, (GetLastRowNumber('Parameters', 3)) - 1, 2).getValues();
    var tarSheet = ss.getSheetByName('Analysis');
    for (var k = 0; k, k < splitSheetNames.length; k++) {
      tarSheet.getRange(k + 2, thecol, 1, 1).setValue(0);
    }
    for (var i = 0; i < thearray.length; i++) {
      for (var j = 0; j < splitSheetNames.length; j++) {
        if (thearray[i][46] == splitSheetNames[j][0]) {
          splitSheetNames[j][1] = splitSheetNames[j][1] + 1;
        }
      }
    }
    for (var j = 0; k, j < splitSheetNames.length; j++) {
      tarSheet.getRange(j + 2, thecol, 1, 1).setValue(splitSheetNames[j][1]);
    }
  };
  //spreadsheet ID 'spreadsheet object'.getId() var ssid = SpreadsheetApp.getActiveSpreadsheet().getId()
  //data array to be written
  //range A1 notation of the range you want to write the array to, this MUST match the array size or it will fail, its this way so you dont have to write to a clean sheet, and can update part of it
  //sheet name string representing the destination sheet, destination sheet must be on the spreadsheet represented by the spreadsheet id
  function bigWrite(data, ssid, sheetname, range) {
    var sheet = SpreadsheetApp.openById(ssid).getSheetByName(sheetname);
    sheet.getRange(range).clearContent();
    SpreadsheetApp.flush();
    var mess = "Attempting bigWrite";
    var subject = "NROF Bigwwrite Stats";
    var didSucceed = false;
    var fullrange = sheetname + '!' + range;
    var trynum = 1;
    var request = {
      'range': fullrange,
      'majorDimension': 'ROWS',
      'values': data
    }
    var totcatch = 0;
    var tottry = 0
    while (!didSucceed && trynum <= 3) {
      try {
        tottry++;
        Sheets.Spreadsheets.Values.update(request, ssid, fullrange, { 'valueInputOption': 'RAW' })
        didSucceed = true;
        mess = mess + " bigWrite NROF tries " + tottry + " with sheetname=" + sheetname;
  
      }
      catch {
        totcatch++;
        mess = mess + " bigWrite NROF caught " + totcatch + " with sheetname=" + sheetname;
        trynum++
      }
    }
    //MailApp.sendEmail("david.typanski@one-line.com", subject, mess);
    addy = "david.typanski@one-line.com";
    //MailApp.sendEmail("8049867959@mms.att.net", subject, mess);
  };
  
  // function changethedate()
  // {
  //   var ss = SpreadsheetApp.getActiveSpreadsheet();
  //   var sheet = ss.getSheetByName('Raw Data');
  //   var lr=sheet.getLastRow();
  //   var thedata=sheet.getRange(2,16,lr-1,1).getValues();
  //   for (var i=2;i<=lr;i++)
  //   { 
  //   thedata[i-2][0]=new Date(thedata[i-2][0]);
  //   }
  //   logaction("Status","changethedate lr="+lr+" i="+i,"changethedate");
  //   ss.getSheetByName('traw').getRange(2,16,lr-1,1).setValues(thedata);
  
  // }
  function fixdatecol(col, fixSSSheet, numOfCols) {
    var lr = fixSSSheet.getLastRow();
    var thedata = fixSSSheet.getRange(2, col, lr - 1, numOfCols).getValues();
    for (var colcounter = 0; colcounter < numOfCols; colcounter++) {
      for (var i = 2; i <= lr; i++) {
        if (!(thedata[i - 2][colcounter] == "")) {
          thedata[i - 2][colcounter] = new Date(thedata[i - 2][colcounter]).toLocaleDateString();
        }
      }
    }
    fixSSSheet.getRange(2, col, lr - 1, numOfCols).setNumberFormat('M/d/yyyy');
    fixSSSheet.getRange(2, col, lr - 1, numOfCols).setValues(thedata);
  };
  function fixthesecolumns(sname, extID1) {
    // SpreadsheetApp.getUi().alert("her12");
    if (null == extID1) {
      extID1 = gThisSSID;
    }
    var fixSS = SpreadsheetApp.openById(extID1);
    var fixSSSheet = fixSS.getSheetByName(sname);
    fixdatecol(16, fixSSSheet, 1);
    fixdatecol(28, fixSSSheet, 1);
    fixdatecol(41, fixSSSheet, 3);
    //fixdatecol(42, fixSSSheet, 1);
    //fixdatecol(43, fixSSSheet, 1);
    //fixdatecol(57, fixSSSheet, 1);
    //fixdatecol(58, fixSSSheet, 1);
    fixdatecol(49, fixSSSheet, 1);
    fixdatecol(51, fixSSSheet, 1);
    fixdatecol(53, fixSSSheet, 1);
    fixdatecol(54, fixSSSheet, 1);
    fixdatecol(56, fixSSSheet, 3);
  };
  
  function sortThisSheet(sName) {
    var Sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sName);
    var parmheight = 11;
    var parmwidth = 7;
    var parms = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet Parameters").getRange(2, 1, parmheight, parmwidth).getValues();
    var foundrow = 0;
    for (var rowindex = 0; rowindex < parms.length; rowindex++) {
      if (!foundrow) {
        if (parms[rowindex][0] == sName) {
  
        }
      }
    }
  
  };
  
  function sheetexists() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dstart = new Date();
    var lastlogline = logaction("Status", "sheetexists starting at " + dstart, "sheetexists");
    var sheetcollection = ss.getSheets();
    var parmsheet = ss.getSheetByName("Sheet Parameters");
    var sheetarray = parmsheet.getRange(2, 1, GetLastRowNumber(parmsheet.getName(), 1) - 1, 1).getValues();
    //SpreadsheetApp.getUi().alert("sheetarray.length - "+sheetarray.length);
    //Logger.log("sheetarray.length - "+sheetarray.length)
    //SpreadsheetApp.getUi().alert("sheetcollection.length - "+sheetcollection.length);
    //Logger.log("sheetcollection.length - "+sheetcollection.length)
    var msg = "";
    for (var i = 0; i < sheetarray.length; i++) {
      //Logger.log("i="+i);
      var found = 0;
      for (var eachsheet = 0; eachsheet < sheetcollection.length; eachsheet++) {
        if (i == 0) {
          logaction("Status", sheetcollection[eachsheet].getSheetName(), "sheetexists", lastlogline++);
        }
        if (sheetcollection[eachsheet].getSheetName() == sheetarray[i][0]) {
          found = 1;
          //Logger.log("found sheet - "+sheetcollection[eachsheet].getSheetName());
        }
      }
      if (!found) {
        //Logger.log("did not find sheet - "+sheetarray[i][0]);
        if (msg == "") {
          msg = "The NROF system has a Major Error condition. Please notify na.servicedesk@one-line.com.\n\n"
        }
        msg = msg + "Sheet " + sheetarray[i][0] + " was not found.  This is a MAJOR system problem.\n"
        //Logger.log("msg="+msg);
      }
    }
    //Logger.log("msg after looping is=\n"+msg)
    if (msg != "") {
      logaction("Status", "There is a Major Error in checking if required sheets exist. Emails are being sent to the following people. " + msg, "sheetexists", lastlogline++);
      var emailarray = parmsheet.getRange(2, 10, GetLastRowNumber(parmsheet.getName(), 10) - 1).getValues();
      for (var sendmsg = 0; sendmsg < emailarray.length; sendmsg++) {
        //Logger.log("sending msg to="+emailarray[sendmsg]);
        logaction("Status", "Sending email to " + emailarray[sendmsg][0], "sheetexists", lastlogline++);
        MailApp.sendEmail(emailarray[sendmsg][0], "NROF Sheets Missing ERROR", msg);
      }
    }
    logaction("Status", "sheetexists end " + (((new Date()) - dstart) / 1000) + " seconds total", "sheetexists", lastlogline++);
  };
  function sortthisrange(sheettosortstring, trueifascending) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheettosort = ss.getSheetByName(sheettosortstring);
    var parmsheet = ss.getSheetByName("Sheet Parameters");
    var sheetarray = parmsheet.getRange(2, 1, GetLastRowNumber(parmsheet.getName(), 1) - 1, 8).getValues();
    //Logger.log("sheetarray.length="+sheetarray.length);
    for (var i = 0; i < sheetarray.length; i++) {
      if (sheettosortstring == sheetarray[i][0]) {
        var colstosort = 0;
        for (var sorts = 2; sorts < 6; sorts++) {
          if (sheetarray[i][sorts] != 0) {
            colstosort++
          }
        }
  
        if (colstosort = 1)
          for (var sorts = 2; sorts < colstosort + 1; sorts++) {
            if (sheetarray[i][sorts] != 0) {
              // Logger.log("sheettosortstring="+sheettosortstring);
              Logger.log("Sorting " + sheetarray[i][sorts] + "sheettosort.getLastRow()=" + sheettosort.getLastRow() + "sheettosort.getLastColumn()=" + sheettosort.getLastColumn())
              //Logger.log("Sorting "+sheetarray[i][sorts]+" on sheet="+sheettosort.getSheetName);
              sheettosort.getRange(2, 1, sheettosort.getLastRow(), sheettosort.getLastColumn()).sort([{ column: sheetarray[i][sorts], ascending: true }])
            }
          }
      }
    }
  };
  /*function callsortthisrangetest() {
    sortthisrange("TestSort", true)
  };*/
  
  function check_sheets_have_data() {
    var dstart = new Date();
    console.log("check_sheets_have_data Start=" + dstart);
    var parmSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Parameters');
    var splitSheetLastLine = GetLastRowNumber('Parameters', 3);
    var splitSheetNames = parmSheet.getRange(2, 3, splitSheetLastLine - 1, 1).getValues();
    console.log("Status", "Checking for " + splitSheetNames.length + " sheets' existence", "check_sheets_have_data");
    var strMsg = "check_sheets_have_data ran and";
    var strMsg2 = "All Business Sheets Summary\n\nChecked existence for sheets:"
    var message = "";
    var subject = "";
    var emailaddy = "";
    for (var i = 0; i < splitSheetNames.length; i++) {
      strMsg2 = strMsg2 + "\n" + splitSheetNames[i][0]
      var testsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(splitSheetNames[i][0]);
      //logaction("Status", " name=" + splitSheetNames[i][0], "check_sheets_have_data", lastlogline++);
      if (testsheet == null) {
        //logaction("Status", "There is no sheet with name=" + splitSheetNames[i][0], "check_sheets_have_data", lastlogline++);
        strMsg = strMsg + ", Sheet " + splitSheetNames[i][0] + " is missing"
      }
    }
    strMsg2 = strMsg2 + "\n"
    if (strMsg !== "check_sheets_have_data ran and") {
      logaction("Status", strMsg, "check_sheets_have_data")
      emailaddy = "david.typanski@one-line.com";
      subject = "NROF Automation ERROR - At least one sheet is Missing!";
      message = strMsg + "\n\nPlease contact email david.typanski@one-line.com or call cell 804.986.7959\n\n" + strMsg2;
      MailApp.sendEmail(emailaddy, subject, message);
      emailaddy = "adam.matthijs@one-line.com";
      MailApp.sendEmail(emailaddy, subject, message);
      emailaddy = "8049867959@mms.att.ne";
      //MailApp.sendEmail(emailaddy, subject, message);
    }
    else {
      emailaddy = "david.typanski@one-line.com";
      subject = "NROF Automation Report - All Business Sheets Present";
      message = strMsg2;
      //MailApp.sendEmail(emailaddy, subject, message);
      //emailaddy = "8049867959@mms.att.net";
      //MailApp.sendEmail(emailaddy, subject, message);
  
    }
    console.log(strMsg2)
    logaction("Status", "check_sheets_have_data Complete " + (((new Date()) - dstart) / (1000)) + " seconds total", "check_sheets_have_data")
  };
  function changeDate() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Parameters");
    sheet.getRange(1, 20).setValue(new Date(new Date().setHours(0, 0, 0, 0))).setNumberFormat('MM/dd/yyyy');
    logaction("Status", "changeDate Ran", "changeDate");
  };
  function maintainworkbook() {
    changeDate();
    del_automation_rows();
    check_sheets_have_data();
  };
  function del_automation_rows() {
    var autosheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Automation Log");
    var lastlogline = logaction("Status", "del_automation_rows Rows=" + autosheet.getMaxRows(), "del_automation_rows");
    console.log(autosheet.getMaxRows());
    if (autosheet.getMaxRows() >= 3000) {
      autosheet.deleteRows(2, 1000);
      logaction("Status", "del_automation_rows Rows Deleted=1000 New Rows=" + autosheet.getMaxRows(), "del_automation_rows");
    }
    else {
      logaction("Status", "del_automation_rows None Deleted - Must be greater than 3000 to delete. Current MaxRows=" + autosheet.getMaxRows(), "del_automation_rows");
    }
  };
  function exportAllNROFData() {
    var dstart = new Date();
    var lastlogline = logaction("Status", "Export Start=" + dstart, "exportAllNROFData");
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var parmSheet = ss.getSheetByName('Parameters');
    var targetSSID = parmSheet.getRange(3, 22).getValue();
    logaction("Status", "Export - Opening and checking target file", "exportAllNROFData", lastlogline++);
    var errorName = "None";
    try {
      var targetSS = SpreadsheetApp.openById(targetSSID);
    }
    catch (err) {
      errorName = "SSID"
      logaction("Status", "Export - The SpreadsheetID " + targetSSID + " could not be found. Please correct in sheet Parameters Column V row 3. Process Stopped due to error.", "exportAllNROFData", lastlogline++);
      SpreadsheetApp.getUi().alert("The SpreadsheetID " + targetSSID + " could not be found. Please correct in sheet Parameters Column V row 3. Ending Process.");
      return;
    }
    var targetSSSheetName = parmSheet.getRange(4, 22).getValue();
    var tarSheet = targetSS.getSheetByName(targetSSSheetName);
    if (null == targetSS) {
      errorName = "SSID";
      logaction("Status", "Export - The SpreadsheetID " + targetSSID + " could not be found. Please correct in sheet Parameters Column V row 3. Process Stopped due to error.", "exportAllNROFData", lastlogline++);
      SpreadsheetApp.getUi().alert("The SpreadsheetID " + targetSSID + " could not be found. Please correct in sheet Parameters Column V row 3. Ending Process.");
    }
    else if (null == tarSheet) {
      errorName = "Sheet Name";
      logaction("Status", "Export - Sheet named '" + targetSSSheetName + "' could not be found. Please correct in sheet Parameters Column V row 4. Process Stopped due to error.", "exportAllNROFData", lastlogline++);
      SpreadsheetApp.getUi().alert("Export - Sheet named '" + targetSSSheetName + "' could not be found. Please correct in sheet Parameters Column V row 4. Ending Process.");
    }
  
  
    if (errorName == "None") {
      var colstowrite = parmSheet.getRange(3, 23, parmSheet.getMaxRows(), 1).getValues();
      var i = 1;
      var strcolnums = colstowrite[0];
      do {
        strcolnums = strcolnums + ", " + colstowrite[i];
        i++;
      } while (colstowrite[i] != "");
      console.log(strcolnums);
      var numColsToWrite = i;
      console.log("numColsToWrite=" + numColsToWrite);
      colstowrite = colstowrite.slice(0, i);
      console.log(strcolnums);
      logaction("Status", "Export - Columns to be read=" + colstowrite.length, "exportAllNROFData", lastlogline++);
      var srcSheet = ss.getSheetByName("All NROF Bookings");
      var srcCols = parmSheet.getRange(1, 6).getValue();
      console.log(srcCols);
      var srcRows = GetLastRowNumber("All NROF Bookings", 1)
      var allData = srcSheet.getRange(1, 1, srcRows, srcCols).getValues();
      var totalsrcRows = allData.length;
      console.log(totalsrcRows);
      var dataitems = [];
      var tardataitems = [];
      var curcul = 0;
      console.log("here1");
      for (var rowindex = 0; rowindex < totalsrcRows; rowindex++) {
        var dataitems = [];
        for (colIndex = 0; colIndex < colstowrite.length; colIndex++) {
          curcul = colstowrite[colIndex];
          dataitems.push(allData[rowindex][curcul - 1])
        }
        tardataitems.push(dataitems);
      }
      console.log("here2");
      if (tarSheet.getMaxRows() > 2) {
        tarSheet.deleteRows(3, (tarSheet.getMaxRows() - 3));
      }
      tarSheet.clearContents();
      console.log("tardataitems.length=" + tardataitems.length);
      tarSheet.getRange(1, 1, tardataitems.length, numColsToWrite).setValues(tardataitems);
      console.log("Done");
    }
    logaction("Status", "Export End " + (((new Date()) - dstart) / 1000) + " seconds total", "Export", lastlogline++);
  };
  
  function getData() {
    var ss = SpreadsheetApp.getActiveSpreadsheet()
    var sheets = ss.getSheets();
    var srows = 0;
    var scols = 0;
    var totcells = 0;
    var sheetdata = [];
  
    for (var i = 0; i < sheets.length; i++) {
      sheetdatarow = {};
      srows = sheets[i].getMaxRows();
      scols = sheets[i].getMaxColumns();
      totcells = totcells + (srows * scols)
      sheetdatarow.sheetname = sheets[i].getSheetName();
      sheetdatarow.rows = srows;
      sheetdatarow.columns = scols;
      sheetdatarow.scells = (srows * scols);
      sheetdatarow.runningtotal = (totcells);
      sheetdata.push(sheetdatarow)
    }
    return sheetdata;
  };
  function getEmailText(sheetsData) {
    var text = "";
    sheetsData.forEach(function (sheetdatarow) {
      text = text + sheetdatarow.sheetname + "\n" + sheetdatarow.rows + "\n" + sheetdatarow.columns + "\n" + sheetdatarow.scells + "\n" + sheetdatarow.runningtotal + "\n-----------------------\n\n";
    });
    return text;
  };
  function sendEmail() {
    var sheetData = getData();
    var body = getEmailText(sheetData);
    var htmlBody = getEmailHtml(sheetData);
    var emaillist = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Parameters").getSheetValues(2, 24,
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Parameters").getMaxRows() - 1, 1);
    for (var i = 0; i < emaillist.length; i++) {
      if (emaillist[i][0] != "") {
        MailApp.sendEmail({
          to: emaillist[i][0],
          subject: "NROF Booking Report (Working File) Size Update",
          body: body,
          htmlBody: htmlBody
        });
        console.log("Number i=" + i + " email=" + emaillist[i][0]);
      }
    }
  };
  function getEmailHtml(sheetData) {
    var htmlTemplate = HtmlService.createTemplateFromFile("summaryemail.html");
    htmlTemplate.sheets = sheetData;
    var htmlBody = htmlTemplate.evaluate().getContent();
    return htmlBody;
  };
  function deleteSheetData(sname) {
  
    return "No rows were deleted from " + sname
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var std = ss.getSheetByName(sname);
    var rstodel = std.getMaxRows() - 2;
    std.deleteRows(3, rstodel);
    return "Rows 3 to " + rstodel + " were deleted from '" + sname + "'";
  };
  
  