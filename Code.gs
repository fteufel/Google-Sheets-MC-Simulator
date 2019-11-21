function onOpen(e) {
  SpreadsheetApp.getUi()
      .createAddonMenu()
      .addItem('Calculate MC Simulation', 'getMCvalues')
      .addToUi();
}

//utilities
function col2row(column) {
  return [column.map(function(row) {return row[0];})];
} 

function row2col(row) {
  return row[0].map(function(elem) {return [elem];});
}

function transpose(a)
{
  return Object.keys(a[0]).map(function (c) { return a.map(function (r) { return r[c]; }); });
}

/**
 * Force Spreadsheet to re-calculate selected cells
 https://gist.github.com/katz/ab751588580469b35e08#file-recalculatesellectedcells-gs
 */
function recalculate(){
  var activeRange = SpreadsheetApp.getActiveRange();
  var originalFormulas = activeRange.getFormulas();
  var originalValues = activeRange.getValues();

  var valuesToEraseFormula = [];
  var valuesToRestoreFormula = [];

  originalFormulas.forEach(function(outerVal, outerIdx){
    valuesToEraseFormula[outerIdx] = [];
    valuesToRestoreFormula[outerIdx] = [];
    outerVal.forEach(function(innerVal, innerIdx){
      if('' === innerVal){
        //The cell doesn't have formula
        valuesToEraseFormula[outerIdx][innerIdx] = originalValues[outerIdx][innerIdx];
        valuesToRestoreFormula[outerIdx][innerIdx] = originalValues[outerIdx][innerIdx];
      }else{
        //The cell has a formula.
        valuesToEraseFormula[outerIdx][innerIdx] = '';
        valuesToRestoreFormula[outerIdx][innerIdx] = originalFormulas[outerIdx][innerIdx];
      }
    })
  })

  activeRange.setValues(valuesToEraseFormula);
  activeRange.setValues(valuesToRestoreFormula);
}

function getMCvalues(){
  var n_mc_smps = 200
  //get which cells should be simulated
  var targetcells = SpreadsheetApp.getActiveRange();
  var cells = targetcells.getValues();
  
  //set up arrays to collect results
  var targetcells = [];
  var simulations = [];
  for(i=0;i<cells.length;i++){
    var targetcell = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(cells[i]);
    targetcells.push(targetcell);
    simulations.push([]);
  }
    

  

  
  Logger.log(targetcell);
  //make new sheet
  var name = (new Date()).toLocaleDateString() + " MC Simulation";
  SpreadsheetApp.getActiveSpreadsheet().insertSheet(name);
  
  
  //iterate over all mc samples
  for(i = 0;i < n_mc_smps ;i++){
    Logger.log("entered loop "+i);
    //calculate
    recalculate();
    
     Logger.log("entering inner loop");
    for(j=0;j<targetcells.length;j++){
      var value = targetcells[j].getValue();
      simulations[j].push(value);
      Logger.log("loop over cells " + j);
    }
    
    //append to list                        
    //var lastrow = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name).getLastRow() +1
    //SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name).getRange("A"+lastrow).setValue(value);
  }

  var lastrow = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name).getLastRow() +1
  //SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name).getRange("A"+lastrow).setValue(simulations[0]);
  //Logger.log(simulations);
  
  // write all at once
  //var writeRange = "A"+lastrow+":B"+(lastrow+n_mc_smps -1);
   Logger.log(simulations);
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name).getRange(1,1,1,targetcells.length).setValues(transpose(cells));
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name).getRange(2,1, n_mc_smps,targetcells.length).setValues(transpose(simulations));
}


