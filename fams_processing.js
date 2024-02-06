function frequencyAnalysis() {
  //analysis sections are array of analysis and respective column
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var resultsSheet = spreadsheet.getSheetByName("stalking sheet") || spreadsheet.insertSheet("stalking sheet");
  resultsSheet.clear(); // Clear existing content

  var allSheets = spreadsheet.getSheets();

  function processAnalysis(col){
    var wordFrequencyMap = {};
    //REPLACE WITH A FAM NAME
    var analysisName = spreadsheet.getSheetByName("bad biddies").getRange(col+"9").getValue();
    allSheets.forEach(function(sheet) {
      var dataRanges = [];
      var nameRanges = sheet.getRange("B2:B122"); // Assuming names are in the range B2:B122

      // Assuming the pattern starts from C10 and increments by 20
      for (var startRow = 10; startRow <= 130; startRow += 20) {
        var dataRange = sheet.getRange(col + startRow + ":" +col + (startRow + 4));
        dataRanges.push.apply(dataRanges, dataRange.getValues());
      }

      dataRanges.forEach(function(row, index) {
        var cellValue = row[0];
        var personNameIndex = Math.floor(index / 5) * 20 + 1;
        var personName = nameRanges.getCell(personNameIndex, 1).getValue();
        
        if (typeof cellValue === 'string' || cellValue instanceof String) {
          //make all slashes spaces, all commas spaces, remove all non alphanumeric
          cellValue = cellValue.replace(/,/g, ' ');
          cellValue = cellValue.replace(/\//g, ' ');
          cellValue = cellValue.replace(/[^a-zA-Z0-9 ]/g, '');

          var words = cellValue.toLowerCase().split(" ");

          for (var i = 0; i < words.length; i++) {
            var fullSubstring = words[i].replace(/s$/, ''); // Remove trailing 's';
            
            // Normalize the word by replacing spaces with hyphens
            var normalizedWord = fullSubstring.replace(/\s/g, '-');

            // Count the individual word
            if (!wordFrequencyMap[normalizedWord]) {
              wordFrequencyMap[normalizedWord] = { frequency: 0, names: [] };
            }

            // Ensure unique names are added to the names array
            if (!wordFrequencyMap[normalizedWord].names.includes(personName)) {
              wordFrequencyMap[normalizedWord].frequency++;
              wordFrequencyMap[normalizedWord].names.push(personName);
            }

            // Count multi-word combinations
            for (var j = i + 1; j < words.length; j++) {
              fullSubstring += '-' + words[j];
              
              // Normalize the multi-word combination
              var normalizedMultiWord = fullSubstring.replace(/\s/g, '-');

              if (!wordFrequencyMap[normalizedMultiWord]) {
                wordFrequencyMap[normalizedMultiWord] = { frequency: 0, names: [] };
              }

              // Ensure unique names are added to the names array
              if (!wordFrequencyMap[normalizedMultiWord].names.includes(personName)) {
                wordFrequencyMap[normalizedMultiWord].frequency++;
                wordFrequencyMap[normalizedMultiWord].names.push(personName);
              }
            }
          }
        }
      });
    });
    
    
    // Convert wordFrequencyMap to an array of objects for sorting
    var wordFrequencyArray = Object.keys(wordFrequencyMap).map(function(word) {
      return { word: word, frequency: wordFrequencyMap[word].frequency, names: wordFrequencyMap[word].names };
    });

    // Filter the array to include only entries with frequency > 1
    var filteredWordFrequencyArray = wordFrequencyArray.filter(function(entry) {
      //filter len >3 and freq >2
      return ((entry.word.length>3 || entry.word =="rnb" || entry.word == "pop" ) &&  entry.frequency > 2) ;
    });

    // Sort the filtered array by frequency in descending order
    filteredWordFrequencyArray.sort(function(a, b) {
      return b.frequency - a.frequency;
    });

    Logger.log(filteredWordFrequencyArray);

    // Append the results to "stalking sheet" sheet
    var lastRow = resultsSheet.getLastRow() + 1;
    resultsSheet.getRange("A" + lastRow).setValue(analysisName);
    resultsSheet.getRange("B" + lastRow).setValue("Frequency");
    resultsSheet.getRange("C" + lastRow).setValue("Liked By");

    //highlight that row
    var rowRange = resultsSheet.getRange(lastRow, 1, 1, resultsSheet.getLastColumn());
    rowRange.setBackground('#ebf584');  // Yellow background color

    // Write data starting from the next row
    for (var i = 0; i < filteredWordFrequencyArray.length; i++) {
      resultsSheet.getRange("A" + (lastRow + i + 1)).setValue(filteredWordFrequencyArray[i].word);
      resultsSheet.getRange("B" + (lastRow + i + 1)).setValue(filteredWordFrequencyArray[i].frequency);
      resultsSheet.getRange("C" + (lastRow + i + 1)).setValue(filteredWordFrequencyArray[i].names.join(", "));
    }
  }

  var analysisTypes = ["B", "C", "D", "E", "F", "G", "H", "I"];
  for (const analysisType of analysisTypes){
    processAnalysis(analysisType);
  }
}

//TODO: Comment this out if you don't want it to run every edit!
function onEdit(e){
  frequencyAnalysis();
}
