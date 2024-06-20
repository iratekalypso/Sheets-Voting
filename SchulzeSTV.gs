function runSchulzeSTV() {
  // Get the active spreadsheet and the "Votes" sheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Votes");
  // Retrieve all data from the sheet
  const data = sheet.getDataRange().getValues();
  // Set the number of winners to be elected
  const numWinners = 5;
  // Initialize an array to store the ballots and a set for unique candidates
  const ballots = [];
  const candidatesSet = new Set();

  // Process each row in the sheet, starting from the second row
  for (let i = 1; i < data.length; i++) {
    // Skip the Voter ID column and get the rest of the row
    const row = data[i].slice(1);
    // Filter out empty cells and convert the candidates to strings
    const ballot = row.filter(cell => cell).map(candidate => candidate.toString());
    // Add the ballot to the ballots array if it is not empty and update the candidates set
    if (ballot.length > 0) {
      ballots.push(ballot);
      ballot.forEach(candidate => candidatesSet.add(candidate));
    }
  }

  // Convert the candidates set to an array
  const candidates = Array.from(candidatesSet);
  // Instantiate the SchulzeSTV class with the ballots, candidates, and number of winners
  const schulzeSTV = new SchulzeSTV(ballots, candidates, numWinners);
  // Calculate the results using the Schulze method
  const results = schulzeSTV.calculateResults();

  // Get or create the sheet named "Results" for output
  let resultSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Results");
  if (resultSheet) {
    resultSheet.clear(); // Clear the sheet if it already exists
  } else {
    resultSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Results");
  }

  // Write headers and detailed results to the "Results" sheet
  let row = 1;
  resultSheet.getRange(row++, 1).setValue("Schulze STV Detailed Results");
  resultSheet.getRange(row++, 1).setValue(`Number of Winners: ${numWinners}`);
  resultSheet.getRange(row++, 1).setValue("Candidates:");
  candidates.forEach(candidate => {
    resultSheet.getRange(row++, 2).setValue(candidate);
  });

  row++;
  resultSheet.getRange(row++, 1).setValue("Pairwise Preferences Matrix:");
  row = writeMatrix(resultSheet, schulzeSTV.d, candidates, row);

  row++;
  resultSheet.getRange(row++, 1).setValue("Strongest Paths Matrix:");
  row = writeMatrix(resultSheet, schulzeSTV.p, candidates, row);

  row++;
  resultSheet.getRange(row++, 1).setValue("Ranking of Candidates:");
  results.forEach((result, index) => {
    resultSheet.getRange(row++, 1).setValue(`Rank ${index + 1}: Candidate ${result.candidate} with score ${result.score}`);
  });

  // Auto-resize columns for better readability
  resultSheet.autoResizeColumns(1, 2);
}

function writeMatrix(sheet, matrix, candidates, startRow) {
  // Write the matrix to the sheet, starting from the specified row
  let row = startRow;
  sheet.getRange(row++, 1).setValue("From/To");
  // Write candidate names as headers
  candidates.forEach((candidate, index) => {
    sheet.getRange(row - 1, index + 2).setValue(candidate);
    sheet.getRange(row + index, 1).setValue(candidate);
  });

  // Write matrix values
  for (let i = 0; i < candidates.length; i++) {
    for (let j = 0; j < candidates.length; j++) {
      sheet.getRange(row + i, j + 2).setValue(matrix[i][j]);
    }
  }

  return row + candidates.length;
}

class SchulzeSTV {
  constructor(ballots, candidates, numWinners) {
    // Initialize the SchulzeSTV instance with ballots, candidates, and number of winners
    this.ballots = ballots;
    this.candidates = candidates;
    this.numWinners = numWinners;
    this.numCandidates = candidates.length;
    // Initialize pairwise preferences and strongest paths matrices
    this.d = this.initializeMatrix(this.numCandidates);
    this.p = this.initializeMatrix(this.numCandidates);
  }

  initializeMatrix(size) {
    // Create a matrix of the specified size, initialized to zeros
    const matrix = [];
    for (let i = 0; i < size; i++) {
      matrix[i] = Array(size).fill(0);
    }
    return matrix;
  }

  calculatePairwisePreferences() {
    // Calculate the pairwise preferences matrix
    const index = (candidate) => this.candidates.indexOf(candidate);
    for (let ballot of this.ballots) {
      for (let i = 0; i < ballot.length; i++) {
        for (let j = i + 1; j < ballot.length; j++) {
          if (ballot[i] !== ballot[j]) {
            this.d[index(ballot[i])][index(ballot[j])] += 1;
          }
        }
      }
    }
    // Ensure diagonal elements are zero
    for (let i = 0; i < this.numCandidates; i++) {
      this.d[i][i] = 0;
    }
  }

  calculateStrongestPaths() {
    // Calculate the strongest paths matrix using the Floyd-Warshall algorithm
    for (let i = 0; i < this.numCandidates; i++) {
      for (let j = 0; j < this.numCandidates; j++) {
        if (i !== j) {
          if (this.d[i][j] > this.d[j][i]) {
            this.p[i][j] = this.d[i][j];
          } else {
            this.p[i][j] = 0;
          }
        }
      }
    }

    for (let i = 0; i < this.numCandidates; i++) {
      for (let j = 0; j < this.numCandidates; j++) {
        if (i !== j) {
          for (let k = 0; k < this.numCandidates; k++) {
            if (i !== k && j !== k) {
              this.p[j][k] = Math.max(this.p[j][k], Math.min(this.p[j][i], this.p[i][k]));
            }
          }
        }
      }
    }
  }

  rankCandidates() {
    // Rank candidates based on the strongest paths matrix
    const ranking = [];
    for (let i = 0; i < this.numCandidates; i++) {
      let count = 0;
      for (let j = 0; j < this.numCandidates; j++) {
        if (i !== j) {
          if (this.p[i][j] > this.p[j][i]) {
            count++;
          }
        }
      }
      ranking.push({ candidate: this.candidates[i], score: count });
    }
    ranking.sort((a, b) => b.score - a.score);
    return ranking.slice(0, this.numWinners);
  }

  calculateResults() {
    // Calculate the election results using the Schulze method
    this.calculatePairwisePreferences();
    this.calculateStrongestPaths();
    return this.rankCandidates();
  }
}
