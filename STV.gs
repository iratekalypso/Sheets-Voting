function runSTV() {
  // Get the active spreadsheet and the sheet named "Votes"
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Votes");
  // Retrieve all data from the sheet
  const data = sheet.getDataRange().getValues();
  // Set the number of winners to be elected
  const numWinners = 5;
  // Initialize an array to store the ballots
  const ballots = [];
  // Create a TieBreaker instance with a range for 16 candidates
  const tieBreaker = new TieBreaker([...Array(16).keys()]);

  // Process each row in the sheet, starting from the second row
  for (let i = 1; i < data.length; i++) {
    // Skip the Voter ID column and get the rest of the row
    const row = data[i].slice(1);
    // Filter out empty cells and convert the candidates to strings
    const ballot = row.filter(cell => cell).map(candidate => candidate.toString());
    // Add the ballot to the ballots array if it is not empty
    if (ballot.length > 0) {
      ballots.push({ ballot, count: 1 });
    }
  }

  // Instantiate the STV class with the ballots, tieBreaker, and number of winners
  const stv = new STV(ballots, tieBreaker, numWinners);
  // Calculate the STV results
  stv.calculate_results();

  // Get or create the sheet named "Results" for output
  let resultSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Results");
  if (resultSheet) {
    resultSheet.clear(); // Clear the sheet if it already exists
  } else {
    resultSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Results");
  }

  let row = 1;
  // Write the Droop Quota to the results sheet
  resultSheet.getRange(row, 1).setValue("Droop Quota");
  resultSheet.getRange(row++, 2).setValue(stv.quota);

  // Write the results of each round to the sheet
  for (const [roundIndex, round] of stv.rounds.entries()) {
    row++;
    resultSheet.getRange(row++, 1).setValue(`Round ${roundIndex + 1}`);
    resultSheet.getRange(row, 1).setValue("Candidate");
    resultSheet.getRange(row++, 2).setValue("Tally");

    // Write the tally for each candidate
    for (const [candidate, tally] of Object.entries(round.tallies)) {
      resultSheet.getRange(row, 1).setValue(candidate);
      resultSheet.getRange(row++, 2).setValue(tally);
    }

    // Write the winners, loser, and any notes for the round
    if (round.winners) {
      resultSheet.getRange(row, 1).setValue("Winners");
      resultSheet.getRange(row++, 2).setValue([...round.winners].join(", "));
    }

    if (round.loser) {
      resultSheet.getRange(row, 1).setValue("Loser");
      resultSheet.getRange(row++, 2).setValue(round.loser);
    }

    if (round.note) {
      resultSheet.getRange(row, 1).setValue("Note");
      resultSheet.getRange(row++, 2).setValue(round.note);
    }
  }

  row++;
  // Write the final winners
  resultSheet.getRange(row, 1).setValue("Final Winners");
  resultSheet.getRange(row++, 2).setValue([...stv.winners].join(", "));
}

class STV {
  constructor(ballots, tie_breaker, required_winners = 1) {
    // Initialize the STV instance with ballots, tie breaker, and number of winners
    this.ballots = ballots;
    this.tie_breaker = tie_breaker;
    this.required_winners = required_winners;
    this.winners = new Set();
    this.rounds = [];
  }

  static droop_quota(ballots, seats = 1) {
    // Calculate the Droop quota for the given ballots and seats
    let voters = 0;
    for (const ballot of ballots) {
      if (ballot.ballot.length > 0) {
        voters += ballot.count;
      }
    }
    return Math.floor(voters / (seats + 1)) + 1;
  }

  calculate_results() {
    // Calculate the election results using the STV method
    const candidates = new Set();
    for (const ballot of this.ballots) {
      candidates.add(...ballot.ballot);
    }
    if (candidates.size < this.required_winners) {
      throw new Error("Not enough candidates provided");
    }

    this.quota = STV.droop_quota(this.ballots, this.required_winners);
    const remaining_candidates = new Set(candidates);

    // Run the election rounds until the required number of winners is reached
    while (this.winners.size < this.required_winners && remaining_candidates.size > 0) {
      const round = { tallies: this.tally_votes(this.ballots) };

      // Check if any candidate meets or exceeds the quota
      if (Math.max(...Object.values(round.tallies)) >= this.quota) {
        round.winners = new Set();
        for (const [candidate, tally] of Object.entries(round.tallies)) {
          if (tally >= this.quota) {
            this.winners.add(candidate);
            round.winners.add(candidate);
            remaining_candidates.delete(candidate);
            this.redistribute_votes(candidate, tally - this.quota);
          }
        }
      } else {
        const loser = this.find_loser(round.tallies);
        round.loser = loser;
        remaining_candidates.delete(loser);
        this.remove_candidate_from_ballots(loser);
      }

      this.rounds.push(round);
    }

    // If not enough winners are found, add remaining candidates until the required number is reached
    if (this.winners.size < this.required_winners) {
      for (const candidate of remaining_candidates) {
        this.winners.add(candidate);
        if (this.winners.size >= this.required_winners) {
          break;
        }
      }
    }
  }

  tally_votes(ballots) {
    // Count the votes for each candidate
    const tallies = {};
    for (const ballot of ballots) {
      if (ballot.ballot.length > 0) {
        const top_choice = ballot.ballot[0];
        if (!tallies[top_choice]) {
          tallies[top_choice] = 0;
        }
        tallies[top_choice] += ballot.count;
      }
    }
    return tallies;
  }

  redistribute_votes(winner, excess) {
    // Redistribute the excess votes of the winner
    for (const ballot of this.ballots) {
      if (ballot.ballot.length > 0 && ballot.ballot[0] === winner) {
        ballot.count *= excess / (excess + this.quota);
        ballot.ballot.shift(); // Remove the winner from the ballot
      }
    }
  }

  remove_candidate_from_ballots(candidate) {
    // Remove a candidate from all ballots
    for (const ballot of this.ballots) {
      const index = ballot.ballot.indexOf(candidate);
      if (index > -1) {
        ballot.ballot.splice(index, 1);
      }
    }
  }

  find_loser(tallies) {
    // Find the candidate with the fewest votes and break ties if necessary
    const minVotes = Math.min(...Object.values(tallies));
    const losers = Object.keys(tallies).filter(candidate => tallies[candidate] === minVotes);
    return this.tie_breaker.break_ties(losers);
  }
}

class TieBreaker {
  constructor(candidate_range) {
    // Initialize and shuffle the candidate order
    this.random_ordering = candidate_range.slice();
    this.shuffle(this.random_ordering);
  }

  shuffle(array) {
    // Shuffle the array using the Fisher-Yates algorithm
    for (let i = array.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [array[i], array[j]] = [array[j], array[i]];
    }
  }

  break_ties(tied_candidates) {
    // Break ties by selecting the first candidate in the shuffled order
    for (const candidate of this.random_ordering) {
      if (tied_candidates.includes(candidate)) {
        return candidate;
      }
    }
    return tied_candidates[0]; // Return the first candidate if no tie-breaker found
  }
}
