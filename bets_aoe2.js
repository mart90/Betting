// Globals

var admin = SpreadsheetApp.getActiveSpreadsheet()
var tech = admin.getSheetByName("Tech")
var scriptparams = admin.getSheetByName("ScriptParams")
var view = SpreadsheetApp.openById(tech.getRange("o7").getValue())
var numplayers = numPlayers()

//--------

function alert(message) {
    tech.getRange(1, 1).setValue("Script message: " + message)
}

function numPlayers() {
    return tech.getRange("H9").getValue()
}

function editPlayerSheets(row, column, value, startatrowoverride, endatrowoverride) {
    var startatrow = startatrowoverride || 3
    var endatrow = endatrowoverride || numplayers + 3

    if (!startatrowoverride && endatrow >= 500) {
        endatrow = 500
    }

    var playersheetids = tech.getRange(3, 6, numplayers).getValues()
    for (var i = startatrow; i < endatrow; i++) {
        SpreadsheetApp.openById(playersheetids[i - 3][0]).getSheetByName('Sheet1').getRange(row, column).setValue(value)
    }
}

function newPlayers(playernames, viewid) {
    var numplayersold = numplayers
    numplayers += playernames.length

    for (var i = 0; i < playernames.length; i++) {
        newPlayer(numplayersold + 3 + i, playernames[i])
    }

    var openmatches = [];
    for (var j = 1; j < 6; j++) {
        if (tech.getRange(5, 8 + j).getValue() == "OPEN")
            openmatches.push(j)
    }

    for (var k = 0; k < openmatches.length; k++) {
        linkSheets(openmatches[k], numplayers, viewid)
    }

    tech.getRange("H9").setValue(numplayers)

    makeMessages(numplayersold + 3)

    function newPlayer(playernumber, playername) {
        var playertemplate = SpreadsheetApp.openById(tech.getRange("o9").getValue())

        var newss = playertemplate.copy(playernumber + '. ' + playername).getId()
        var newsheet = SpreadsheetApp.openById(newss).getSheetByName('Sheet1')
        var newfile = DriveApp.getFileById(newss)

        newfile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT)

        newsheet.getRange(23, 10).setValue('=importrange("' + tech.getRange("o11").getValue() + '","Tech!A' + playernumber + '")')

        // Insert new player into admin sheet
        if (!tech.getRange(playernumber, 1).isBlank()) {
            // Already a name in this cell. Probably an error in numPlayers
            alert("NewPlayer(): Tried to add new player but tech row was already populated")
            return
        }
        
        tech.getRange(playernumber, 1).setValue(playername)
        tech.getRange(playernumber, 6).setValue(newss)
        tech.getRange(playernumber, 7).setValue("=1500+Q" + playernumber + "+D" + playernumber + "-C" + playernumber)
    }
}

// Generates PMs for new sign ups
function makeMessages(startatrow) {
    for (var i = startatrow; i < numplayers + 3; i++) {
        var playername = tech.getRange(i, 1).getValue()
        var playersheetid = tech.getRange(i, 6).getValue()

        tech.getRange(i, 18).setValue(
            'Hi ' + playername + ',[br][/br]' +
            '[br][/br]' +
            'Welcome to and thank you for participating in T2T 1v1 Chameleon tournament betting! [url=https://docs.google.com/spreadsheets/d/' + playersheetid + '/edit#gid=0&vpid=A1][u]Here[/u][/url] is a link to your personal betting sheet. [br][/br]' +
            'You\'re probably going to want to bookmark this link. Share it at your own risk, anyone who has it will be able to edit your bets.[br][/br]' +
            '[br][/br]' +
            'And [url=https://docs.google.com/spreadsheets/d/1jqUly_w72S5lSVRxWBIuB5WIwK39dFO91gi0Yfd1tYU/edit#gid=557307683][u]here[/u][/url] is a link to the stats spreadsheet.[br][/br]' +
            'Notice there are multiple tabs (switch tabs on the bottom left). One for the current standings, one for an overview of betting history this event, one per finished match, and one per current betting match where you can see everyone else\'s current bets.[br][/br]' +
            '[br][/br]' +
            'Information about new betting matches and results will be posted in [url=http://eso-community.net/viewtopic.php?f=468&t=13665][u]the discussion thread[/u][/url] where you will be tagged from now on. Feel free to discuss anything related to betting there.[br][/br]' +
            '[br][/br]' +
            'You can read back the signups thread containing info and rules [url=http://eso-community.net/viewtopic.php?f=468&t=13664][u]here[/u][/url].[br][/br]' +
            '[br][/br]' +
            'If you have any questions or feedback feel free to reply to this message.[br][/br] Have fun! :flowers:')
    }
}

function open(matchnumber, numberofmatches) {
    for (var a = matchnumber; a < matchnumber + numberofmatches; a++) {
        tech.getRange(5, a + 8).setValue("OPEN")
        view.getSheetByName("Match " + a).getRange("A2").setValue("Bets open")

        linkSheets(a)
    }
}

function close(matchnumber, numberofmatches) {
    for (var a = matchnumber; a < matchnumber + numberofmatches; a++) {
        tech.getRange(5, a + 8).setValue("CLOSED")
        view.getSheetByName("Match " + a).getRange("A2").setValue("Bets closed")

        unlinkSheets(a)
        
        removeOverbets(a)
    }
}

// Opens bets by linking player sheets to admin sheet
// Bet changes in player sheets are now recorded in the system
function linkSheets(matchnumber) {
    var adminmatchsheet = admin.getSheetByName('M00' + matchnumber)

    if (matchnumber == 1) column = 'B'
    else if (matchnumber == 2) column = 'F'
    else if (matchnumber == 3) column = 'J'
    else if (matchnumber == 4) column = 'N'
    else if (matchnumber == 5) column = 'R'

    for (var i = 3; i < numplayers + 3; i++) {
        for (var a = 0; a < 12; a++) {
            adminmatchsheet.getRange(i, a + 3).setValue('=Bets!' + column + (12 * (i - 3) + a + 2).toString())
        }
    }
}

// Closes bets by unlinking player sheets
// Copies the current bets and pastes the values only, so that they can no longer be influenced
// Limits bets in case someone got past the input validation
function unlinkSheets(matchnumber) {
    var adminmatchsheet = admin.getSheetByName('M00' + matchnumber)
    var range = adminmatchsheet.getRange(3, 3, numplayers, 12)

    range.copyValuesToRange(adminmatchsheet, 3, 12, 3, numplayers + 2)

    limitBets()

    function limitBets() {
        var winninglimit = 10000, resultlimit = 1000

        var bets = range.getValues()
        for (var row = 0; row < bets.length; row++) {
            var rowdata = bets[row]

            for (var column = 0; column < rowdata.length; column++) {
                if (column < 2) {
                    if (rowdata[column] > winninglimit) {
                        adminmatchsheet.getRange(row + 3, column + 3).setValue(winninglimit)
                    }
                }
                else if (rowdata[column] > resultlimit) {
                    adminmatchsheet.getRange(row + 3, column + 3).setValue(resultlimit)
                }
            }
        }
    }
}

// Removes bets on a match by players who bet more than they have
function removeOverbets(matchnumber) {
    var matchsheet = admin.getSheetByName('M00' + matchnumber)
    var playerbalances = tech.getRange(3, 2, numplayers).getValues()

    for (var i = 3; i < numplayers + 3; i++) {
        if (playerbalances[i - 3][0] < -2) {
            matchsheet.getRange(i, 3, 1, 10).setValue(0)
        }
    }
}

// Sets match result in Q15:R16 of match sheet
// Payouts are calculated on the sheet itself (column M) based on what is set here
function setWinner(matchnumber, result) {
    var adminmatchsheet = admin.getSheetByName("M00" + matchnumber)
    var viewmatchsheet = view.getSheetByName("Match " + matchnumber)

    for (i = 6; i <= 15; i++) {
        if (adminmatchsheet.getRange(i, 17).getValue() != result) {
            continue
        }

        // Result matched
        var wincolumn, winodds, resultcolumn, resultodds

        if (i < 10) {
            wincolumn = "C";
            winodds = roundToTwo(adminmatchsheet.getRange("R4").getValue())
        }
        else {
            wincolumn = "D";
            winodds = roundToTwo(adminmatchsheet.getRange("R5").getValue())
        }

        switch (i) {
            case 6: resultcolumn = "E"; break
            case 7: resultcolumn = "F"; break
            case 8: resultcolumn = "G"; break
            case 9: resultcolumn = "H"; break
            case 10: resultcolumn = "I"; break
            case 11: resultcolumn = "J"; break
            case 12: resultcolumn = "K"; break
            case 13: resultcolumn = "L"; break
            case 14: resultcolumn = "M"; break
            case 15: resultcolumn = "N"; break
        }

        resultodds = roundToTwo(adminmatchsheet.getRange(i, 18).getValue())
    }

    adminmatchsheet.getRange("S17").setValue(wincolumn)
    adminmatchsheet.getRange("S18").setValue(winodds)
    adminmatchsheet.getRange("T17").setValue(resultcolumn)
    adminmatchsheet.getRange("T18").setValue(resultodds)

    viewmatchsheet.getRange("S17").setValue(wincolumn)
    viewmatchsheet.getRange("S18").setValue(winodds)
    viewmatchsheet.getRange("T17").setValue(resultcolumn)
    viewmatchsheet.getRange("T18").setValue(resultodds)

    function roundToTwo(num) {
        return +(Math.round(num + "e+2") + "e-2");
    }
}

// End match script
// Run after setWinner()
function endMatch(matchnumber) {
    archive(matchnumber)
    updateTech(matchnumber)
    updateHistory(matchnumber)
    // resetBets(matchnumber)
    clearMatchSheets(matchnumber)
    // consistencyCalc()
}

// Archives view match sheet by duplicating it and renaming to "[PlayerA] vs [PlayerB]"
// The new sheet will be in the view spreadsheet
function archive(matchnumber) {
    var matchsheet = view.getSheetByName('Match ' + matchnumber)
    var rangetocopy = matchsheet.getRange('A1:T1000')
    var playera = matchsheet.getRange('C2').getValue()
    var playerb = matchsheet.getRange('D2').getValue()

    matchsheet.copyTo(view).setName(playera + ' vs ' + playerb)
    var newsheet = view.getSheetByName(playera + ' vs ' + playerb)

    rangetocopy.copyValuesToRange(newsheet, 1, 20, 1, 1000)
    newsheet.getRange('A2').clear()
    newsheet.getRange(3, 1, numplayers, 16).sort({column: 16, ascending: false})

    removeNonParticipants()

    function removeNonParticipants() {
        var pointsbet = newsheet.getRange(3, 2, numplayers).getValues()

        for (var i = 0; i < pointsbet.length; i++) {
            if (pointsbet[i] == 0)  {
                var startrow = i + 3
                break
            }
        }

        if (!startrow) {
            return
        }

        for (var j = startrow; j < pointsbet.length; j++) {
            if (pointsbet[j] > 0) {
                var endrow = j + 2
                break
            }
        }

        newsheet.getRange(endrow + 1, 1, 1000 - startrow, 16).copyValuesToRange(newsheet, 1, 16, startrow, 1000)
    }
}

// Updates admin.tech:
// - Column Q
// - Column E
// - Match history
function updateTech(matchnumber) {
    var matchsheet = admin.getSheetByName('M00' + matchnumber)
    var roundnumber = tech.getRange('H14').getValue() + 1

    var totalprofits = tech.getRange(3, 17, numplayers).getValues()
    var pastbets = tech.getRange(3, 5, numplayers).getValues()
    var profits = matchsheet.getRange(3, 16, numplayers).getValues()
    var currentbets = matchsheet.getRange(3, 2, numplayers).getValues()

    var syspointsbefore = tech.getRange("i10").getValue()

    // Add profits to total profits and bets made to past bets
    for (var i = 3; i < numplayers + 3; i++) {
        tech.getRange(i, 17).setValue(parseInt(totalprofits[i - 3]) + parseInt(profits[i - 3]))
        tech.getRange(i, 5).setValue(parseInt(pastbets[i - 3]) + parseInt(currentbets[i - 3]))
    }

    setMatchStats()

    function setMatchStats() {
        setInflation()
        setScore()

        tech.getRange('H14').setValue(roundnumber);
        tech.getRange(17 + roundnumber, 9).setValue(matchsheet.getRange('T15').getValue()) 
        tech.getRange(17 + roundnumber, 10).setValue(matchsheet.getRange('R19').getValue()) 
        tech.getRange(17 + roundnumber, 11).setValue(matchsheet.getRange('R20').getValue()) 

        var playera = tech.getRange(3, 8 + matchnumber).getValue()
        var playerb = tech.getRange(4, 8 + matchnumber).getValue()
        tech.getRange(17 + roundnumber, 12).setValue(playera)
        tech.getRange(17 + roundnumber, 13).setValue(playerb)

        function setInflation() {
            var syspointsafter = 0;

            var networths = tech.getRange(3, 7, numplayers).getValues()
            for (var i = 0; i < networths.length; i++) {
                syspointsafter += networths[i][0]
            }

            if (syspointsafter == syspointsbefore) {
                // Net worths probably weren't updated properly
                alert("SetInflation(): Points in system after was equal to before")
            }

            var inflation = (syspointsafter - syspointsbefore) / syspointsbefore
            tech.getRange(17 + roundnumber, 15).setValue(inflation)
        }

        function setScore() {
            var scorecolumn = matchsheet.getRange("T17").getValue()
            var scorerow

            switch (scorecolumn) {
                case "E": scorerow = 6; break
                case "F": scorerow = 7; break
                case "G": scorerow = 8; break
                case "H": scorerow = 9; break
                case "I": scorerow = 10; break
                case "J": scorerow = 11; break
                case "K": scorerow = 12; break
                case "L": scorerow = 13; break
                case "M": scorerow = 14; break
                case "N": scorerow = 15; break
            }

            if (scorerow == undefined) {
                alert("SetScore(): Scorerow not set")
                return
            }

            tech.getRange(17 + roundnumber, 14).setValue(matchsheet.getRange(scorerow, 17).getValue())
        }
    }
}

// Adds new columns in view.history and populates with data
function updateHistory(matchnumber) {
    var history = view.getSheetByName('History')
    var matchsheet = admin.getSheetByName('M00' + matchnumber)

    // Create new columns
    history.insertColumnsBefore(1, 4)
    var newcolumns = history.getRange(1, 1)
    history.getRange(1, 5, 300, 3).copyFormatToRange(history, 1, 3, 1, 300)
    history.getRange(1, 5, 300, 3).copyTo(newcolumns)

    // Get data from admin sheet
    var roundnumber = tech.getRange('H14').getValue()
    var players = tech.getRange(3, 1, numplayers, 1).getValues()
    var networths = tech.getRange(3, 7, numplayers, 1).getValues()
    var profits = matchsheet.getRange(3, 16, numplayers, 1).getValues()

    var rownumber = roundnumber + 17
    var playera = matchsheet.getRange('C2').getValue()
    var playerb = matchsheet.getRange('D2').getValue()
    var result = tech.getRange(rownumber, 14).getValue()
    var totalbets = tech.getRange(rownumber, 9).getValue()
    var bigwin = tech.getRange(rownumber, 10).getValue()
    var bigloss = tech.getRange(rownumber, 11).getValue()

    var amountofbetters = getAmountOfBetters()
    var odds = getResultOdds()

    // Populate history
    history.getRange(1, 1).setValue('Round ' + roundnumber)
    history.getRange(4, 2).setValue(playera)
    history.getRange(5, 2).setValue(playerb)
    history.getRange(6, 2).setValue(result)
    history.getRange(7, 2).setValue(odds)
    history.getRange(8, 2).setValue(totalbets + " (" + amountofbetters + " bets)")
    history.getRange(9, 2).setValue(bigwin)
    history.getRange(10, 2).setValue(bigloss)

    history.getRange(13, 1, numplayers, 1).setValues(players)
    history.getRange(13, 2, numplayers, 1).setValues(networths)
    history.getRange(13, 3, numplayers, 1).setValues(profits)
    history.getRange(13, 1, numplayers, 3).sort({column: 2, ascending: false})

    history.getRange(1, 4, 300, 1).clear()

    function getAmountOfBetters() {
        var amountofbetters = 0
        for (i = 0; i < profits.length; i++) {
            if (profits[i] != 0)
                amountofbetters++
        }
        return amountofbetters
    }

    function getResultOdds() {
        for (i = 0; i < 8; i++) {
            if (matchsheet.getRange(6 + i, 17).getValue() == result) {
                return matchsheet.getRange(6 + i, 18).getValue()
            }
        }
    }
}

// Resets all bets to zero on player sheets for given matchnumber
function resetBets(matchnumber) {
    for (var i = 3; i < numplayers + 3; i++) {
        SpreadsheetApp.openById(tech.getRange(i, 6).getValue()).getSheetByName("Sheet1").getRange(8, matchnumber * 4 - 2, 12).setValue(0)
    }
}

// Clears bets on match sheets for given matchnumber
function clearMatchSheets(matchnumber) {
    var adminmatchsheet = admin.getSheetByName('M00' + matchnumber)
    var viewmatchsheet = view.getSheetByName('Match ' + matchnumber)

    adminmatchsheet.getRange(3, 3, numplayers, 12).clearContent()

    adminmatchsheet.getRange('S17:T18').clearContent()
    viewmatchsheet.getRange('S17:T18').clearContent()
}

// Calculates consistency based on data in view.history
// Can be configured to recalculate for all matches if tech.h16 is set to zero
function consistencyCalc() {
    var stats = admin.getSheetByName("Stats")
    var history = view.getSheetByName("History")
    var numplayers = numPlayers()

    var playernames = tech.getRange(3, 1, numplayers).getValues()
    var participations = tech.getRange(3, 19, numplayers).getValues()
    var wincounts = tech.getRange(3, 20, numplayers).getValues()

    var numrounds = tech.getRange("h14").getValue()
    numrounds -= tech.getRange("h16").getValue()

    if (numrounds == 0) {
        return
    }

    var resultarrays = getResultArrays(numrounds)

    for (var i = 3; i < numplayers + 3; i++) {
        processPlayer(i)
    }

    tech.getRange("h16").setValue(tech.getRange("h14").getValue())

    function processPlayer(playerrow) {
        var playername = playernames[playerrow - 3]
        var participation = participations[playerrow - 3]
        var wincount = wincounts[playerrow - 3]
        var participated = false

        for (var j = 0; j < resultarrays.length; j++) {
            var resultarray = resultarrays[j]

            for (var k = 0; k < resultarray.length; k++) {
                var resultplayername = resultarray[k][0]
                if (resultplayername != playername) {
                    continue
                }

                var profit = resultarray[k][2]
                if (profit == 0) {
                    continue
                }

                // Player won or lost points
                participated = true

                participation++
                if (profit > 0) {
                    wincount++
                }
            }
        }

        if (!participated) {
            return
        }

        tech.getRange(playerrow, 19).setValue(participation)
        tech.getRange(playerrow, 20).setValue(wincount)

        if (participation >= 5) {
            stats.getRange(playerrow, 5).setValue(wincount / participation)
        }
    }

    function getResultArrays(numrounds) {
        var resultarrays = [];
        for (var round = 1; round <= numrounds; round++) {
            resultarrays.push(history.getRange(13, 1 + 4 * (numrounds - round), numplayers, 3).getValues())
        }

        return resultarrays
    }
}

// Restores points based on archived sheet view."[PlayerA] vs [PlayerB]"
function revertMatch(sheetname) {
    var archivedsheet = view.getSheetByName(sheetname)

    for (var i = 3; i < numplayers + 3; i++) {
        var playername = archivedsheet.getRange(i, 1).getValue()

        for (var a = 3; a < numplayers + 3; a++) {
            if (tech.getRange(a, 1).getValue() == playername) {
                tech.getRange(a, 17).setValue(tech.getRange(a, 17).getValue() - archivedsheet.getRange(i, 14).getValue())
                break
            }
        }
    }
}

function donateToAll(amount) {
    var donationsin = tech.getRange(3, 4, numplayers).getValues()

    for (var i = 0; i < donationsin.length; i++) {
        tech.getRange(3 + i, 4).setValue(donationsin[i][0] += amount)
    }
}

// Prepares new event
//
// Manual actions if a new admin sheet was created:
// - Duplicate reference sheet
// - Duplicate player template
// - Duplicate view template
// - Replace all instances of current admin ID in new reference sheet AND new view sheet with new admin ID
// - Replace all instances of current reference ID in new player template with new reference ID
// - Replace IDs in tech.o9 - o13
// - Run below script
function prepareNewEvent(eventname) {
    var viewtemplate = SpreadsheetApp.openById(tech.getRange("o13").getValue())
    var newview = viewtemplate.copy(eventname)
    var newviewid = newview.getId()
    tech.getRange("o7").setValue(newviewid)

    // Update player template
    var docslink = '=HYPERLINK("https://docs.google.com/spreadsheets/d/'
    var playertemplate = SpreadsheetApp.openById(tech.getRange("o9").getValue()).getSheetByName("Sheet1")
    playertemplate.getRange("i24").setValue(docslink + newviewid + '", "Standings")')

    var matchsheetids = []
    for (var i = 1; i < 6; i++) {
        matchsheetids.push(newview.getSheetByName("Match " + i).getSheetId())
    }
    playertemplate.getRange(21, 1).setValue(docslink + newviewid + '/edit#gid=' + matchsheetids[0] + '", "View bets")')
    playertemplate.getRange(21, 5).setValue(docslink + newviewid + '/edit#gid=' + matchsheetids[1] + '", "View bets")')
    playertemplate.getRange(21, 9).setValue(docslink + newviewid + '/edit#gid=' + matchsheetids[2] + '", "View bets")')
    playertemplate.getRange(21, 13).setValue(docslink + newviewid + '/edit#gid=' + matchsheetids[3] + '", "View bets")')
    playertemplate.getRange(21, 17).setValue(docslink + newviewid + '/edit#gid=' + matchsheetids[4] + '", "View bets")')

    resetAdmin(true)
}

// Resets the event, optionally removes the links to player sheets
function resetEvent(removePlayers) {
    resetView()
    resetAdmin(removePlayers)
}

// Resets admin spreadsheet
// Existing player sheets are no longer linked to the system but not deleted
function resetAdmin(removePlayers) {
    var stats = admin.getSheetByName("Stats")
    stats.getRange(3, 5, 1000).clearContent()

    tech.getRange(3, 1, 1000).clearContent()
    tech.getRange(3, 3, 1000, 3).setValue(0)
    
    if (removePlayers) {
        tech.getRange(3, 6, 1000).clearContent()
    }
    else {
        for (var i = 1; i < 6; i++) {
            resetBets(i)
        }
    }
    
    tech.getRange(3, 7, 1000).setValue(-1)
    tech.getRange(3, 17, 1000).setValue(0)
    tech.getRange(3, 18, 1000).clearContent()
    tech.getRange(3, 19, 1000, 2).setValue(0)

    tech.getRange("h9").setValue(0)
    tech.getRange("h14").setValue(0)
    tech.getRange("h16").setValue(0)
    tech.getRange(18, 9, 40, 7).clearContent()
    tech.getRange(3, 14, 10).setValue("-")
    tech.getRange("o3").clearContent()
    tech.getRange("o5").clearContent()
    tech.getRange(7, 9, 1, 5).setValue('=HYPERLINK(""; "Not yet scheduled")')
    tech.getRange(3, 9, 2, 5).clearContent()
    tech.getRange(5, 9, 1, 5).setValue("CLOSED")

    admin.getSheetByName("M001").getRange(3, 3, 1000, 12).clearContent()
    admin.getSheetByName("M002").getRange(3, 3, 1000, 12).clearContent()
    admin.getSheetByName("M003").getRange(3, 3, 1000, 12).clearContent()
    admin.getSheetByName("M004").getRange(3, 3, 1000, 12).clearContent()
    admin.getSheetByName("M005").getRange(3, 3, 1000, 12).clearContent()
}

// Resets view spreadsheet
function resetView() {
    var history = view.getSheetByName("History")
    var roundcount = tech.getRange("H14").getValue()

    if (roundcount == 0) {
        return
    }
    
    history.deleteColumns(4, 4 * roundcount)
    history.getRange("a1").setValue("Round 0")
    history.getRange(4, 2, 7).clearContent()
    history.getRange(13, 1, 1000, 3).clearContent()

    // Delete archived match sheets (name contains " vs ")
    var viewsheets = view.getSheets()
    for (var i = 0; i < viewsheets.length; i++) {
        if (viewsheets[i].getName().indexOf(" vs ") != -1) {
            view.deleteSheet(viewsheets[i])
        }
    }
}
