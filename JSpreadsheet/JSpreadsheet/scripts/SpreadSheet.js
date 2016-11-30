// Dynamic HTML to build a spreadsheet table

var nCharsAllowed = "0123456789abcdefghijklmnopqrstuvwxyz=():";

var isIE = ((window.navigator.userAgent).indexOf("MSIE") != -1);
var isChrome = ((window.navigator.userAgent).indexOf("Chrome") != -1);
var isFF = ((window.navigator.userAgent).indexOf("Firefox") != -1);

//Selected Cell Variables
var prevSelect = "1_1";
var currSelect = "";

//Creating the selection area for formulas
var formulaStartRow = 0;
var formulaEndRow = 0;
var formulaStartColumn = 0;
var formulaEndColumn = 0;


// ********************
// button event handler
function createSpreadsheet() {
    //var rows = parseInt(document.getElementById("txtRows").value);
    //var columns = parseInt(document.getElementById("txtColumns").value);
    var rows = 20;
    var columns = 10;
    document.getElementById("SpreadsheetTable").innerHTML = buildTable(rows, columns);
}

// ***************************************************
// function builds the table based on rows and columns
function buildTable(rows, columns) {
    // start with the table declaration
    var divHTML = "<table border='1' cellpadding='0' cellspacing='0' class='TableClass'>";

    // next do the column header labels
    divHTML += "<tr><th></th>";

    for (var j = 0; j < columns; j++)
        divHTML += "<th>" + String.fromCharCode(j + 65) + "</th>";

    // now do the main table area
    for (var i = 1; i <= rows; i++) {
        divHTML += "<tr>";
        // ...first column of the current row (row label)
        divHTML += "<td id='" + i + "_0' class='BaseColumn'>" + i + "</td>";

        // ... the rest of the columns //ADD HIDDEN DIVS FOR THE FORMULA!
        for (var j = 1; j <= columns; j++)
            divHTML += "<div id='div" + i + "_" + j + "'" + " style='display:none;'><td id='" + i + "_" + j + "' class='AlphaColumn' onclick='clickCell(this)'></td><div>";

        // ...end of row
        divHTML += "</tr>";
    }

    // finally add the end of table tag
    divHTML += "</table>";

    return divHTML;
}




// *******************************************
// event handler fires when user clicks a cell
function clickCell(ref) {
    //var rcArray = ref.id.split('_');    
    // alert("You selected row " + rcArray[0] + " and column " + rcArray[1]);

    //Set Colour of previous id back to the rest of the grid
    document.getElementById(prevSelect).style.background = "#C3DAF9";

    //Set Colour of selected id
    document.getElementById(ref.id).style.background = "#00FFFF";
    prevSelect = ref.id;
    currSelect = ref.id;

    var selecVal = document.getElementById(currSelect).innerHTML;
    var selecDivVal = document.getElementById("div" + currSelect).innerHTML;

    document.getElementById("formulaBox").focus();

    //if div contains a formula, update formula box to show the formula.
    if (document.getElementById("div" + currSelect).innerHTML.substring(0, 5) == "SUM=(") {
        document.getElementById("formulaBox").value = selecDivVal;
    }
    else {
        document.getElementById("formulaBox").value = selecVal;
    }


}


// *********************************************************
// Cross browser code to filter data entry on numeric fields
function editTextNum(e) {
    if (isFF) {
        // get the key code pressed
        var key = e.which;

        // if this is a backspace then process
        if (key == 8)
            return true;
        else if (key == 13) {
            //alert("You pressed the enter key");
            editCurrentCell();
            calculateAllCells();
            return true;
        }
        else {
            // check to see if this is an allowable key
            if (!nCharOK(key))
                return false;
            else
                return true;
        }
    }
    else {
        // IE uses the .keyCode
        if (isIE) {
            if (window.event.keyCode == 13) {
                //alert("You pressed the enter key");
                //Calculate all the cells and re-update them.
                editCurrentCell();
                calculateAllCells();
            }
            else if (!nCharOK(window.event.keyCode)) {
                window.event.keyCode = null;
            }

            return true;
        }
        else {
            // Chrome and Safari use returnValue
            if (window.event.keyCode == 13) {
                //alert("You pressed the enter key");
                editCurrentCell();
                calculateAllCells();
            }
            else if (window.event.keyCode != 13) {
                //Do nothing, process the event.
            }


            else if (!nCharOK(window.event.keyCode)) {
                window.event.returnValue = null;
            }

            return true;
        }
    }
}

// *************************************************************
// filter the currently entered character to see that it is part
// of the acceptable numeric (integer) character set
function nCharOK(c) {
    var ch = (String.fromCharCode(c));
    ch = ch.toUpperCase();

    if (nCharsAllowed.indexOf(ch) != -1)
        return true;
    else
        return false;
}


//Function to set content of currently selected cell to the value of the textbox.
function editCurrentCell() {
    //if its a formula, then add the formula to the div
    if (formulaBox.value.substring(0, 5) == "SUM=(") {
        document.getElementById("div" + currSelect).innerHTML = formulaBox.value;
    }
    else {
        document.getElementById(currSelect).innerHTML = formulaBox.value;
    }

}

function getSum(StartColumn, EndColumn, StartRow, EndRow) {
    //Go through all of the cells in the grid
    var sum = 0;
    for (var i = StartColumn; i <= EndColumn; i++) {
        for (var j = StartRow; j <= EndRow; j++) {
            if (!isNaN(parseFloat(document.getElementById(j + "_" + i).innerHTML))) {
                var numstring = document.getElementById(j + "_" + i).innerHTML;
                sum += parseFloat(numstring);
            }

        }
    }
    return sum;
}

//Function to loop through all of the cells in the spreadsheet and perform calculations
//On the cells, if calculations need to be performed.
function calculateAllCells() {

    //Determine if the cell contains a literal text string (numeric or alpha),  or a formula 
    //If you detect a formula, you must perform the designated (sum) calculation and output the result.
    //Go through all of the cells



    //Traverse all columns
    for (var i = 1; i <= 9; i++) {
        //Do all Rows in the Column
        for (var j = 1; j <= 19; j++) {
            //Check if it's a formula or not
            //SUM=(Ax:By)

            if (document.getElementById("div" + j + "_" + i).innerHTML.substring(0, 5) == "SUM=(") {
                //pull the start row, end row, startcol and endcol

                //if forumula is (ax:bx)
                if (document.getElementById("div" + j + "_" + i).innerHTML.length == 11) {
                    var startcol = document.getElementById("div" + j + "_" + i).innerHTML.substring(5, 6);
                    var startrw = document.getElementById("div" + j + "_" + i).innerHTML.substring(6, 7);
                    var endcol = document.getElementById("div" + j + "_" + i).innerHTML.substring(8, 9);
                    var endrw = document.getElementById("div" + j + "_" + i).innerHTML.substring(9, 10);
                }

                //if forumula is (ax:bxx)
                if (document.getElementById("div" + j + "_" + i).innerHTML.length == 12 &&
                    document.getElementById("div" + j + "_" + i).innerHTML.substring(7, 8) == ":") {
                    var startcol = document.getElementById("div" + j + "_" + i).innerHTML.substring(5, 6);
                    var startrw = document.getElementById("div" + j + "_" + i).innerHTML.substring(6, 7);
                    var endcol = document.getElementById("div" + j + "_" + i).innerHTML.substring(8, 9);
                    var endrw = document.getElementById("div" + j + "_" + i).innerHTML.substring(9, 11);
                }

                //if formula is (axx:bx)
                if (document.getElementById("div" + j + "_" + i).innerHTML.length == 12 &&
                document.getElementById("div" + j + "_" + i).innerHTML.substring(8, 9) == ":") {
                    var startcol = document.getElementById("div" + j + "_" + i).innerHTML.substring(5, 6);
                    var startrw = document.getElementById("div" + j + "_" + i).innerHTML.substring(6, 8);
                    var endcol = document.getElementById("div" + j + "_" + i).innerHTML.substring(9, 10);
                    var endrw = document.getElementById("div" + j + "_" + i).innerHTML.substring(10, 11);

                }

                //if formula is (axx:bxx)
                if (document.getElementById("div" + j + "_" + i).innerHTML.length == 13) {
                    var startcol = document.getElementById("div" + j + "_" + i).innerHTML.substring(5, 6);
                    var startrw = document.getElementById("div" + j + "_" + i).innerHTML.substring(6, 8);
                    var endcol = document.getElementById("div" + j + "_" + i).innerHTML.substring(9, 10);
                    var endrw = document.getElementById("div" + j + "_" + i).innerHTML.substring(10, 12);
                }

                var startcolumn;
                switch (startcol) {
                    case 'A':
                        startcolumn = 1;
                        break;
                    case 'B':
                        startcolumn = 2;
                        break;
                    case 'C':
                        startcolumn = 3;
                        break;
                    case 'D':
                        startcolumn = 4;
                        break;
                    case 'E':
                        startcolumn = 5;
                        break;
                    case 'F':
                        startcolumn = 6;
                        break;
                    case 'G':
                        startcolumn = 7;
                        break;
                    case 'H':
                        startcolumn = 8;
                        break;
                    case 'I':
                        startcolumn = 9;
                        break;
                    case 'J':
                        startcolumn = 10;
                        break;
                }
                var endcolumn;
                switch (endcol) {
                    case 'A':
                        endcolumn = 1;
                        break;
                    case 'B':
                        endcolumn = 2;
                        break;
                    case 'C':
                        endcolumn = 3;
                        break;
                    case 'D':
                        endcolumn = 4;
                        break;
                    case 'E':
                        endcolumn = 5;
                        break;
                    case 'F':
                        endcolumn = 6;
                        break;
                    case 'G':
                        endcolumn = 7;
                        break;
                    case 'H':
                        endcolumn = 8;
                        break;
                    case 'I':
                        endcolumn = 9;
                        break;
                    case 'J':
                        endcolumn = 10;
                        break;
                }
                var startrow = startrw;
                var endrow = endrw;

                document.getElementById(j + "_" + i).innerHTML = getSum(startcolumn, endcolumn, startrow, endrow);

            }


        }
    }


}



// Clear all literals and formulas from the spreadsheet
function clearTable() {
    //Clear the Spreadsheet
    document.getElementById("SpreadsheetTable").innerHTML = buildTable(20, 10);

    //Clear Formulas
    document.getElementById("formulaBox").value = "";
}



//Function to save the spreadsheet to JSON
function savetoLocalStorage() {
    if (typeof (Storage) !== "undefined") {
        // Code for localStorage/sessionStorage.
        //Set Colour of previous id back to the rest of the grid
        document.getElementById(prevSelect).style.background = "#C3DAF9";

        //Spreadsheet Data to save to JSON / file
        var tableData = document.getElementById("SpreadsheetTable").innerHTML;
        var JSONTable = JSON.stringify(tableData);
        localStorage.setItem("spreadsheet", JSONTable);
    } else {
        // Sorry! No Web Storage support..
    }
}

//Function to load from localStorage
function loadFromLocalStorage() {
    var tableFromStorage = localStorage.getItem("spreadsheet");
    var htmlTable = JSON.parse(tableFromStorage);
    document.getElementById("SpreadsheetTable").innerHTML = htmlTable;
}