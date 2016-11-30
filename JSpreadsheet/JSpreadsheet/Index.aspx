<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Index.aspx.cs" Inherits="JSpreadsheet.Index" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <script src="scripts/SpreadSheet.js"></script>
    <link href="Styles/SpreadSheet.css" rel="stylesheet" />
</head>
<body onload="createSpreadsheet()">
    <form id="form1" runat="server">
    <div>
<%--            <input id="btnBuild" type="button" value="Build" onclick="createSpreadsheet();" />
        <br /><br />
        <input type="text" id="txtRows" onkeypress="return editTextNum(event);" value="20" />
        <input type="text" id="txtColumns" onkeypress="return editTextNum(event);" value="10" />
        <br /><br />--%>
        
        <div id="SpreadsheetTable">
        </div>
    </div>
    </form>

    <br />

    <label id="formulaLabel">Enter Data for currently selected cell.</label>
    <input type="text" id="formulaBox" onkeypress="return editTextNum(event);"  />

    

    <br />
    <input id="btnClear" type="button" value="Clear Spreadsheet and Formulas" onclick="clearTable();" />
    <input id="btnSave" type="button" value="Save Spreadsheet to local storage" onclick="savetoLocalStorage();" /> 
    <input id="btnLoad" type="button" value="Load Spreadsheet from local storage" onclick="loadFromLocalStorage();" />

</body>
</html>


