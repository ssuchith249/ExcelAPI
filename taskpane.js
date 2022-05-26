/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

/* 
    Authors: Raghavender Hariharan, Singuri Suchith, Rohit Bhrugumalla, Ujwala Parupudi
    Team: Platform Technology Group
    Description: This code is used to develop an Excel Add-In that lets the user upload their 
    worksheets into ION via IMS (ION V2 Messaging Service) or into Infor Datalake via Data fabric 
    ingestion APIs. Users can also retrieve data from Infor Datalake into Excel Worksheets.
    Taskpane Folder: Contains the HTML and Javascript code that handles all the UI and functionalities
    for the Excel Add-In. 
*/

var _dlg;             // var for Dialog box
var error_var = 0;    // var to check if Log Sheet exists
var access_token;     // var for storing Bearer token
var no_of_rows;       // var for storing rows in Log Sheet
var tenant;           // var for storing tenant value
var lid;              // var for storing logical id value
var color;            // var for color used in log sheet

// Load Fetch Library
const fetch = (...args) => import('node-fetch').then(({ default: fetch }) => fetch(...args));

// Window OnLoad
window.addEventListener('load', addSheet);

// Add Status_Overview Sheet
export async function addSheet() {
  try {
    await Excel.run(async (context) => {
      // Add log sheet
      let sheets = context.workbook.worksheets;
      let log_sheet = sheets.add("Sheet_Overview");
      log_sheet.load("name, position");
      await context.sync();

      // Add Log Sheet Headers
      log_sheet = sheets.getItem("Sheet_Overview");
      let headers = [
        ["Sheet_Name", "Object_Schema", "Size(in Bytes)", "No of Rows", "Date", "Time", "Status", "Error-Message"],
      ];
      let range = log_sheet.getRange("A1:H1");
      range.values = headers;
      range.format.autofitColumns();
      let header_range = log_sheet.getRange("A1:H1");
      header_range.format.fill.color = "#4472C4";
      header_range.format.font.color = "white";
      await context.sync();
      error_var = -1;
    });
  } catch (error) {
    error_var = 1;
    if (error.code == "InvalidOperationInCellEditMode") {
      // Modal for Editing Mode Error
      var myModal = new bootstrap.Modal(document.getElementById("myModal"));
      document.getElementById("modalHeading").innerHTML = "Load Add-In";
      document.getElementById("modalText").innerHTML = "Excel cell in Edit Mode. Please Exit Edit mode by using the Enter or Tab keys, or by selecting another cell, and then load the Add-In again.";
      myModal.show();
    }

  }
}

// Load Worksheets Dropdown 
export async function loadDropdown() {
  try {
    await Excel.run(async (context) => {
      var list = document.getElementById("sheetDropdown");

      list.length = 1;
      let sheets = context.workbook.worksheets;
      sheets.load("name");

      return context.sync().then(async function () {
        for (var k = 0; k < sheets.items.length; k++) {
          var opt = sheets.items[k].name;
          if (opt == "Sheet_Overview")
            continue
          var text = document.createTextNode(opt);
          var option = document.createElement("option");
          option.appendChild(text);
          list.appendChild(option);
        }
      });

    });
  } catch (error) {
    console.error(error);
  }
}

// Check which Option has been chosen
function radioButtonCheck() {
  const radioButtons = document.querySelectorAll('input[name="inlineRadioOptions"]');
  let selectedSize;
  for (const radioButton of radioButtons) {
    if (radioButton.checked) {
      selectedSize = radioButton.value;
      loadDropdown();
      // To IMS
      if (selectedSize === "option1") {
        document.getElementById("retrieve_select").style.display = "none";
        document.getElementById("sheet_select").style.display = "block";

        // Text box
        document.getElementById('textbox').style.display = "block";
      }
      // To Datalake
      else if (selectedSize === "option2") {
        document.getElementById("retrieve_select").style.display = "none";
        document.getElementById("sheet_select").style.display = "block";

        // Text box
        document.getElementById('textbox').style.display = "block";
      }
      // Retrieve Data
      else if (selectedSize === "option3") {
        document.getElementById("sheet_select").style.display = "none";
        document.getElementById("retrieve_select").style.display = "block";
      }
      break;
    }
  }
}


Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Set Style
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // Authenticate Button
    document.getElementById("signIn").onclick = signIn;

    // Instructions Button
    document.getElementById("instructions").onclick = openInstructions;

    // Refresh Sheet Dropdown
    document.getElementById('refreshDropdown').addEventListener('click', loadDropdown);

    // Send Data Button
    document.getElementById("sendData").onclick = sendData;

    // Retrieve Data Button
    document.getElementById("run").onclick = run;

    // Check Radio button
    document.getElementById("inlineRadio1").addEventListener('click', radioButtonCheck);
    document.getElementById("inlineRadio2").addEventListener('click', radioButtonCheck);
    document.getElementById("inlineRadio3").addEventListener('click', radioButtonCheck);

  }
});

// Open Instructions.html
function openInstructions() {
  Excel.run(context => {
    // // sync the context to run the previous API call, and return.
    Office.context.ui.displayDialogAsync('https://localhost:3000/instructions.html',
      // change these to your preference
      { height: 70, width: 45 },

      function (asyncResult) {

        // note _dlg is globally defined
        _dlg = asyncResult.value;

        _dlg.addEventHandler(Office.EventType.DialogMessageReceived,
          processDialogCallback);
      }
    );
    return context.sync();

  });
}

// Sign In Button Click
function signIn() {
  Excel.run(context => {

    // Change color of Authenticate Button
    document.getElementById('signIn').classList.remove('btn-success');
    document.getElementById('signIn').classList.add('btn-primary');

    // Set Tenant Name to h6 tag
    document.getElementById('tenant_name').innerHTML = "";

    // Hide all the option when Authenticate button is pressed
    document.getElementById("option_list").style.display = "none";
    document.getElementById("sheet_select").style.display = "none";
    document.getElementById("retrieve_select").style.display = "none";

    // // sync the context to run the previous API call, and return.
    Office.context.ui.displayDialogAsync('https://localhost:3000/commands.html',
      // change these to your preference
      { height: 70, width: 45 },

      function (asyncResult) {

        // note _dlg is globally defined
        _dlg = asyncResult.value;

        _dlg.addEventHandler(Office.EventType.DialogMessageReceived,
          processDialogCallback);
      }
    );
    return context.sync();

  });
}

// Process Message received from Dialog 
function processDialogCallback(arg) {
  var messageFromDialog = JSON.parse(arg.message);
  if (messageFromDialog.messageType === "token") {
    access_token = messageFromDialog.access_token;
    if (typeof (access_token) == 'undefined') {
      _dlg.close();
      var myModal = new bootstrap.Modal(document.getElementById("myModal"));
      document.getElementById("modalHeading").innerHTML = "Sign In";
      document.getElementById("modalText").innerHTML = "Failed to Sign In.";
      myModal.show();
    }
    else {
      _dlg.close();

      // Change color of Authenticate Button
      document.getElementById('signIn').classList.remove('btn-primary');
      document.getElementById('signIn').classList.add('btn-success');

      // Display send and recieve buttons and make checked attribute as false
      var opts = document.getElementsByClassName('form-check');
      document.getElementById('inlineRadio1').checked = false;
      document.getElementById('inlineRadio2').checked = false;
      document.getElementById('inlineRadio3').checked = false;
      opts[0].style.display = 'block';

      // Set Tenant Name to h6 tag
      document.getElementById('tenant_name').innerHTML = `Tenant: ${tenant}`;
    }
  }

  else if (messageFromDialog.messageType === "tenant") {
    tenant = messageFromDialog.tenant_name;
  }
}

// Add Logs in Sheet Overview
export async function logSheet(sheet_name, schema_name, size_sheet, no_of_rows, date_time, currTime, status, error_msg, sheet_color) {
  try {
    await Excel.run(async (context) => {
      // Add logs in Sheet_Overview
      let sheets = context.workbook.worksheets;
      let log_sheet = sheets.getItem("Sheet_Overview");

      let range = log_sheet.getUsedRange();
      range.load("values");
      await context.sync();

      var sheetOverview_row = (range.values).length + 1;
      let row_range = log_sheet.getRange(`A${sheetOverview_row}:H${sheetOverview_row}`);
      if (no_of_rows == 1)
        no_of_rows = 0; // Used to assign number of rows if sheet is empty
      row_range.values = [
        [sheet_name, schema_name, size_sheet, no_of_rows, date_time, currTime, status, error_msg],
      ];
      row_range.format.autofitColumns();

      //To set color to the error messsages in Sheet_Overview
      let color_range = log_sheet.getRange(`G${sheetOverview_row}:H${sheetOverview_row}`);
      color_range.format.font.color = sheet_color;

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

// Check if schema exists in the tenant or not
async function getObjectList() {
  try {
    var myHeaders = new Headers();
    myHeaders.append("Authorization", `Bearer ${access_token}`);

    var requestOptions = {
      method: 'GET',
      headers: myHeaders,
      redirect: 'follow'
    };

    const response = await fetch(`https://mingle-ionapi.inforcloudsuite.com/${tenant}/IONSERVICES/datacatalog/v1/object/list`, requestOptions);
    const obj = await response.json();
    return obj;
  } catch (error) {
    console.log(error);
  }
}


// Extract the final data and send it
export async function extractAndSend(json_text, schema_name, sheet_name) {
  try {
    await Excel.run(async (context) => {
      var empty_str = "";
      no_of_rows = 0;
      // Loop through all the sheets
      for (var i = 1; i < json_text.length; i++) {
        no_of_rows += 1;
        for (var j = 0; j < json_text[i].length; j++) {
          empty_str += json_text[i][j];
          if (j == json_text[i].length - 1)
            continue
          empty_str += ',';
        }
        if (i == json_text.length - 1)
          continue
        empty_str += '\n';
      }

      if (empty_str == "") {
        color = "Red";
        await logSheet(sheet_name, "-", 0, 0, new Date().toLocaleDateString(), new Date().toLocaleTimeString(), "", "Data Not Found", color);
        return;
      }
      // Store Log Sheet Details
      var size_sheet = byteCount(empty_str);
      var date_time = "";
      var status = "";
      var error_msg = "";

      const radioButtons = document.querySelectorAll('input[name="inlineRadioOptions"]');
      let selectedSize;
      for (const radioButton of radioButtons) {
        if (radioButton.checked) {
          selectedSize = radioButton.value;
          // Send via IMS
          if (selectedSize === "option1") {

            //Size of Data Check
            if (size_sheet > 5000000) {
              date_time = new Date().toLocaleDateString();
              error_msg = "Too Large File cannot upload";
              color = "Red";
            }

            else {
              var data = JSON.stringify({
                "documentName": schema_name,
                "messageId": schema_name + Math.floor(Math.random() * 1000001).toString(),
                "fromLogicalId": `lid://${lid}`,
                "toLogicalId": "lid://default",
                "document": {
                  "value": empty_str,
                  "encoding": "NONE",
                  "characterSet": "UTF-8"
                }
              });

              var config = {
                method: 'post',
                url: `https://mingle-ionapi.inforcloudsuite.com/${tenant}/CustomerApi/EXCELWrapperAPI/v2/message`,
                headers: {
                  'Authorization': `Bearer ${access_token}`,
                  'Content-Type': 'application/json',
                  'cache-control': 'no-cache'
                },
                data: data
              };

              date_time = new Date().toLocaleDateString();
              var axios = require('axios');
              color = "Green";
              var result = await axios(config).catch(function (error) {
                if (error.response) {
                  status = error.response.status;
                  color = "Red";
                  if (status == 401)
                    error_msg = error.response.data["error"];
                  else
                    error_msg = error.response.data["errors"];
                }
              });

              if (result !== undefined) {
                status = `${result.data["code"]}. ${result.data["message"]}`;
                error_msg = "";
              }
            }

            // Add logs in Sheet_Overview
            await logSheet(sheet_name, schema_name, size_sheet, no_of_rows, date_time, new Date().toLocaleTimeString(), status, error_msg, color);
          }
          // Send to DataLake
          else if (selectedSize === "option2") {
            var response = await getObjectList();
            if (response.objects.find(({ name }) => name === schema_name) == undefined) {
              color = "Red";
              date_time = new Date().toLocaleDateString();
              // Add logs in Sheet_Overview
              await logSheet(sheet_name, schema_name, size_sheet, no_of_rows, date_time, new Date().toLocaleTimeString(), "Data Not Sent.", "Object Schema Does not Exist", color);
            }

            else {
              // Perform Pako Deflate
              const pako = require('pako');
              var fileAsArray = pako.deflate(empty_str, { to: 'string' });
              const compressedFile = fileAsArray.buffer;
              const dataToUpload = new Blob([compressedFile], { type: 'text/csv;charset=utf-8' });
              color = "Green";
              var axios = require('axios');
              var FormData = require('form-data');
              //var fs = require('fs');
              var data = new FormData();
              data.append('dl_document_name', schema_name);
              data.append('dl_from_logical_id', `lid://${lid}`);
              data.append('file', dataToUpload);

              var config = {
                method: 'post',
                url: `https://mingle-ionapi.inforcloudsuite.com/${tenant}/DATAFABRIC/ingestion/v1/dataobjects`,
                headers: {
                  'Authorization': `Bearer ${access_token}`
                },
                data: data
              };

              date_time = new Date().toLocaleDateString();
              var axios = require('axios');
              var result = await axios(config).catch(function (error) {
                color = "Red";
                console.error(error);
                if (error.response) {
                  status = error.response.status;
                  if (status == 401)
                    error_msg = error.response.data["error"];
                  else if (status = 400) {
                    error_msg = error.response.data["errors"][0].message + " (dl_from_logical_id refers to the Logical ID being entered while sending the data)";
                  }
                  else
                    error_msg = error.response.data["errors"];
                }
              });
              if (result !== undefined) {
                status = `${result.status}. Published Successfully`;
                error_msg = "";
              }
              if (dataToUpload.size > 5000000) {
                error_msg += "Warning:The compressed file is above 5MB";
                color = "Orange";
              }

              await logSheet(sheet_name, schema_name, size_sheet, no_of_rows, date_time, new Date().toLocaleTimeString(), status, error_msg, color);

            }
          }
          break;
        }
      }
    });
  }
  catch (error) {
    console.error(error);
  }
}

// Split the Sheet Data into equal chunks
export async function splitDataToChunks(name) {
  try {
    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getItem(name);
      sheet.activate();
      let range = sheet.getUsedRange();
      range.load("address");
      await context.sync();

      // Extract rows and column from range address
      var range_str = range.address;
      range_str = range_str.slice(range_str.indexOf("!") + 1);


      var arr = range_str.split(':');
      //Get start Row and Column
      var startCol = arr[0].replace(/[0-9]/g, '');
      var startRow = arr[0].replace(/\D/g, '');

      // If Data is Empty
      if ((arr.length == 1) && (startCol + startRow == "A1")) {
        color = "Red";
        await logSheet(name, "-", 0, 0, new Date().toLocaleDateString(), new Date().toLocaleTimeString(), "", "Data Not Found", color);
      }

      else {
        // Get End Row and Column
        var endCol = arr[1].replace(/[0-9]/g, '');
        var endRow = arr[1].replace(/\D/g, '');
        startRow = parseInt(startRow);

        var number = endRow;
        var chunk_array = new Array(Math.floor(number / 10000)).fill(10000).concat(number % 10000);

        // Var to Hold Data
        var json_text = [];
        var rowend = 0;
        for (var j = 0; j < chunk_array.length; j++) {
          rowend += chunk_array[j];
          let end = endCol + rowend;
          end = `${startCol}${startRow}: ${end}`;
          let range = sheet.getRange(end);
          range.load("values");
          await context.sync();
          json_text.push(range.values);
          startRow = rowend + 1;
        }

        var finalDataToSend = [];

        for (var i = 0; i < json_text.length; i++) {
          for (var j = 0; j < json_text[i].length; j++) {
            finalDataToSend.push(json_text[i][j]);
          }
        }

        await extractAndSend(finalDataToSend, finalDataToSend[0][0], name);
      }
    });
  }
  catch (error) {
    console.error(error);
  }
}


// Send Data Button Click
export async function sendData() {
  try {
    await Excel.run(async (context) => {

      // Check logical ID
      lid = document.getElementById("lid").value;

      const radioButtons = document.querySelectorAll('input[name="inlineRadioOptions"]');
      let selectedSize;
      for (const radioButton of radioButtons) {
        if (radioButton.checked) {
          selectedSize = radioButton.value;
          if (lid == "") {
            var myModal = new bootstrap.Modal(document.getElementById("myModal"));
            document.getElementById("modalHeading").innerHTML = "Send Data";
            document.getElementById("modalText").innerHTML = "Please Enter Logical ID";
            myModal.show();
            return;
          }
        }
      }
      var myModal = new bootstrap.Modal(document.getElementById("myModal"));
      document.getElementById("modalHeading").innerHTML = "Send Data";
      document.getElementById("modalText").innerHTML = "Sending Data. Please Wait";
      myModal.show();
      var select = document.getElementById('sheetDropdown');
      var text = select.options[select.selectedIndex].text;

      if (text == "ALL") {
        let sheets = context.workbook.worksheets;
        sheets.load("name");

        // Check is Log Sheet exists or not
        if (error_var == 0) {
          addSheet();
        }

        return context.sync().then(async function () {

          for (var k = 0; k < sheets.items.length - 1; k++) {
            let sheet = context.workbook.worksheets.getItem(sheets.items[k].name);
            sheet.activate();
            sheet.load("name");
            await context.sync();

            if (sheet.name === 'Sheet_Overview') {
              break;
            }

            // Split the data into chunks
            await splitDataToChunks(sheet.name);
          }

          // If All Sheets being sent, then make Sheet_Overview active
          let sheet = context.workbook.worksheets.getItem('Sheet_Overview');
          sheet.activate();
          document.getElementById("modalText").innerHTML = "Sent Data. Please check Sheet_Overview for more details.";
        }).catch(e => {
          console.log(e);
        });
      }

      else {
        // Check is Log Sheet exists or not
        if (error_var == 0) {
          addSheet();
        }

        let sheet = context.workbook.worksheets.getItem(text);
        sheet.activate();

        sheet.load("name");
        await context.sync();

        // Split the data into chunks
        await splitDataToChunks(sheet.name);

        document.getElementById("modalText").innerHTML = "Sent Data. Please check Sheet_Overview for more details.";

        // If Single Sheet is being sent, then make Sheet_Overview active
        let sheet_overview = context.workbook.worksheets.getItem('Sheet_Overview');
        sheet_overview.activate();
      }

    });
  } catch (error) {
    console.error(error);
  }
}

function byteCount(s) {
  return encodeURI(s).split(/%..|./).length - 1;
}


// Data Retrieval Process

// Get Query ID
async function getQueryId(fname) {
  try {
    var myHeaders = new Headers();
    myHeaders.append("Authorization", `Bearer ${access_token}`);
    myHeaders.append("Content-Type", "text/plain");

    var raw = fname;

    var requestOptions = {
      method: 'POST',
      headers: myHeaders,
      body: raw,
      redirect: 'follow'
    };
    const response = await fetch(`https://mingle-ionapi.inforcloudsuite.com/${tenant}/DATAFABRIC/compass/v2/jobs/`, requestOptions);
    const obj = await response.json();
    return obj;
  } catch (error) {
    console.error(error);
  }
}

// Check Status of Query
async function checkStatus(queryId) {
  try {
    var myHeaders = new Headers();
    myHeaders.append("Authorization", `Bearer ${access_token}`);

    var requestOptions = {
      method: 'GET',
      headers: myHeaders,
      redirect: 'follow'
    };

    const response = await fetch(`https://mingle-ionapi.inforcloudsuite.com/${tenant}/DATAFABRIC/compass/v2/jobs/${queryId}/status/`, requestOptions);
    const obj = await response.json();
    return obj;
  } catch (error) {
    console.error(error);
  }
}

// Get data from Result
async function getResult(queryId) {
  try {
    var myHeaders = new Headers();
    myHeaders.append("Authorization", `Bearer ${access_token}`);

    var requestOptions = {
      method: 'GET',
      headers: myHeaders,
      redirect: 'follow'
    };

    const response = await fetch(`https://mingle-ionapi.inforcloudsuite.com/${tenant}/DATAFABRIC/compass/v2/jobs/${queryId}/result?limit=100000&offset=0`, requestOptions);
    const obj = await response.text();
    return obj;
  } catch (error) {
    console.log(error);
  }
}

// Rename worksheet with provided Sheet name 
export async function renameWorksheet(sheet) {
  try {
    await Excel.run(async (context) => {
      var name = document.getElementById('sheet_name').value;
      sheet.name = name;
      await context.sync();
    });
  }
  catch (error) {
    console.error(error);
  }
}

// Get Column Name from Number
function printString(columnNumber) {
  // To store result (Excel column name)
  let columnName = [];

  while (columnNumber > 0) {
    // Find remainder
    let rem = columnNumber % 26;

    // If remainder is 0, then a
    // 'Z' must be there in output
    if (rem == 0) {
      columnName.push("Z");
      columnNumber = Math.floor(columnNumber / 26) - 1;
    }
    else // If remainder is non-zero
    {
      columnName.push(String.fromCharCode((rem - 1) + 'A'.charCodeAt(0)));
      columnNumber = Math.floor(columnNumber / 26);
    }
  }
  columnName = columnName.reverse().join("")
  return columnName;
}

// Function to cause delay for every iteration
function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// Get response data
export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      var name = document.getElementById('sheet_name').value;
      if (name.length > 31) {
        var myModal = new bootstrap.Modal(document.getElementById("myModal"));
        document.getElementById("modalHeading").innerHTML = "Retrieve Data";
        document.getElementById("modalText").innerHTML = "Please provide the sheet name which has less than 31 characters";
        myModal.show();
        return;
      }

      var myModal = new bootstrap.Modal(document.getElementById("myModal"));
      document.getElementById("modalHeading").innerHTML = "Retrieve Data";
      document.getElementById("modalText").innerHTML = "Retrieving Data. Please Wait";
      myModal.show();
      let fname = document.getElementById('fname').value;

      let sheet = context.workbook.worksheets.getActiveWorksheet();

      const csv_to_array = (data, delimiter = ',', omitFirstRow = false) =>
        data
          .slice(omitFirstRow ? data.indexOf('\n') + 1 : 0)
          .split('\n')
          .map(v => v.split(delimiter));

      var response = await getQueryId(fname);
      var queryId = response.queryId;
      var count = 60;
      while (count >= 0) {
        response = await checkStatus(queryId);
        if (response.status == "FINISHED") {
          response = await getResult(queryId);
          response = response.replace(/['"]+/g, '');

          var size_sheet = byteCount(response); // Get size of data
          let rowno = 0;
          let columnno = 0;

          let data = csv_to_array(response);
          data.pop();
          rowno = data.length;
          columnno = data[0].length;

          let columnname = printString(columnno);

          // For Loop to Split the Retrieval Process
          if (data.length <= 10000) {
            let end = columnname + rowno;
            end = `A1: ${end}`;
            let range = sheet.getRange(end);
            range.values = data;
            range.format.autofitColumns();
            await context.sync();
          }

          else {
            var index = 0;
            var chunk_size = 10000;
            var arrayLength = data.length;
            var tempArray = [];

            var rowstart = 1;
            var rowend = 0;
            columnno = data[0].length;

            let columnname = printString(columnno);

            for (index = 0; index < arrayLength; index += chunk_size) {
              tempArray.push(data.slice(index, index + chunk_size));
            }

            for (var j = 0; j < tempArray.length; j++) {
              rowend += tempArray[j].length;
              let end = columnname + rowend;
              end = `A${rowstart}: ${end}`;
              let range = sheet.getRange(end);
              range.values = tempArray[j];
              range.format.autofitColumns();
              await context.sync();
              rowstart = rowend + 1;
            }
          }

          await renameWorksheet(sheet);
          document.getElementById("modalText").innerHTML = "Retrieved Data Successfully.";

          // Log details into Sheet_Overview
          sheet.load('name');
          await context.sync();
          var sheet_name = sheet.name;
          var date_time = new Date().toLocaleDateString();
          var status = "";
          status = status + "Retrieved Data Successfully. Query Id: " + queryId;
          var error_msg = "";
          color = "Green";
          await logSheet(sheet_name, "-", size_sheet, rowno, date_time, new Date().toLocaleTimeString(), status, error_msg, color);
          break;
        }

        else if (response.status == "FAILED") {
          var date_time = new Date().toLocaleDateString();
          var status = response.status;
          status = status + ". Couldn't retrieve the data please check the query.";
          var error_msg = "Couldn't retrieve the data please check the query. Query Id: " + queryId;
          color = "Red";
          await logSheet("-", "-", 0, 0, date_time, new Date().toLocaleTimeString(), status, error_msg, color);
          document.getElementById("modalText").innerHTML = "Could not retrieve data ,please check Sheet_Overview for more details";
          break;
        }
        else if (count == 1 && response.status == "RUNNING") {
          var date_time = new Date().toLocaleDateString();
          var status = response.status;
          var error_msg = "Please Re-Run the Query. Query Id: " + queryId;
          color = "Orange";
          await logSheet("-", "-", 0, 0, date_time, new Date().toLocaleTimeString(), status, error_msg, color);
          document.getElementById("modalText").innerHTML = "Please Re-Run the Query.";
          break;
        }
        count--;
        await sleep(5000);
      }
    });
  } catch (error) {
    console.error(error);
  }
}
