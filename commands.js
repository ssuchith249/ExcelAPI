/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

/* 
    Authors: Raghavender Hariharan, Singuri Suchith, Rohit Bhrugumalla, Ujwala Parupudi
    Team: Platform Technology Group
    Description: This code is used to develop an Excel Add-In that lets the user upload their 
    worksheets into ION via IMS (ION V2 Messaging Service) or into Infor Datalake via Data fabric 
    ingestion APIs. Users can also retrieve data from Infor Datalake into Excel Worksheets.
    Commands Folder: Contains the HTML and Javascript code that handles the authorization for the Add-In.
*/


var code;        // var for Auth Code  
var auth_obj;    // var for storing Auth Object
var token;       // var for storing Bearer token
var flag = 0;    // Check for same ionapi file
var closebtns;
var items;

// OnLoad Extract Auth Code and URL
window.addEventListener('load', () => {

  // Load Dropdown with List of ION API Files
  Object.keys(localStorage).sort().reverse().forEach(function (key) {
    if (key.includes('ionAPI')) {
      var opt = localStorage.getItem(key);
      var name = JSON.parse(opt).cn;

      var ul = document.getElementById("dynamic-list");
      var li = document.createElement("li");
      li.setAttribute('id', name);
      li.innerHTML = `<button type="button" title="Click to Sign In" class="btn btn-link btn-sm">${name}</button>
    <span class="close"><svg xmlns="http://www.w3.org/2000/svg" width="16" height="16"
      fill="currentColor" class="bi bi-x-circle-fill" viewBox="0 0 16 16">
      <path
          d="M16 8A8 8 0 1 1 0 8a8 8 0 0 1 16 0zM5.354 4.646a.5.5 0 1 0-.708.708L7.293 8l-2.647 2.646a.5.5 0 0 0 .708.708L8 8.707l2.646 2.647a.5.5 0 0 0 .708-.708L8.707 8l2.647-2.646a.5.5 0 0 0-.708-.708L8 7.293 5.354 4.646z" />
  </svg></span>
  </li>`;
      ul.appendChild(li);
    }
  })

  // Add Event Listeners to the buttons
  setCloseEventListener();
  setButtonEventListener();

  // Extract window URL
  const urlParams = new URLSearchParams(location.search);

  for (const [key, value] of urlParams) {
    if (key == "code") {
      code = value;

      // Send tenant details to taskpane
      var messageObject = { messageType: "tenant", tenant_name: localStorage.getItem('ti') };
      var jsonMessage = JSON.stringify(messageObject);
      Office.context.ui.messageParent(jsonMessage);

      // Fetch Token Generation
      const fetch = (...args) => import('node-fetch').then(({ default: fetch }) => fetch(...args));

      var myHeaders = new Headers();
      myHeaders.append("Content-Type", "application/x-www-form-urlencoded");

      var urlencoded = new URLSearchParams();
      urlencoded.append("grant_type", "authorization_code");
      urlencoded.append("client_id", localStorage.getItem('ci'));
      urlencoded.append("client_secret", localStorage.getItem('cs'));
      urlencoded.append("code", code);
      urlencoded.append("redirect_uri", "https://localhost:3000/commands.html");

      var requestOptions = {
        method: 'POST',
        headers: myHeaders,
        body: urlencoded,
        redirect: 'follow'
      };

      const address = fetch(`https://mingle-ionapi.inforcloudsuite.com/${localStorage.getItem('ti')}/CustomerApi/EXCELWrapperAPI/tokenAPI`, requestOptions)
        .then(response => response.json())
        .then((result) => {
          return result
        })
        .catch(error => console.log('error', error));

      const printAddress = async () => {
        const a = await address;
        var messageObject = { messageType: "token", access_token: a.access_token };
        var jsonMessage = JSON.stringify(messageObject);
        Office.context.ui.messageParent(jsonMessage);
      };

      // Display Modal for Sign In
      var myModal = new bootstrap.Modal(document.getElementById("myModal"));
      document.getElementById("modalHeading").innerHTML = "Sign In";
      document.getElementById("modalText").innerHTML = `Signed in Successfully.`;
      myModal.show();

      // Remove ci, cs and ti from localStorage
      localStorage.removeItem('ci');
      localStorage.removeItem('cs');
      localStorage.removeItem('ti');

      printAddress();
    }

    else {
      // Remove CI, CS and TI if exits
      if ('ci' in localStorage)
        localStorage.removeItem('ci');
      if ('cs' in localStorage)
        localStorage.removeItem('cs');
      if ('ti' in localStorage)
        localStorage.removeItem('ti');
    }
  }
});

// Login Function when ION API File is Clicked
function logIn(name) {

  Object.keys(localStorage).every(function (key) {
    if (key.includes('ionAPI')) {
      var auth_obj = JSON.parse(localStorage.getItem(key));

      if (auth_obj.cn == name) {
        // Set CI , CS and TI
        localStorage.setItem(`ci`, auth_obj.ci);
        localStorage.setItem(`cs`, auth_obj.cs);
        localStorage.setItem(`ti`, auth_obj.ti);
        window.location.replace(`${auth_obj['pu']}${auth_obj['oa']}?client_id=${auth_obj['ci']}&response_type=code&redirect_uri=${auth_obj['ru']}`);
        return false;
      }
    }
    return true;
  })
}

// Set Event Listener to Remove ION API File from dropdown
function setCloseEventListener() {
  /* Get all elements with class="close" */
  closebtns = document.getElementsByClassName("close");
  var i;

  /* Loop through the elements, and hide the parent, when clicked on */
  for (i = 0; i < closebtns.length; i++) {
    closebtns[i].addEventListener("click", function () {
      this.parentElement.style.display = 'none';
      var name = this.parentElement.id;
      var ul = document.getElementById("dynamic-list");
      var item = document.getElementById(name);
      ul.removeChild(item);
      Object.keys(localStorage).every(function (key) {
        if (key.includes('ionAPI')) {
          var opt = JSON.parse(localStorage.getItem(key));

          if (opt.cn == name) {
            localStorage.removeItem(key);
            return false;
          }
        }
        return true;
      })

    });
  }
}

// Set Event Listener when user selects the ION API file
function setButtonEventListener() {
  // Add listener to li elements
  items = document.getElementsByClassName('btn-link');
  var i;

  /* Loop through the elements, and hide the parent, when clicked on */
  for (i = 0; i < items.length; i++) {
    items[i].addEventListener("click", function () {
      logIn(this.parentElement.id);
    });
  }
}


Office.onReady((info) => {
  // If needed, Office.js is ready to be called
  if (info.host === Office.HostType.Excel) {

    // Button Event Capture
    document.getElementById("uploadBtn").addEventListener('click', openDialog);
    document.getElementById("fileid").addEventListener('change', readSingleFile, false);

  }
});

// Upload Button Click
function openDialog() {
  document.getElementById('fileid').click();
}

// Read the ION API File
function readSingleFile(e) {

  var file = e.target.files[0];
  if (file.name.split('.').pop() == "ionapi") {
    if (!file) {
      return;
    }

    var reader = new FileReader();
    reader.onload = function (e) {
      auth_obj = JSON.parse(e.target.result);
      // If ION API File is not of type WebApp 
      if (!('ru' in auth_obj)) {
        var myModal = new bootstrap.Modal(document.getElementById("myModal"));
        document.getElementById("modalHeading").innerHTML = "Upload Profile";
        document.getElementById("modalText").innerHTML = "Please Upload '.ionapi' file of type WebApp";
        myModal.show();
        return;
      }

      // Check if User Uploads same ionapi File
      flag = 0;

      // Using For Each Loop
      Object.keys(localStorage).every(function (key) {
        if (key.includes('ionAPI')) {
          var opt = JSON.parse(localStorage.getItem(key));
          if (opt.cn == auth_obj.cn) {
            localStorage.removeItem(key);
            localStorage.setItem(key, JSON.stringify(auth_obj));
            flag = 1;

            // Display Modal
            var myModal = new bootstrap.Modal(document.getElementById("myModal"));
            document.getElementById("modalHeading").innerHTML = "Upload Profile";
            document.getElementById("modalText").innerHTML = `Profile Uploaded Successfully. Please Sign in into ${auth_obj.ti} tenant.`;
            myModal.show();

            return false;
          }
        }
        return true;
      })
      // If User Uploads a New File
      if (flag == 0) {
        var date = new Date();
        var timestamp = date.getDate().toString() + (date.getMonth() + 1).toString() + date.getFullYear().toString() + date.getHours().toString() + date.getMinutes().toString() + date.getSeconds().toString();
        localStorage.setItem(`ionAPI_${timestamp}`, JSON.stringify(auth_obj));

        // Load List value
        var opt = localStorage.getItem(`ionAPI_${timestamp}`);
        var ul = document.getElementById("dynamic-list");
        var li = document.createElement("li");
        li.setAttribute('id', JSON.parse(opt).cn);
        ul.appendChild(li);

        var btn = document.createElement("button");
        btn.type = "button";
        btn.title = "Click to Sign In";
        btn.classList.add('btn');
        btn.classList.add('btn-link');
        btn.classList.add('btn-sm');
        btn.innerHTML = JSON.parse(opt).cn;
        li.appendChild(btn);
        li.innerHTML += `<span class="close"><svg xmlns="http://www.w3.org/2000/svg" width="16" height="16"
          fill="currentColor" class="bi bi-x-circle-fill" viewBox="0 0 16 16">
          <path
              d="M16 8A8 8 0 1 1 0 8a8 8 0 0 1 16 0zM5.354 4.646a.5.5 0 1 0-.708.708L7.293 8l-2.647 2.646a.5.5 0 0 0 .708.708L8 8.707l2.646 2.647a.5.5 0 0 0 .708-.708L8.707 8l2.647-2.646a.5.5 0 0 0-.708-.708L8 7.293 5.354 4.646z" />
      </svg></span>
      </li>`;
        setButtonEventListener();
        setCloseEventListener();

        var myModal = new bootstrap.Modal(document.getElementById("myModal"));
        document.getElementById("modalHeading").innerHTML = "Upload Profile";
        document.getElementById("modalText").innerHTML = `Profile Uploaded Successfully. Please Sign in into ${JSON.parse(opt).ti} tenant.`;
        myModal.show();
      }
    };
    reader.readAsText(file);

  }

  // If File is not an IONAPI File
  else {
    var myModal = new bootstrap.Modal(document.getElementById("myModal"));
    document.getElementById("modalHeading").innerHTML = "Upload Profile";
    document.getElementById("modalText").innerHTML = "Please Upload '.ionapi' file.";
    myModal.show();
    e.target.value = '';
    return;
  }

  e.target.value = '';
}

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
      ? window
      : typeof global !== "undefined"
        ? global
        : undefined;
}

const g = getGlobal();

// The add-in command functions need to be available in global scope
g.action = action;
