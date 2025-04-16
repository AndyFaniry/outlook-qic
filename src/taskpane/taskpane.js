/* global document, Office, localStorage, window, process, URLSearchParams, console */

import { initSessionManager, setToken, deleteToken } from "../common/sessionManager";

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Outlook) {
    // document.getElementById("sideload-msg").style.display = "none";
    // document.getElementById("app-body").style.display = "flex";
    const token = await localStorage.getItem("token");
    const tokenDate = parseInt(localStorage.getItem("tokenDate"), 10);
    const now = Date.now();
    const tokenExpire = 86400 * 1000;
    if (token && now - tokenDate < tokenExpire) {
      window.location.href = "acceuil.html";
    } else {
      deleteToken();
    }
    document.getElementById("loginClick").onclick = async function () {
      await openLogin();
    };
    const params = new URLSearchParams(window.location.search);
    if (params.get("action") === "connexion") {
      await openLogin();
    }
    // Office.context.mailbox.getSelectedItemsAsync((asyncResult) => {
    //   asyncResult.value.forEach((message) => {
    //     //Read ou Compose
    //     console.log(`Item mode: ${message.itemMode}`);
    //   });
    // });
  }
});

export async function run() {
  /**
   * Insert your Outlook code here
   */

  const item = Office.context.mailbox.item;
  const userProfile = Office.context.mailbox.userProfile;
  const msgSender = Office.context.mailbox.item.sender;
  let insertAt = document.getElementById("item-subject");
  let label = document.createElement("b").appendChild(document.createTextNode("Subject: "));
  insertAt.appendChild(label);
  insertAt.appendChild(document.createElement("br"));
  insertAt.appendChild(document.createTextNode(item.subject));
  insertAt.appendChild(document.createElement("br"));
  insertAt.appendChild(document.createTextNode(userProfile.displayName));
  insertAt.appendChild(document.createElement("br"));
  insertAt.appendChild(document.createTextNode(userProfile.emailAddress));
  insertAt.appendChild(document.createElement("br"));
  insertAt.appendChild(document.createTextNode(msgSender.displayName));
  insertAt.appendChild(document.createElement("br"));
  insertAt.appendChild(document.createTextNode(msgSender.emailAddress));
}

async function openLogin() {
  await Office.context.ui.displayDialogAsync(
    `${process.env.API_URL}/login`,
    { height: 90, width: 100 },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log(asyncResult.error.code + ": " + asyncResult.error.message);
      } else {
        var dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          dialog.close();
          processMessage(arg);
        });
      }
    }
  );
}

function processMessage(arg) {
  const data = JSON.parse(arg.message);
  initSessionManager();
  setToken(data.token, data.mail_gestionnaire);
  window.location.href = "acceuil.html";
}

// function setupToken(token) {
//   localStorage.removeItem("token");
//   localStorage.setItem("token", token);
// }
