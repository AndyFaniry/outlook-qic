/* global document, Office , console, window, setTimeout*/

import { initSessionManager, getQueryParams, deleteToken } from "../common/sessionManager";
import { showLoading, hideLoading } from "../common/loadingManager";

try {
  showLoading();
  Office.onReady(async (info) => {
    if (info.host === Office.HostType.Outlook) {
      initSessionManager();
      const receivedData = await getQueryParams();
      initData(receivedData);
      document.getElementById("return-acceuil").onclick = async function () {
        window.location.href = `acceuil.html`;
      };
      document.getElementById("deconnexion").onclick = async function () {
        await deconnexion();
      };
      document.getElementById("changer-compte").onclick = async function () {
        await changerCompte();
      };
      await new Promise((resolve) => setTimeout(resolve, 2000));
    }
  });
} catch (error) {
  console.error("Error during action:", error);
} finally {
  hideLoading();
}

function initData(data) {
  console.log(data);
  //boite mail connecter
  document.getElementById("mail-card").textContent = data.boite_mail_connetcted;
  //compte QIC
  document.getElementById("mail-container").textContent = data.mail_gestionnaire;
  document.getElementById("compte-id").textContent = data.gestionnaire_id;
  document.getElementById("boite-mail").textContent = data.boite_mail_gestionnaire;
}

async function deconnexion() {
  await deleteToken();
  window.location.href = `taskpane.html`;
}

async function changerCompte() {
  await deleteToken();
  window.location.href = "taskpane.html?action=connexion";
}
