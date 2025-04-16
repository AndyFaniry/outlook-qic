/* global document, Office , process, console, window, setTimeout, URLSearchParams, localStorage,$*/
import { initSessionManager } from "../common/sessionManager";
import { showLoading, hideLoading } from "../common/loadingManager";

const token = localStorage.getItem("token");

try {
  showLoading();
  Office.onReady(async (info) => {
    if (info.host === Office.HostType.Outlook) {
      initSessionManager();
      document.getElementById("return-acceuil").onclick = async function () {
        window.location.href = "acceuil.html";
      };
      document.querySelectorAll(".list-item").forEach(async (item) => {
        await item.addEventListener("click", () => {
          item.classList.toggle("expanded");
        });
      });
      const { clientId, logo, clientObject } = await getData();
      document.getElementById("open-crm").onclick = async function () {
        await Office.context.ui.openBrowserWindow(`${process.env.API_URL}/client/mon/${clientId}`);
      };
      await initData(logo, clientObject);
      //check if input as changed and display button
      document.body.addEventListener("input", async (event) => {
        if (event.target.tagName === "INPUT") {
          document.getElementById("button-changed").style.display = "block";
        }
      });
      //click undo button
      document.getElementById("undo").onclick = async function () {
        await displayDetails(clientObject);
        document.getElementById("button-changed").style.display = "none";
      };
      //click save button
      document.getElementById("save").onclick = async function () {
        await updateClient(clientId);
      };
      await new Promise((resolve) => setTimeout(resolve, 2000));
    }
  });
} catch (error) {
  console.error("Error during action:", error);
} finally {
  hideLoading();
}

async function getData() {
  const params = await new URLSearchParams(window.location.search);
  const clientId = await params.get("clientId");
  const logo = await params.get("logo");
  const client = await params.get("client");
  const clientObject = client ? await JSON.parse(decodeURIComponent(client)) : null;
  console.log(clientObject);
  return { clientId, logo, clientObject };
}

async function initData(logo, clientObject) {
  document.getElementById("company-logo").src = logo;
  await displayDetails(clientObject);
}

async function displayDetails(clientObject) {
  const nomClient = clientObject?.nom_client ? clientObject?.nom_client : clientObject?.nomClient;
  const prenomClient = clientObject?.prenom_client ? clientObject?.prenom_client : clientObject?.prenomClient;
  const name = nomClient + " " + prenomClient || "";
  const clientName = document.getElementById("client-name");
  clientName.textContent = name.toUpperCase();
  const headerName = document.getElementById("header-name");
  headerName.textContent = name.toUpperCase();
  const clientMail = document.getElementById("client-mail");
  const clientEmail = clientObject?.mail_client ? clientObject?.mail_client : clientObject?.mailClient;
  clientMail.textContent = clientEmail;
  clientMail.href = `mailto:${clientObject.mail_client}`;
  const clientAdresse = clientObject?.adresse1_client ? clientObject?.adresse1_client : clientObject?.adresse1Client;
  const clientNaissance = clientObject?.date_naissance_client
    ? clientObject?.date_naissance_client
    : clientObject?.dateNaissanceClient;
  const clientTelephone = clientObject?.telephone_client
    ? clientObject?.telephone_client
    : clientObject?.telephoneClient;
  const clientNum = clientObject?.num_adresse_client
    ? clientObject?.num_adresse_client
    : clientObject?.numAdresseClient;
  const clientBoite = clientObject?.num_boite_adresse_client
    ? clientObject?.num_boite_adresse_client
    : clientObject?.numBoiteAdresseClient;
  const clientCp = clientObject?.cp_client ? clientObject?.cp_client : clientObject?.cPClient;
  const clientVille = clientObject?.ville_client ? clientObject?.ville_client : clientObject?.villeClient;
  document.getElementById("content-nom").value = nomClient;
  document.getElementById("content-prenom").value = prenomClient;
  document.getElementById("content-adresse").value = clientAdresse;
  document.getElementById("content-dateNaissance").value = clientNaissance;
  document.getElementById("content-telephone").value = clientTelephone;
  document.getElementById("content-numero").value = clientNum;
  document.getElementById("content-boite").value = clientBoite;
  document.getElementById("content-cp").value = clientCp;
  document.getElementById("content-ville").value = clientVille;
}

function checkInputValue() {
  const boite = document.getElementById("content-boite").value;
  const cp = document.getElementById("content-cp").value;
  var errorMsg = document.getElementById("errorMsg");
  if (boite === "") {
    errorMsg.textContent = "La boîte ne doit pas être vide";
    return false;
  }
  if (boite.length > 4) {
    errorMsg.textContent = "La boîte ne doit pas être plus long que 4 caractères";
    return false;
  }
  if (cp.length < 4) {
    errorMsg.textContent = "Le code postal doit être plus long que 4 caractères";
    return false;
  }
  if (cp.length > 5) {
    errorMsg.textContent = "Le code postal ne doit pas être plus long que 5 caractères";
    return false;
  }
  return true;
}

async function updateClient(clientId) {
  const check = checkInputValue();
  if (check) {
    document.getElementById("errorMsg").value = "";
    const data = {
      id: clientId,
      nom_client: document.getElementById("content-nom").value,
      prenom_client: document.getElementById("content-prenom").value,
      adresse1_client: document.getElementById("content-adresse").value,
      telephone_client: document.getElementById("content-telephone").value,
      date_naissance_client: document.getElementById("content-dateNaissance").value || null,
      num_adresse_client: +document.getElementById("content-numero").value,
      num_boite_adresse_client: document.getElementById("content-boite").value,
      cp_client: +document.getElementById("content-cp").value,
      ville_client: document.getElementById("content-ville").value,
    };
    await $.ajax({
      url: `${process.env.API_URL}/api/clients/${clientId}`,
      method: "PATCH",
      dataType: "json",
      headers: {
        Authorization: `Bearer ${token}`,
      },
      data: JSON.stringify(data),
      beforeSend: function () {
        showLoading();
      },
      complete: function () {
        hideLoading();
      },
      success: async function () {
        document.getElementById("button-changed").style.display = "none";
      },
      error: function (error) {
        console.error("Error fetching data:", error);
      },
    });
  }
}
