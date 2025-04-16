/* global document, Office , Image, process, $, console, localStorage, window, URLSearchParams, setTimeout, clearTimeout*/

import { initSessionManager } from "../common/sessionManager";
import { showLoading, hideLoading } from "../common/loadingManager";

const token = localStorage.getItem("token");
const mail_gestionnaire = localStorage.getItem("mail");
let clientId = 0;
let logo = "../../assets/profil.png";
let user = "";
let client = "";
let searchClients = [];
let conversationId = "";
let emailBody = "";
let checkBoxContact = "checkBoxContact";
let checkBoxDossier = "checkBoxDossier";
let contactDiv = "email-group-contacts";
let dossierDiv = "email-group-dossiers";
let checkedContactId = [];
let checkedDossierId = [];
let dossierId = [];
let emailId = null;
try {
  showLoading();
  Office.onReady(async (info) => {
    if (info.host === Office.HostType.Outlook) {
      initSessionManager();
      let userAccount = Office.context.mailbox.userProfile;
      const msgSender = Office.context.mailbox.item.sender;
      conversationId = Office.context.mailbox.item.conversationId;
      conversationId = conversationId.replace(/\//g, "");
      // console.log(Office.context.mailbox.item);

      const name = msgSender.displayName;
      const email = msgSender.emailAddress;
      await loadUser();
      await loadData(email);
      await checkEmail();
      await getBodyMail();
      await checkDossierIdObject();
      // await searchClient();
      document.getElementById("add-crm").onclick = async function () {
        await addClient(email, name);
      };
      document.getElementById("open-crm").onclick = async function () {
        Office.context.ui.openBrowserWindow(`${process.env.API_URL}/client/mon/${clientId}`);
      };
      document.getElementById("afficher-contact").onclick = async function () {
        const dataToPass = {
          clientId: clientId,
          logo: logo,
          client: JSON.stringify(client),
        };
        let queryString = new URLSearchParams();
        queryString.append("clientId", dataToPass.clientId);
        queryString.append("logo", dataToPass.logo);
        queryString.append("client", encodeURIComponent(dataToPass.client));
        window.location.href = `profil.html?${queryString.toString()}`;
      };
      document.getElementById("open-parameter").onclick = async function () {
        const dataToPass = {
          gestionnaire_id: user.id,
          boite_mail_connetcted: userAccount.emailAddress,
          mail_gestionnaire: user.mail_gestionnaire,
          boite_mail_gestionnaire: user.mail_gestionnaire.split("@")[1],
        };
        const queryString = new URLSearchParams(dataToPass).toString();
        window.location.href = `parametre.html?${queryString}`;
      };
      document.getElementById("dropdownToggle").onclick = function () {
        toggleDropdown();
      };
      document.getElementById("closeToggle").onclick = function () {
        toggleDropdown();
      };
      document.querySelectorAll(".email-item").forEach((item) => {
        item.addEventListener("click", () => {
          const checkbox = item.querySelector('input[type="checkbox"]');
          checkbox.checked = !checkbox.checked;
        });
      });
      document.addEventListener("click", function (event) {
        const searchContainer = document.querySelector(".search-bar");
        const resultsDiv = document.getElementById("search-results");
        if (!searchContainer.contains(event.target)) {
          resultsDiv.style.display = "none";
        }
      });
      const searchInput = document.getElementById("search-input");
      let timeout = null;
      searchInput.addEventListener("input", async () => {
        if (timeout) clearTimeout(timeout);
        timeout = await setTimeout(async () => {
          await searchFunction(searchInput.value);
        }, 300);
      });
      document.getElementById("saveToggle").onclick = function () {
        modal.style.display = "block";
      };
      const modal = document.getElementById("myModal");
      document.querySelector(".close-modal").addEventListener("click", () => {
        modal.style.display = "none";
      });
      // Close modal if user clicks outside the modal content
      window.addEventListener("click", (e) => {
        if (e.target === modal) {
          modal.style.display = "none";
        }
      });
      document.getElementById("saveMail").onclick = function () {
        assignEmail();
      };
      await new Promise((resolve) => setTimeout(resolve, 2000));
    }
  });
} catch (error) {
  console.error("Error during action:", error);
} finally {
  hideLoading();
}

async function loadData(email) {
  await displayLogo(email);
  await loadClients(email);
}

export async function displayLogo(email) {
  let domain = getDomainFromEmail(email);
  if (!isPersonalEmail(domain)) {
    await checkLogoExists(domain, function (exists, logoUrl) {
      if (exists) {
        document.getElementById("company-logo").src = logoUrl;
        logo = logoUrl;
      } else {
        document.getElementById("company-logo").src = "../../assets/profil.png";
      }
    });
  } else {
    document.getElementById("company-logo").src = "../../assets/profil.png";
  }
}

function displayClientInformation(response, isClient) {
  if (isClient) {
    client = response;
    const nomClient = response?.nom_client ? response?.nom_client : response?.nomClient;
    const prenomClient = response?.prenom_client ? response?.prenom_client : response?.prenomClient;
    const name = nomClient + " " + prenomClient || "";
    const clientName = document.getElementById("client-name");
    clientName.textContent = name.toUpperCase();
    addEmailItem(response.id, response, checkBoxContact, contactDiv, true);
    if (response?.dossiers?.length > 0) {
      for (let index = 0; index < response.dossiers.length; index++) {
        const pushed = pushCheck(dossierId, response.dossiers[index].id);
        // dossierId.push(response.dossiers[index].id);
        if (pushed)
          addEmailItem(response.dossiers[index].id, response.dossiers[index], checkBoxDossier, dossierDiv, false);
      }
    }
  }
  const clientMail = document.getElementById("client-mail");
  clientMail.textContent = response.mail_client;
  clientMail.href = `mailto:${response.mail_client}`;
}

function getDomainFromEmail(email) {
  return email.substring(email.lastIndexOf("@") + 1);
}

function isPersonalEmail(domain) {
  // Liste des domaines à exclure (fournisseurs de messagerie)
  const personalEmailDomains = [
    "gmail.com",
    "yahoo.com",
    "yahoo.fr",
    "yahoo.co.uk",
    "hotmail.com",
    "hotmail.fr",
    "hotmail.co.uk",
    "outlook.com",
    "outlook.fr",
    "outlook.co.uk",
    "live.com",
    "live.fr",
    "aol.com",
    "aol.fr",
    "icloud.com",
  ];

  return personalEmailDomains.includes(domain);
}

async function checkLogoExists(domain, callback) {
  let logoUrl = `${process.env.LOGO_API_URL}/${domain}`;
  let img = new Image();
  img.onload = async function () {
    await callback(true, logoUrl);
  };
  img.onerror = async function () {
    await callback(false, logoUrl);
  };
  img.src = logoUrl;
}

async function loadClients(email) {
  await $.ajax({
    url: `${process.env.API_URL}/api/clients/${email}`,
    method: "GET",
    dataType: "json",
    headers: {
      Authorization: `Bearer ${token}`,
    },
    success: function (response) {
      if (response) {
        clientId = response.id;
        $("#add-crm").hide();
        displayClientInformation(response, true);
      } else {
        $("#afficher-contact").prop("disabled", true).css("cursor", "not-allowed").css("width", "100%");
        $("#dropdownToggle").prop("disabled", true).css("cursor", "not-allowed").css("color", "lightgray");
        $("#open-crm").hide();
        let response = {};
        response.mail_client = email;
        displayClientInformation(response, false);
      }
    },
    error: function (error) {
      console.error("Error fetching data:", error);
    },
  });
}
async function addClient(email, name) {
  try {
    showLoading();
    var data = getDataClient(email, name);
    await $.ajax({
      url: `${process.env.API_URL}/api/clients`,
      method: "POST",
      dataType: "json",
      data: JSON.stringify(data),
      contentType: "application/json",
      headers: {
        Authorization: `Bearer ${token}`,
      },
      success: function (response) {
        if (response) {
          clientId = response.id;
          $("#afficher-contact").prop("disabled", false).css("cursor", "pointer").css("width", "85%");
          $("#open-crm").show();
          $("#add-crm").hide();
          displayClientInformation(response, true);
        }
      },
      error: function (error) {
        console.error("Error fetching data:", error);
      },
    });
  } catch (error) {
    console.error("Error during action:", error);
  } finally {
    hideLoading();
  }
}

function capitalize(word) {
  return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
}

function getNameFromEmail(email) {
  const localPart = email.split("@")[0];
  const nameParts = localPart.split(/[._-]/);
  if (nameParts.length > 2) {
    const nom_client = capitalize(nameParts[0]);
    const prenom_client = capitalize(nameParts[1]);
    return { nom_client, prenom_client };
  } else {
    return { nom_client: capitalize(nameParts[0]), prenom_client: "" };
  }
}

function getDataClient(email, name) {
  var data = {};
  if (email === name) {
    data = getNameFromEmail(email);
  } else {
    const nameParts = name.split(" ");
    if (nameParts.length >= 2) {
      data = { nom_client: nameParts[0], prenom_client: nameParts[1] };
    } else {
      data = { nom_client: nameParts[0], prenom_client: "" };
    }
  }
  data.mail_client = email;
  return data;
}

async function loadUser() {
  await $.ajax({
    url: `${process.env.API_URL}/api/gestionnaires/${mail_gestionnaire}`,
    method: "GET",
    dataType: "json",
    headers: {
      Authorization: `Bearer ${token}`,
    },
    success: function (response) {
      user = response;
    },
    error: function (error) {
      console.error("Error fetching data:", error);
    },
  });
}

function toggleDropdown() {
  const dropdown = document.getElementById("emailDropdown");
  dropdown.style.display = dropdown.style.display === "block" ? "none" : "block";
}

async function searchClient(mots) {
  await $.ajax({
    url: `${process.env.API_URL}/api/clients/search/${mots}`,
    method: "GET",
    dataType: "json",
    headers: {
      Authorization: `Bearer ${token}`,
    },
    success: function (response) {
      searchClients = response;
    },
    error: function (error) {
      console.error("Error fetching data:", error);
    },
  });
}

let emailCount = 0;

function addEmailItem(id, data, name, div, isContact, isChecked) {
  // Increment the counter for unique IDs
  emailCount++;

  // Create the wrapper div
  const emailItem = document.createElement("div");
  emailItem.className = "email-item";
  emailItem.id = id;

  // Create the input checkbox with unique ID
  const checkbox = document.createElement("input");
  const checkboxId = "check" + emailCount;
  checkbox.type = "checkbox";
  checkbox.id = checkboxId;
  checkbox.value = id;
  checkbox.checked = isChecked;
  checkbox.name = name;

  // Create the label and SVG inside it
  const label = document.createElement("label");
  label.setAttribute("for", checkboxId);
  label.className = "custom-checkbox";
  label.innerHTML = `
            <svg xmlns="http://www.w3.org/2000/svg" width="25" height="25" viewBox="0 0 25 25" fill="none">
                <path d="M9.375 11.4583L12.5 14.5833L22.9167 4.16667M21.875 12.5V19.7917C21.875 20.3442 21.6555 20.8741 21.2648 21.2648C20.8741 21.6555 20.3442 21.875 19.7917 21.875H5.20833C4.6558 21.875 4.12589 21.6555 3.73519 21.2648C3.34449 20.8741 3.125 20.3442 3.125 19.7917V5.20833C3.125 4.6558 3.34449 4.12589 3.73519 3.73519C4.12589 3.34449 4.6558 3.125 5.20833 3.125H16.6667" stroke="#0091AE" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
            </svg>
        `;

  // Create the span for the email address
  const emailSpan = document.createElement("span");
  emailSpan.className = "email-address";
  const nom = data?.nom_client + " " + data?.prenom_client || "";
  if (isContact) {
    emailSpan.textContent = nom;
  } else {
    emailSpan.textContent = data?.id;
  }

  // if (isDisabled) {
  //   checkbox.disabled = true;
  //   emailItem.style.pointerEvents = "none";
  //   emailItem.style.cursor = "not-allowed";
  //   label.style.border = "2px solid rgba(202, 203, 203, 0.77)"; // Set border color to blue (default)
  // }

  // Append all elements to the email item
  emailItem.appendChild(checkbox);
  emailItem.appendChild(label);
  emailItem.appendChild(emailSpan);

  // Append the new item to the email group container
  document.getElementById(div).appendChild(emailItem);
}

async function searchFunction(mots) {
  const input = document.getElementById("search-input");
  const filter = input.value.toLowerCase();
  const resultsDiv = document.getElementById("search-results");
  const arrowsDiv = document.getElementById("arrow");

  // Vider les résultats précédents
  resultsDiv.innerHTML = "";

  if (filter === "") {
    resultsDiv.style.display = "none";
    arrowsDiv.style.display = "none";
    return;
  }

  // Filtrer les données
  await searchClient(mots);
  const filteredData = searchClients.filter((item) => item.id !== clientId && item.mail_client !== null);
  if (filteredData.length === 0) {
    resultsDiv.innerHTML = "<div>Aucun résultat</div>";
  } else {
    filteredData.forEach((item) => {
      const resultItem = document.createElement("div");
      resultItem.textContent = item?.nom_client + " " + item?.prenom_client;

      // Clique sur un résultat
      resultItem.addEventListener("click", () => {
        addEmailItem(item.id, item, checkBoxContact, contactDiv, true, true);
        //ajouter dossier du client
        if (item?.dossiers?.length > 0) {
          for (let index = 0; index < item.dossiers.length; index++) {
            let checked = false;
            if (item.dossiers[index]?.dossier_statut?.statut === "Ouvert") {
              checked = true;
            }
            const pushed = pushCheck(dossierId, item.dossiers[index].id);
            // dossierId.push(item.dossiers[index].id);
            if (pushed)
              addEmailItem(item.dossiers[index].id, item.dossiers[index], checkBoxDossier, dossierDiv, false, checked);
          }
        }
        resultsDiv.style.display = "none";
        arrowsDiv.style.display = "none";
      });

      resultsDiv.appendChild(resultItem);
    });
  }

  resultsDiv.style.display = "block";
  arrowsDiv.style.display = "block";
}

async function checkEmail() {
  await $.ajax({
    url: `${process.env.API_URL}/api/emails/${conversationId}`,
    method: "GET",
    dataType: "json",
    headers: {
      Authorization: `Bearer ${token}`,
    },
    beforeSend: function () {
      hideAssing();
    },
    complete: function () {
      showAssing();
    },
    success: async function (data) {
      const assignText = document.getElementById("assign-block");
      const nameEmail = document.getElementById("name-email-input");
      if (data) {
        emailId = data.id;
        nameEmail.value = data.conversation_name;
        assignText.textContent = "Cet e-mail est suivi";
        const clients = data.clients;
        const dossiers = data.dossiers;
        if (clients.length > 0) {
          for (let index = 0; index < clients.length; index++) {
            if (clients[index].id !== clientId) {
              addEmailItem(clients[index].id, clients[index], checkBoxContact, contactDiv, true, true);
            } else {
              $(`#${clients[index].id} input[type="checkbox"]`).prop("checked", true);
            }
            pushCheck(checkedContactId, clients[index].id);
            // checkedContactId.push(clients[index].id);
          }
        }
        if (dossiers.length > 0) {
          for (let index = 0; index < dossiers.length; index++) {
            if (!dossierId.includes(dossiers[index].id)) {
              addEmailItem(dossiers[index].id, dossiers[index], checkBoxDossier, dossierDiv, false, true);
            } else {
              $(`#${dossiers[index].id} input[type="checkbox"]`).prop("checked", true);
            }
            pushCheck(checkedDossierId, dossiers[index].id);
            // checkedDossierId.push(dossiers[index].id);
          }
        }
      } else {
        emailId = null;
        assignText.textContent = "Cet e-mail n'est pas suivi";
      }
    },
    error: function (error) {
      console.error("Error fetching data:", error);
    },
  });
}

function showAssing() {
  document.getElementById("spinner-block").style.display = "none";
  document.getElementById("assign-block").style.display = "block";
  document.getElementById("dropdownToggle").classList.remove("disabled");
}

function hideAssing() {
  document.getElementById("spinner-block").style.display = "block";
  document.getElementById("assign-block").style.display = "none";
  document.getElementById("dropdownToggle").classList.add("disabled");
}

async function assignEmail() {
  const checkedContacts = getCheckedValues(checkBoxContact);
  const checkedDossiers = getCheckedValues(checkBoxDossier);
  const mail = Office.context.mailbox.item;
  const conversationName = $("#name-email-input").val();
  const data = {
    recipient: mail.to[0]?.emailAddress,
    sender: mail.sender.emailAddress,
    subject: mail.normalizedSubject,
    body: emailBody,
    sent_at: mail.dateTimeCreated,
    conversation_id: conversationId,
    conversation_name: conversationName,
  };
  await $.ajax({
    url: `${process.env.API_URL}/api/emails`,
    method: "POST",
    dataType: "json",
    headers: {
      Authorization: `Bearer ${token}`,
    },
    data: JSON.stringify(data),
    beforeSend: function () {
      showLoadingMail();
    },
    complete: function () {
      hideLoadingMail();
    },
    success: async function (data) {
      emailId = data.id;
      if (checkedContacts) {
        const checkedSet = new Set(checkedContacts);
        const checkedContactIdSet = new Set(checkedContactId);
        const toAdd = [...checkedSet].filter((id) => !checkedContactIdSet.has(+id));
        const toRemove = [...checkedContactIdSet].filter((id) => !checkedSet.has(+id));
        await Promise.all([
          ...toAdd.map((id) => addEmailToClient(emailId, id)),
          ...toRemove.map((id) => removeEmailToClient(emailId, id)),
        ]);
      }
      if (checkedDossiers) {
        const checkedSet = new Set(checkedDossiers);
        const checkedDossierIdSet = new Set(checkedDossierId);
        const toAdd = [...checkedSet].filter((id) => !checkedDossierIdSet.has(+id));
        const toRemove = [...checkedDossierIdSet].filter((id) => !checkedSet.has(+id));
        await Promise.all([
          ...toAdd.map((id) => addEmailToDossier(emailId, id)),
          ...toRemove.map((id) => removeEmailToDossier(emailId, id)),
        ]);
      }
    },
    error: function (error) {
      console.error("Error fetching data:", error);
    },
  });
}

async function getBodyMail() {
  await Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, async (bodyResult) => {
    if (bodyResult.status === Office.AsyncResultStatus.Failed) {
      console.log(`Failed to get body: ${bodyResult.error.message}`);
      return;
    }
    emailBody = bodyResult.value;
  });
}

function showLoadingMail() {
  document.getElementById("spinner-block-mail").style.display = "block";
  document.getElementById("dropdown-block").style.display = "none";
}

function hideLoadingMail() {
  document.getElementById("spinner-block-mail").style.display = "none";
  document.getElementById("dropdown-block").style.display = "block";
  document.getElementById("assign-block").textContent = "Cet e-mail est suivi";
  document.getElementById("myModal").style.display = "none";
  toggleDropdown();
}

async function addEmailToClient(emailId, clientId) {
  const data = {
    email_id: emailId,
  };
  await $.ajax({
    url: `${process.env.API_URL}/api/clients/email/${clientId}/assign-client`,
    method: "POST",
    dataType: "json",
    headers: {
      Authorization: `Bearer ${token}`,
    },
    data: JSON.stringify(data),
    success: async function () {
      pushCheck(checkedContactId, clientId);
      // checkedContactId.push(clientId);
    },

    error: function (error) {
      console.error("Error fetching data:", error);
    },
  });
}

async function removeEmailToClient(emailId, idClient) {
  const data = {
    email_id: emailId,
  };
  await $.ajax({
    url: `${process.env.API_URL}/api/clients/email/${idClient}/remove-client`,
    method: "POST",
    dataType: "json",
    headers: {
      Authorization: `Bearer ${token}`,
    },
    data: JSON.stringify(data),
    success: async function () {
      if (idClient !== clientId) {
        $(`#${idClient}`).remove();
      }
      checkedContactId = checkedContactId.filter((item) => item !== idClient);
    },

    error: function (error) {
      console.error("Error fetching data:", error);
    },
  });
}

async function addEmailToDossier(emailId, dossierId) {
  const data = {
    email_id: emailId,
  };
  await $.ajax({
    url: `${process.env.API_URL}/api/clients/email/${dossierId}/assign-dossier`,
    method: "POST",
    dataType: "json",
    headers: {
      Authorization: `Bearer ${token}`,
    },
    data: JSON.stringify(data),
    success: async function () {
      pushCheck(checkedDossierId, dossierId);
      // checkedDossierId.push(dossierId);
    },

    error: function (error) {
      console.error("Error fetching data:", error);
    },
  });
}

async function removeEmailToDossier(emailId, idDossier) {
  const data = {
    email_id: emailId,
  };
  await $.ajax({
    url: `${process.env.API_URL}/api/clients/email/${idDossier}/remove-dossier`,
    method: "POST",
    dataType: "json",
    headers: {
      Authorization: `Bearer ${token}`,
    },
    data: JSON.stringify(data),
    success: async function () {
      if (!dossierId.includes(idDossier)) {
        $(`#${idDossier}`).remove();
      }
      checkedDossierId = checkedDossierId.filter((item) => item !== idDossier);
    },

    error: function (error) {
      console.error("Error fetching data:", error);
    },
  });
}

function getCheckedValues(name) {
  const checkboxes = document.querySelectorAll(`input[name="${name}"]:checked`);
  const values = Array.from(checkboxes).map((cb) => +cb.value);
  return values;
}

async function checkDossierIdObject() {
  const mail = Office.context.mailbox.item;
  const subject = mail.normalizedSubject;
  const numbers = subject.match(/\d+/g)?.map(Number);
  if (numbers && numbers.length > 0) {
    for (const number of numbers) {
      await $.ajax({
        url: `${process.env.API_URL}/api/clients/dossier/${number}`,
        method: "GET",
        headers: {
          Authorization: `Bearer ${token}`,
        },
        success: async function (response) {
          if (response) {
            let checked = false;
            if (response?.dossier_statut?.statut === "Ouvert") {
              checked = true;
            }
            const pushed = pushCheck(dossierId, response.id);
            // dossierId.push(response.id);
            if (pushed) addEmailItem(response.id, response, checkBoxDossier, dossierDiv, false, checked);
          }
        },
        error: function (error) {
          console.error("Error fetching data:", error);
        },
      });
    }
  }
}

function pushCheck(array, value) {
  if (!array.includes(value)) {
    array.push(value);
    return true;
  }
  return false;
}
