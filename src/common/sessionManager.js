/* global setInterval, localStorage, window, URLSearchParams */
// const inactivityLimit = 10 * 60 * 1000; // 10 minutes in milliseconds
const inactivityLimit = 10 * 60 * 1000000; // 10 minutes in milliseconds
export function setToken(token, mail) {
  const timestamp = Date.now();
  localStorage.setItem("token", token);
  localStorage.setItem("mail", mail);
  localStorage.setItem("lastActivity", timestamp.toString());
  localStorage.setItem("tokenDate", timestamp.toString());
}

function checkSession() {
  const token = localStorage.getItem("token");
  const lastActivity = parseInt(localStorage.getItem("lastActivity"), 10);
  const now = Date.now();

  if (token && lastActivity && now - lastActivity > inactivityLimit) {
    deleteToken();
    window.location.href = "taskpane.html";
  }
}

function resetActivity() {
  localStorage.setItem("lastActivity", Date.now().toString());
}

export function initSessionManager() {
  setInterval(checkSession, 60 * 1000);
  window.addEventListener("mousemove", resetActivity);
  window.addEventListener("keypress", resetActivity);
}

export async function getQueryParams() {
  const params = await new URLSearchParams(window.location.search);
  const data = {};
  for (const [key, value] of params.entries()) {
    data[key] = value;
  }
  return data;
}

export function deleteToken() {
  localStorage.removeItem("token");
  localStorage.removeItem("mail");
  localStorage.removeItem("lastActivity");
  localStorage.removeItem("tokenDate");
}
