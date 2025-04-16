/* global document */
export function showLoading() {
  document.getElementById("loading").style.display = "block";
  document.getElementById("content").style.display = "none";
}

export function hideLoading() {
  document.getElementById("loading").style.display = "none";
  document.getElementById("content").style.display = "block";
}
