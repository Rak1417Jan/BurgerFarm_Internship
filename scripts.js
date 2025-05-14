// Show loading overlay when navigating between pages
function showLoading(event) {
  const loadingOverlay = document.querySelector(".loading-overlay");
  loadingOverlay.classList.add("show-loading");

  // Prevent immediate navigation to simulate loading
  event.preventDefault();

  // Navigate after a short delay to show the loading animation
  setTimeout(() => {
    window.location.href = event.currentTarget.getAttribute("href");
  }, 80);
}

// Show toast notification function
function showToast(message, duration = 3000) {
  const toast = document.getElementById("notification");
  const messageElement = document.getElementById("notification-message");

  messageElement.textContent = message;
  toast.classList.add("show");

  setTimeout(() => {
    toast.classList.remove("show");
  }, duration);
}

// Check if the page was redirected from another tool
window.addEventListener("load", () => {
  const urlParams = new URLSearchParams(window.location.search);
  const message = urlParams.get("message");

  if (message) {
    showToast(decodeURIComponent(message));

    // Clean up the URL
    const newUrl = window.location.pathname;
    window.history.replaceState({}, document.title, newUrl);
  }
});

// Add current date to the footer
document.addEventListener("DOMContentLoaded", () => {
  const footer = document.querySelector("footer");
  const currentYear = new Date().getFullYear();

  // Update the copyright year if needed
  footer.innerHTML = footer.innerHTML.replace(/\d{4}/, currentYear);
});
