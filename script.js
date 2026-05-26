const navToggle = document.querySelector(".nav-toggle");
const mainNav = document.querySelector(".main-nav");

if (navToggle && mainNav) {
  navToggle.addEventListener("click", () => {
    const expanded = navToggle.getAttribute("aria-expanded") === "true";
    navToggle.setAttribute("aria-expanded", String(!expanded));
    mainNav.classList.toggle("is-open");
  });
}

document.querySelectorAll("form").forEach((form) => {
  form.addEventListener("submit", (event) => {
    event.preventDefault();
    const button = form.querySelector("button[type='submit']");
    if (!button) return;
    const original = button.textContent;
    button.textContent = "Submitted";
    button.disabled = true;
    setTimeout(() => {
      button.textContent = original;
      button.disabled = false;
      form.reset();
    }, 1800);
  });
});
