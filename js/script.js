// Small offline-friendly interactions for the project demonstration.
const menuButton = document.querySelector('.menu-toggle');
const navigation = document.querySelector('#primary-nav');

if (menuButton && navigation) {
  menuButton.addEventListener('click', () => {
    const isOpen = menuButton.getAttribute('aria-expanded') === 'true';
    menuButton.setAttribute('aria-expanded', String(!isOpen));
    navigation.classList.toggle('is-open', !isOpen);
  });
}

const requirementsButton = document.querySelector('#requirements-toggle');
const requirementsPanel = document.querySelector('#requirements-panel');

if (requirementsButton && requirementsPanel) {
  requirementsButton.addEventListener('click', () => {
    const isOpen = requirementsButton.getAttribute('aria-expanded') === 'true';
    requirementsButton.setAttribute('aria-expanded', String(!isOpen));
    requirementsPanel.hidden = isOpen;
    requirementsButton.textContent = isOpen ? 'Show course requirements' : 'Hide course requirements';
  });
}

const inquiryForm = document.querySelector('#inquiry-form');
const feedback = document.querySelector('#form-feedback');

if (inquiryForm && feedback) {
  inquiryForm.addEventListener('submit', (event) => {
    event.preventDefault();
    feedback.textContent = 'Demo only: your inquiry was checked locally and was not sent or stored.';
  });
}
