/**
 * Reveal on Scroll — Travel by Luxe
 */

const REVEAL_OPTIONS = {
  threshold: 0.15,
  rootMargin: '0px 0px -60px 0px',
};

function revealOnScroll() {
  const cards = document.querySelectorAll('.itinerary-feature-card');
  const observer = new IntersectionObserver((entries) => {
    entries.forEach((entry) => {
      if (entry.isIntersecting) {
        entry.target.classList.add('revealed');
      }
    });
  }, REVEAL_OPTIONS);

  cards.forEach((card) => observer.observe(card));
}

document.addEventListener('DOMContentLoaded', revealOnScroll);
