/**
 * Itinerary Architect — Slider Logic & Dynamic Preview
 * Travel by Luxe
 */

const SLIDER_IDS = ['pace', 'culture', 'food', 'exclusive'];
const LABEL_IDS = ['lbl-pace', 'lbl-culture', 'lbl-food', 'lbl-exclusive'];

const CONTENT_MAP = {
  pace: [
    'Chauffeur Lead',
    'Light Walking',
    'Active Exploration',
    'Urban Trekker',
    'Cobblestone Adventurer',
  ],
  culture: [
    'Iconic Highlights',
    'Museum Immersion',
    'Deep History',
    "Scholar-Led",
    'Closed-Door Archives',
  ],
  food: [
    'Authentic Trattoria',
    'Vineyard Dining',
    'Michelin Selection',
    'Artisan Workshops',
    "Chef's Private Table",
  ],
  exclusive: [
    'Private Luxury',
    'Elite Access',
    'After-Hours Access',
    'Royal Standard',
    'Presidential Concierge',
  ],
};

const ITINERARY_PROFILES = {
  balanced: {
    title: 'The Balanced Immersion',
    desc: 'A 14-day journey designed to flow seamlessly between Rome, Florence, and Venice. Every detail curated for those who appreciate both grandeur and intimacy.',
    tags: ['Multi-Day Immersion', 'Private Driver-Guide', '5-Day Minimum'],
    features: ['Executive Class Transport', 'Hand-picked 5-Star Stays', 'Flexible Daily Rhythm'],
  },
  scholar: {
    title: "The Scholar's Deep Dive",
    desc: 'A deep-dive into the hidden history of Italy, featuring private access to Vatican archives, Florentine ateliers, and scholar-led museum experiences.',
    tags: ['Scholar-Led', 'Closed-Door Access', '14-Day Grand Tour'],
    features: [
      'Scholar-Led Museum Deep-Dives',
      'Behind-the-scenes Artisan Visits',
      'Private Vatican After-Hours',
      'Renaissance Manuscript Viewings',
    ],
  },
  epicurean: {
    title: 'The Epicurean Legend',
    desc: 'A journey for the palate. Private truffle hunting in Umbria, sunset dinners in Chianti vineyards, and chef-led experiences across Italy\'s finest tables.',
    tags: ['Gastronomy-First', 'Private Estates', '14-Day Grand Tour'],
    features: [
      'Private Estate Wine Tastings',
      'Artisan Pasta Masterclass',
      'Truffle Hunting & Farm-to-Table',
      "Chef's Private Table Dinners",
    ],
  },
  executive: {
    title: 'The Executive Retreat',
    desc: 'Door-to-door chauffeur service, presidential-level concierge, and exclusive access reserved for those who expect nothing less than perfection.',
    tags: ['Chauffeur-Only', 'Presidential Concierge', 'Zero-Friction'],
    features: [
      'Strictly Door-to-Door Chauffeur Service',
      'Private Jet Transfers Available',
      '24/7 Concierge Coordination',
      'Elite Hotel & Restaurant Access',
    ],
  },
  adventurer: {
    title: 'The Cobblestone Adventurer',
    desc: 'For those who prefer to explore on foot. Active days through medieval villages, coastal paths, and hidden piazzas—with luxury awaiting each evening.',
    tags: ['Active Pace', 'Walking-Focused', '10-Day Minimum'],
    features: [
      'Guided Walking Tours',
      'Coastal Path Explorations',
      'Luxury Evening Accommodation',
      'Flexible Daily Itineraries',
    ],
  },
  hybrid: {
    title: 'Your Custom Itinerary',
    desc: 'A bespoke blend of culture, gastronomy, and exclusivity. Our travel designers will craft a journey that reflects your unique preferences.',
    tags: ['Bespoke Design', 'Multi-Day Immersion', '5-Day Minimum'],
    features: [
      'Executive Class Transport',
      'Hand-picked 5-Star Stays',
      'Personalized Daily Experiences',
      'Complimentary Itinerary Planning',
    ],
  },
};

function getSliderValues() {
  return {
    pace: parseInt(document.getElementById('pace')?.value || 1, 10),
    culture: parseInt(document.getElementById('culture')?.value || 1, 10),
    food: parseInt(document.getElementById('food')?.value || 1, 10),
    exclusive: parseInt(document.getElementById('exclusive')?.value || 1, 10),
  };
}

function determineProfile(values) {
  const { pace, culture, food, exclusive } = values;

  if (pace === 1 && exclusive >= 4) return 'executive';
  if (pace >= 4 && (culture >= 3 || food >= 3)) return 'adventurer';
  if (culture >= 4 && food >= 4) return 'scholar';
  if (culture >= 4) return 'scholar';
  if (food >= 4) return 'epicurean';
  if (exclusive >= 4 && pace <= 2) return 'executive';

  return 'balanced';
}

function updateLabels(values) {
  SLIDER_IDS.forEach((id, i) => {
    const labelEl = document.getElementById(LABEL_IDS[i]);
    if (labelEl && CONTENT_MAP[id]) {
      const idx = values[id] - 1;
      labelEl.textContent = CONTENT_MAP[id][Math.max(0, idx)];
    }
  });
}

function updatePreviewCard(profileKey) {
  const profile = ITINERARY_PROFILES[profileKey] || ITINERARY_PROFILES.hybrid;
  const card = document.getElementById('preview-card');
  if (!card) return;

  const tagsEl = card.querySelector('#preview-tags');
  const titleEl = card.querySelector('#preview-title');
  const descEl = card.querySelector('#preview-desc');
  const featuresEl = card.querySelector('#preview-features');

  if (tagsEl) {
    tagsEl.innerHTML = profile.tags
      .map((t) => `<span class="feature-tag">${t}</span>`)
      .join('');
  }
  if (titleEl) titleEl.textContent = profile.title;
  if (descEl) descEl.textContent = profile.desc;
  if (featuresEl) {
    featuresEl.innerHTML = profile.features
      .map((f) => `<li>${f}</li>`)
      .join('');
  }
}

function showLeadForm() {
  const form = document.getElementById('form-container');
  if (form) {
    form.classList.remove('lead-form--hidden');
  }
}

function updateUI() {
  const values = getSliderValues();
  updateLabels(values);
  const profile = determineProfile(values);
  updatePreviewCard(profile);
  showLeadForm();
}

function initFromUTM() {
  const params = new URLSearchParams(window.location.search);
  const utmContent = (params.get('utm_content') || '').toLowerCase();

  const paceEl = document.getElementById('pace');
  const cultureEl = document.getElementById('culture');
  const foodEl = document.getElementById('food');
  const exclusiveEl = document.getElementById('exclusive');

  if (!paceEl || !cultureEl || !foodEl || !exclusiveEl) return;

  switch (utmContent) {
    case 'scholar':
      cultureEl.value = 5;
      exclusiveEl.value = 4;
      break;
    case 'epicurean':
      foodEl.value = 5;
      paceEl.value = 2;
      break;
    case 'executive':
      paceEl.value = 1;
      exclusiveEl.value = 5;
      break;
    case 'adventurer':
      paceEl.value = 5;
      cultureEl.value = 3;
      break;
    default:
      break;
  }
}

function initSliders() {
  initFromUTM();
  updateUI();

  SLIDER_IDS.forEach((id) => {
    const el = document.getElementById(id);
    if (el) {
      el.addEventListener('input', updateUI);
      el.addEventListener('change', updateUI);
    }
  });
}

document.addEventListener('DOMContentLoaded', initSliders);
