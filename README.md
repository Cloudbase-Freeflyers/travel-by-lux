# Travel by Luxe — Italy 2026 Landing Page

A high-end, bespoke landing page for Travel by Luxe. Digital concierge aesthetic with deep navy, matte black, and metallic gold accents.

## Quick Start

Open `index.html` in a browser, or run a local server:

```bash
npx serve -l 3333
```

Then visit `http://localhost:3333`.

## UTM Parameters

Pre-populate the Itinerary Architect based on ad campaigns:

- `?utm_content=scholar` — Culture 100%, Exclusivity high
- `?utm_content=epicurean` — Gastronomy 100%, Light pace
- `?utm_content=executive` — Chauffeur-only, Presidential concierge
- `?utm_content=adventurer` — Active pace, moderate culture

## Structure

```
TravelByLux/
├── index.html          # Main layout
├── css/
│   ├── variables.css   # Luxe palette, typography, spacing
│   ├── base.css        # Reset, body, typography
│   ├── layout.css      # Hero, trust, architect grid
│   ├── components.css  # Sliders, cards, buttons, glassmorphism
│   └── animations.css  # Fade-up, Ken Burns
└── js/
    ├── slider-logic.js # Itinerary Architect, UTM, dynamic preview
    └── animations.js   # Reveal on scroll
```

## Design Tokens

- **Primary:** Deep Navy `#0a1128`, Matte Black `#050810`
- **Accent:** Metallic Gold `#c5a059`
- **Typography:** Cinzel (serif), Montserrat (sans)
- **Shadows:** `0 20px 50px rgba(0,0,0,0.8)` for premium elevation
