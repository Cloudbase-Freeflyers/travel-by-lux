# Travel by Luxe — Brand manifest & kit

This document captures **brand identity** and **design tokens** used on the live site, with colors and typography taken from the public Elementor global stylesheet ([Elementor kit 153](https://travelbyluxe.com/wp-content/uploads/elementor/css/post-153.css)), which powers pages such as [Italy itineraries](https://travelbyluxe.com/italy/itineraries/).

---

## Brand overview

| Field | Detail |
|--------|--------|
| **Primary name** | Travel by Luxe |
| **Regional product** | Italy by Luxe (Italy hub) |
| **Website** | [https://travelbyluxe.com/italy/itineraries/](https://travelbyluxe.com/italy/itineraries/) |
| **Positioning** | Private, luxury, multi-day and day tours across Europe; bespoke itineraries; English-speaking guides |
| **Contact (US)** | +1 (424) 766-5955 |
| **Contact (Italy)** | +39 06 9453 4617 |
| **Email** | hello@travelbyluxe.com |

---

## Voice & messaging (from live site)

- **Tone:** Expert, warm, confident; emphasis on “private,” “custom,” “complimentary planning,” and “specialist tour planners.”
- **Proof points:** TripAdvisor Travelers’ Choice, 5-star Google/TripAdvisor reviews, licensed guides, verified operator.
- **Legal:** Ts & Cs, Privacy Policy, Cookie policy (links on main site footer).

---

## Color system (brand kit)

| Token | Hex | Use |
|--------|-----|-----|
| **Primary / black** | `#000000` | Pure black (Elementor global; primary label) |
| **Heading / slate** | `#21313B` | H1–H6, strong headlines, nav emphasis |
| **Body / secondary** | `#6F757E` | Default paragraph and secondary UI text |
| **Accent (teal)** | `#2FA3A3` | Primary CTAs, links, key icons, focus rings |
| **Accent light** | `#4DB8B8` | Hover state on teal buttons |
| **White** | `#FFFFFF` | Cards, text on dark/teal buttons |
| **Page background (warm)** | `#FDFBF9` | Large section backgrounds |
| **Neutral gray** | `#F3F3F4` | Alternate bands, soft panels |
| **Tint (teal wash)** | `#EAF5F6` | Soft highlights, subtle fills |
| **Secondary accent (peach)** | `#FFBC7D` | Page transition / secondary highlight (Elementor kit) |

**CSS variables (local project)** — see `luxe.html` `:root`:

- `--text` / `--text-muted` → `#6F757E`
- `--text-heading` → `#21313B`
- `--accent` → `#2FA3A3`
- `--accent-light`, `--accent-soft`, `--on-accent`
- `--bg`, `--bg-soft`, `--bg-tint`, `--bg-card`
- `--border`, `--radius-*`, `--shadow-*`

---

## Typography (brand kit)

| Role | Family | Notes (from live site kit) |
|------|--------|------------------------------|
| **Display / headings** | **Italiana** | H1–H6, hero titles; elegant serif for “luxury travel” feel |
| **Body / UI** | **Inter** | 14px base on site; forms, labels, buttons |

**Google Fonts (subset used in `luxe.html`):**

- Inter: weights 400, 500, 600, 700 (with variable opsz where loaded)
- Italiana: regular weight for headings

**Button pattern (live site):** Teal background (`#2FA3A3`), white text, **pill shape** (`border-radius: 100px`), generous horizontal padding (~15–20px).

---

## Logo & imagery

- **Logo:** “Italy by Luxe” wordmark in script style; brand materials often use a **turquoise/teal** field behind the logo (see [itineraries page](https://travelbyluxe.com/italy/itineraries/) hero and assets). Align marketing backgrounds with **`#2FA3A3`** or **`#EAF5F6`** for consistency.
- **Photos:** Rounded corners on site imagery (Elementor kit sets **20px** radius on images globally).

---

## Social & reviews

- TripAdvisor, Google, Facebook, Twitter/X, YouTube, Instagram (footer on main site).
- Review snippets and “Travelers’ Choice” badges appear on Italy itineraries and similar hubs.

---

## Implementation note

This manifest is aligned with **public CSS** from `travelbyluxe.com` as of extraction from `post-153.css` (Elementor global kit). If the live site changes its theme, re-sync tokens from:

`https://travelbyluxe.com/wp-content/uploads/elementor/css/post-153.css`

---

*Internal project file: keep in sync with `luxe.html` when adjusting campaign landing pages.*
