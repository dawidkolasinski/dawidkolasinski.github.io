# Dawid Kolasiński — portfolio and app pages

Public GitHub Pages site for Dawid Kolasiński, Sunchucks and Moonchucks.
The account-level hostname keeps `app-ads.txt` at the domain root, as required by
AdMob's app verification when this URL is used as the App Store Marketing URL.

English is the default language. Privacy, Terms and Support have separate Polish pages
and language links. Language selection uses explicit links, with no automatic
browser-language redirect. Both translations must be checked against the actual
app and SDK behavior before each release and updated together.

App support routes:

| Page | English | Polish |
| --- | --- | --- |
| App support directory | `/support/` | `/pl/support/` |
| Sunchucks support | `/sunchucks/support/` | `/sunchucks/pl/support/` |
| Moonchucks support | `/moonchucks/support/` | `/moonchucks/pl/support/` |

Sunchucks legal pages are under `/privacy/`, `/terms/`, `/pl/privacy/` and
`/pl/terms/`. Moonchucks has its own legal pages under `/moonchucks/`.

Support and privacy contact: sunchucks.support@gmail.com.

The portfolio is at `/`, with product pages at `/moonchucks/` and `/sunchucks/`.
Update the portfolio and the corresponding product page together when an app's
availability changes. Add an App Store link only after verifying the public listing.

The portfolio uses app icons exported from the app repositories and landscape
marketing artwork from the previous site draft. The landscapes are illustrative.

Preview from this directory with `python3 -m http.server 8080 --bind 127.0.0.1`.
