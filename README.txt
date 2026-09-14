GENERIC Vanhier SSO ROUTER

Bestanden:
- mapping.json: configuratie van nieuwe SSO-routes
- sso/index.html: generieke MSAL pagina
- sso/sso-router.js: router + Entra authenticatie

Belangrijk:
De huidige Nmbrs/Yuki/ChatGPT-flow blijft ongemoeid.

Voor een nieuwe route voeg je een object toe aan mapping.json.

Voorbeeld redirect:
{
  "path": "/start-app1",
  "applicationId": "CLIENT-ID",
  "outputUrl": "https://app.example.nl",
  "type": "redirect"
}

Voor formulierflows:
{
  "path": "/start-twinfield",
  "applicationId": "CLIENT-ID",
  "outputUrl": "https://twinfield.nl",
  "type": "form",
  "form": {
    "usernameSelector": "#username",
    "passwordSelector": "#password",
    "submitSelector": "#login"
  }
}

LET OP:
De router gebruikt een MSAL redirectUri naar /sso/.
Daarom moet de App Registration deze redirect URI bevatten:
https://vanhier-accountants-adviseurs.github.io/vanhier-chrome-sso-extensie/sso/

De /start-twinfield URL moet op GitHub Pages naar de router kunnen leiden.
Omdat GitHub Pages statisch is, is hiervoor nog een kleine route/404-koppeling nodig in de bestaande repo.
