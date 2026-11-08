# WJ Projects Agenda, Microsoft Graph

Afzonderlijke, openbare **alleen-lezen** technische proef van een Outlook-agendakoppeling via Microsoft Graph. De bestaande WJ Projects Airtable Agenda, productiebase, automations en persoonlijke Outlook-agenda's worden door deze applicatie niet gewijzigd.

## Microsoft Entra

- Tenant: WJ Projects, alleen organisatieaccounts.
- Beoogde appnaam: `WJ Projects Agenda` (pas deze ook afzonderlijk aan in Microsoft Entra).
- Appregistratie: **Single-page application (SPA)**.
- Redirect URI: `https://ahashwj.github.io/wj-projects-agenda-auth/redirect.html`
- Delegated permissions: `User.Read`, `Calendars.Read`.
- **Geen client secret of certificaat**. De OAuth-flow gebruikt Authorization Code met PKCE in de browser.
- Aanmeldingen verlopen via een popup en bij embedden via een same-origin popup relay.

## Gebruik

1. Publiceer deze repository via GitHub Pages en controleer dat de redirect URI bereikbaar is.
2. Voeg de exacte SPA redirect URI toe aan de Entra-app onder **Authentication**.
3. Open `https://ahashwj.github.io/wj-projects-agenda-auth/` op een laptop.
4. Klik **Inloggen met Microsoft** en meld je aan met een WJ Projects-account.
5. Controleer alleen-lezen afspraken van de komende 14 dagen.

De data wordt uitsluitend direct vanuit de browser met Microsoft Graph gelezen. De interface bewaart geen afspraken in de repository of op GitHub. Tokens worden door MSAL beheerd in browser-sessionStorage, niet op de server. Bestaande Outlook-apparaten blijven werken.

## Beperkingen

- Dit is een alleen-lezen proef buiten de bestaande Airtable Custom Element, geen volledige Outlook-vervanging en geen bidirectionele synchronisatie.
- Voor een productieomgeving moeten onder andere gebruikersautorisatie, privacy, beheer van agenda's, afspraakwijzigingen, Teams-integratie, wijzigingen vanuit Outlook, dubbele records, beveiligingsheaders en foutafhandeling worden ontworpen en getest.
- Een embed in Airtable vereist een afzonderlijke browser- en Airtable SDK-test. Niet als werkend aannemen voordat dat is aangetoond.
- GitHub Pages geeft geen volledige controle over HTTP-responsheaders. Voor een productie-login is geschikte hosting met beheersbare security headers waarschijnlijk nodig.

## Ontwikkelen

```sh
npm ci
npm run build
```

Er mogen geen wachtwoorden, secrets, toegangstokens of privégegevens aan de repository worden toegevoegd.
