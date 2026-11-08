import {PublicClientApplication, InteractionRequiredAuthError} from '@azure/msal-browser';

// Public OAuth client identifiers only. No application secret is used in this SPA.
const msal = new PublicClientApplication({
  auth: {
    clientId: '3f887f86-9c36-4dd4-83e7-857c2b231789',
    authority: 'https://login.microsoftonline.com/dcc6cd1a-0b51-4889-9c2f-e4d97c65f29d',
    redirectUri: new URL('./redirect.html', window.location.href).href,
    ...(window.self !== window.top ? {
      popupRelayUri: new URL('./popup-relay.html', window.location.href).href,
    } : {}),
    postLogoutRedirectUri: new URL('./redirect.html', window.location.href).href,
  },
  cache: {cacheLocation: 'sessionStorage'},
});

const SCOPES = ['User.Read', 'Calendars.Read'];
const elements = Object.fromEntries(
  ['connect', 'refresh', 'signout', 'account', 'status', 'events']
    .map(id => [id, document.getElementById(id)]),
);

let busy = false;
let signedInAccount = null;

function report(message, error = false) {
  elements.status.textContent = message;
  elements.status.setAttribute('role', error ? 'alert' : 'status');
}

function renderControls() {
  elements.connect.disabled = busy || Boolean(signedInAccount);
  elements.refresh.disabled = busy || !signedInAccount;
  elements.signout.disabled = busy || !signedInAccount;
  elements.account.textContent = signedInAccount
    ? `Verbonden: ${signedInAccount.username ?? signedInAccount.name ?? 'Microsoft 365'}`
    : 'Niet aangemeld';
}

async function withBusy(task) {
  if (busy) return;
  busy = true;
  renderControls();
  try {
    await task();
  } catch (error) {
    // Never display an OAuth response, authorization code, access token or raw error payload.
    report(error instanceof InteractionRequiredAuthError
      ? 'Je Microsoft-sessie is verlopen. Klik opnieuw op Afspraken ophalen om aan te melden.'
      : 'De actie is niet gelukt. Controleer de browserinstellingen en Microsoft-toestemming.', true);
  } finally {
    busy = false;
    renderControls();
  }
}

function formatEventDate(dateTime, timeZone) {
  if (!dateTime) return 'Tijd onbekend';
  // Microsoft Graph returns timestamps without a Z when the requested Outlook time zone is used.
  const iso = /(?:Z|[+-]\d\d:\d\d)$/i.test(dateTime) ? dateTime : `${dateTime}Z`;
  const date = new Date(iso);
  if (Number.isNaN(date.getTime())) return 'Tijd onbekend';
  return new Intl.DateTimeFormat('nl-NL', {
    timeZone: timeZone === 'UTC' ? 'UTC' : 'Europe/Amsterdam',
    weekday: 'short', day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit',
  }).format(date);
}

function renderEvents(events) {
  elements.events.replaceChildren();
  for (const event of events) {
    const item = document.createElement('li');
    const when = document.createElement('span');
    when.className = 'time';
    if (event.isAllDay) {
      const startDate = event.start?.dateTime?.slice(0, 10);
      when.textContent = `${startDate ?? 'Datum onbekend'} · hele dag`;
    } else {
      when.textContent = `${formatEventDate(event.start?.dateTime, 'UTC')} tot ${formatEventDate(event.end?.dateTime, 'UTC')}`;
    }
    const name = document.createElement('span');
    name.className = 'subject';
    name.textContent = event.isCancelled ? `[Geannuleerd] ${event.subject ?? 'Zonder titel'}` : event.subject ?? 'Zonder titel';
    item.append(when, name);
    elements.events.append(item);
  }
}

async function graphToken() {
  const request = {account: signedInAccount, scopes: SCOPES};
  try {
    return (await msal.acquireTokenSilent(request)).accessToken;
  } catch (error) {
    if (!(error instanceof InteractionRequiredAuthError)) throw error;
    return (await msal.acquireTokenPopup(request)).accessToken;
  }
}

async function refresh() {
  if (!signedInAccount) return;
  report('Outlook-afspraken worden opgehaald…');
  const token = await graphToken();
  const start = new Date();
  const end = new Date(start.getTime() + 14 * 24 * 60 * 60 * 1000);
  const params = new URLSearchParams({
    startDateTime: start.toISOString(),
    endDateTime: end.toISOString(),
    '$select': 'subject,start,end,isAllDay,isCancelled',
    '$top': '50',
    '$orderby': 'start/dateTime',
  });
  let next = `https://graph.microsoft.com/v1.0/me/calendarView?${params}`;
  const events = [];
  do {
    const url = new URL(next);
    if (url.origin !== 'https://graph.microsoft.com' || !url.pathname.startsWith('/v1.0/me/calendarView')) {
      throw new Error('Onverwachte Microsoft Graph-paginering');
    }
    const response = await fetch(url, {
      headers: {Authorization: `Bearer ${token}`, Accept: 'application/json'},
      cache: 'no-store',
    });
    if (!response.ok) throw new Error(`Microsoft Graph HTTP ${response.status}`);
    const data = await response.json();
    if (!Array.isArray(data.value)) throw new Error('Ongeldig agenda-antwoord');
    events.push(...data.value);
    next = typeof data['@odata.nextLink'] === 'string' ? data['@odata.nextLink'] : null;
  } while (next && events.length < 250);
  events.sort((a, b) => (a.start?.dateTime ?? '').localeCompare(b.start?.dateTime ?? ''));
  renderEvents(events);
  report(events.length ? `${events.length} afspraak/afspraken opgehaald. Alleen-lezen DEV-test.` : 'Geen afspraken in de komende 14 dagen.');
}

elements.connect.addEventListener('click', () => withBusy(async () => {
  report('Microsoft-aanmelding wordt geopend…');
  const result = await msal.loginPopup({scopes: SCOPES, prompt: 'select_account'});
  signedInAccount = result.account;
  msal.setActiveAccount(signedInAccount);
  renderControls();
  await refresh();
}));

elements.refresh.addEventListener('click', () => withBusy(refresh));

elements.signout.addEventListener('click', () => withBusy(async () => {
  // Local sign-out only: do not sign the user out of other Microsoft applications.
  if (signedInAccount) await msal.clearCache({account: signedInAccount});
  signedInAccount = null;
  elements.events.replaceChildren();
  report('De lokale aanmeldsessie is gewist. Outlook blijft aangemeld.');
}));

withBusy(async () => {
  await msal.initialize();
  signedInAccount = msal.getActiveAccount() ?? msal.getAllAccounts()[0] ?? null;
  report(signedInAccount ? 'Je sessie is gevonden. Klik op Afspraken ophalen.' : 'Klik op Inloggen met Microsoft. Er worden geen agenda-afspraken gewijzigd.');
});
