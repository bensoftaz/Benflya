# Benson & Precious — QR RSVP + Wedding Day Check-In

## Files
- `Benson-Precious-Wedding-Fixed.html` — main wedding website with structured RSVP form and WhatsApp contact/reminder.
- `apps-script.gs` — Google Sheets + Apps Script backend.
- `checkin.html` — protected usher/admin scanner.

## Backend setup
1. Create a Google Sheet.
2. Extensions → Apps Script.
3. Paste `apps-script.gs` into Code.gs.
4. Change `ADMIN_PASSWORD` from the placeholder.
5. Deploy → New deployment → Web app.
6. Execute as: Me.
7. Who has access: Anyone.
8. Copy the `/exec` URL.

## Connect the website
Open the main HTML and set `RSVP_API_URL` to the `/exec` URL.
The structured RSVP form sends JSON and receives a unique `BP-XXXXXX` guest ID.
For attending guests with an email address, Apps Script emails a QR code.

## Connect the scanner
Open `checkin.html` and set `API_URL` to the same `/exec` URL.
The usher signs in, starts the camera, and scans a guest QR code.
The backend prevents duplicate check-ins and returns the guest name, party size and dietary notes.

## Security
The scanner requires a username/password and receives a short-lived token. Change the default admin password before deployment. Do not expose the `?action=list` endpoint publicly; it requires the admin token.

## Important
The QR image is generated through `api.qrserver.com`. Email delivery uses the Google account's MailApp quota. The website and scanner need internet access for the backend/check-in operation.
