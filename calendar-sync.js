// calendar-sync-app/index.js

require('dotenv').config();
const express = require('express');
const { google } = require('googleapis');
const { Client } = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');

const app = express();
app.use(express.json());

const SYNC_TAG = '[SyncedByMyApp]';

// ========== CALENDAR SYNC LIST ========== //

const SYNC_CALENDARS = [
	{ name: 'Work Google', provider: 'google', calendarId: 'work@example.com' },
	{ name: 'Personal Google', provider: 'google', calendarId: 'primary' },
	{ name: 'Personal Outlook', provider: 'outlook', calendarId: 'me' },
	// Add more calendars here
];

// ========== GOOGLE SETUP ========== //
const oauth2Client = new google.auth.OAuth2(
	process.env.GOOGLE_CLIENT_ID,
	process.env.GOOGLE_CLIENT_SECRET,
	process.env.GOOGLE_REDIRECT_URI
);

oauth2Client.setCredentials({
	access_token: process.env.GOOGLE_ACCESS_TOKEN,
	refresh_token: process.env.GOOGLE_REFRESH_TOKEN
});

const calendar = google.calendar({ version: 'v3', auth: oauth2Client });

// ========== MICROSOFT SETUP ========== //
const microsoftClient = Client.init({
	authProvider: (done) => {
		done(null, process.env.MS_ACCESS_TOKEN);
	},
});

// ========== HELPERS ========== //
function buildOutlookEvent(googleEvent) {
	const description = googleEvent.description || '';
	
	return {
		subject: googleEvent.summary,
		start: {
			dateTime: googleEvent.start.dateTime || googleEvent.start.date,
			timeZone: 'UTC',
		},
		end: {
			dateTime: googleEvent.end.dateTime || googleEvent.end.date,
			timeZone: 'UTC',
		},
		body: {
			contentType: 'Text',
			content: `${description}\n\n${SYNC_TAG}`,
		},
	};
}

function buildGoogleEvent(outlookEvent) {
	const content = outlookEvent.body?.content || '';
	
	return {
		summary: outlookEvent.subject,
		start: {
			dateTime: outlookEvent.start.dateTime,
		},
		end: {
			dateTime: outlookEvent.end.dateTime,
		},
		description: `${content}\n\n${SYNC_TAG}`,
	};
}

function buildGoogleToGoogleEvent(sourceEvent) {
	const description = sourceEvent.description || '';
	return {
		summary: sourceEvent.summary,
		start: sourceEvent.start,
		end: sourceEvent.end,
		description: `${description}\n\n${SYNC_TAG}`,
	};
}

function buildOutlookToOutlookEvent(sourceEvent) {
	const content = sourceEvent.body?.content || '';
	return {
		subject: sourceEvent.subject,
		start: sourceEvent.start,
		end: sourceEvent.end,
		body: {
			contentType: 'Text',
			content: `${content}\n\n${SYNC_TAG}`,
		},
	};
}

// ========== SYNC FROM GOOGLE TO OUTLOOK ========== //
async function handleGoogleWebhook(req, res) {
	const calendarId = 'primary';
	try {
		const eventsRes = await calendar.events.list({
			calendarId,
			timeMin: new Date().toISOString(),
			maxResults: 1,
			singleEvents: true,
			orderBy: 'updated',
		});
		
		const [event] = eventsRes.data.items;
		if (event && !(event.description || '').includes(SYNC_TAG)) {
			// NOTE: Re-building the event again for each calendar isn't optimal, but this is easier to read.
			for (const calendar of SYNC_CALENDARS) {
				if (calendar.provider == 'google') {
					const newEvent = buildGoogleToGoogleEvent(event);
					await calendar.events.insert({ calendarId: calendar.calendarId, requestBody: newEvent });
				} else if (calendar.provider == 'outlook') {
					const outlookEvent = buildOutlookEvent(event);
					await microsoftClient.api(`/users/${outlookCalendarId}/events`).post(outlookEvent);
				} else {
					console.error('Unrecognized calendar provider type:', calendar);
					res.status(500).send('Error');
				}
			}
		}
		res.status(200).send('OK');
	} catch (err) {
		console.error('Google Webhook Error:', err);
		res.status(500).send('Error');
	}
}

// ========== SYNC FROM OUTLOOK TO GOOGLE ========== //
async function handleOutlookWebhook(req, res) {
	const eventData = req.body?.value?.[0];
	if (!eventData) return res.status(400).send('No event data');
	
	try {
		const outlookEvent = await microsoftClient.api(`/me/events/${eventData.id}`).get();
		const content = outlookEvent.body?.content || '';
		if (!content.includes(SYNC_TAG)) {
			// NOTE: Re-building the event again for each calendar isn't optimal, but this is easier to read.
			for (const calendar of SYNC_CALENDARS) {
				if (calendar.provider == 'google') {
					const googleEvent = buildGoogleEvent(outlookEvent);
					await calendar.events.insert({ calendarId: calendar.calendarId, requestBody: googleEvent });
				} else if (calendar.provider == 'outlook') {
					const newEvent = buildOutlookToOutlookEvent(outlookEvent);
					await microsoftClient.api(`/users/${outlookCalendarId}/events`).post(newEvent);
				} else {
					console.error('Unrecognized calendar provider type:', calendar);
					res.status(500).send('Error');
				}
			}
		}
		res.status(200).send('OK');
	} catch (err) {
		console.error('Outlook Webhook Error:', err);
		res.status(500).send('Error');
	}
}

// ========== ROUTES ========== //
app.post('/webhook/google', handleGoogleWebhook);
app.post('/webhook/outlook', handleOutlookWebhook);

// Verification for MS Graph webhook subscription
app.get('/webhook/outlook', (req, res) => {
	if (req.query && req.query.validationToken) {
		res.send(req.query.validationToken);
	} else {
		res.status(400).send('Missing validationToken');
	}
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
