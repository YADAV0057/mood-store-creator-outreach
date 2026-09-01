#!/usr/bin/env node
import 'dotenv/config';
import { google } from 'googleapis';
import { readFileSync } from 'fs';
import path from 'path';
import { generateEmail, analyzeReviewsWithAI } from './emailGen.js';

/**
 * PHASE 2 BRIDGE (2026-09-01) — see the identical helper in server.js/index.js.
 * trustpilot.js's scrapeReviews() is no longer called; creator discovery data
 * already lives in Sheet1 (columns G/J/K/L in the reconciled schema).
 */
function buildCreatorSignal({ nicheTags, platform, subscribers, instagramHandle }) {
  const parts = [];
  if (nicheTags) parts.push(`Niche: ${nicheTags}`);
  if (platform) parts.push(`Platform: ${platform}`);
  if (subscribers) parts.push(`Followers/Subscribers: ${subscribers}`);
  if (instagramHandle) parts.push(`Instagram: ${instagramHandle}`);

  return [{
    rating: '',
    date: '',
    title: nicheTags || 'creator profile',
    text: parts.length > 0 ? parts.join(' | ') : 'No discovery data available for this creator.'
  }];
}

const credentials = JSON.parse(readFileSync(path.resolve('./credentials.json'), 'utf8'));
const auth = new google.auth.GoogleAuth({
  credentials,
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

const sheets = google.sheets({ version: 'v4', auth });
const spreadsheetId = process.env.GOOGLE_SHEET_ID;

async function regenerateEmails() {
  console.log('╔═══════════════════════════════════════════════════════════════╗');
  console.log('║          REGENERATING EMAILS WITH NEW FORMAT                  ║');
  console.log('╚═══════════════════════════════════════════════════════════════╝\n');

  // Get emails tab data
  const emailsData = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: "'emails'!A:G",
  });

  const rows = emailsData.data.values || [];
  if (rows.length <= 1) {
    console.log('No emails to regenerate.');
    return;
  }

  // Get Sheet1 data. NOTE (2026-09-01, Phase 2): this used to read "Sheet1!A:Q"
  // with company/email/website at indices 5/6/13 — that was a leftover from an
  // older, wider Sheet1 layout and did not match the current reconciled schema
  // (A=Status, B=FirstName, C=LastName, D=Company, E=Email, F=Website,
  // G=NicheTags, H=EmailDraft, I=DraftID, J=Platform, K=Subscribers,
  // L=InstagramHandle — see src/sheets.js). Fixed to use the real indices.
  const sheet1Data = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: "Sheet1!A:L",
  });
  const sheet1Rows = sheet1Data.data.values || [];

  // Create lookup by company name
  const companyLookup = {};
  sheet1Rows.slice(1).forEach(row => {
    const company = row[3]?.trim();
    if (company) {
      companyLookup[company] = {
        ceoName: `${row[1] || ''} ${row[2] || ''}`.trim(),
        email: row[4]?.trim() || '',
        website: row[5]?.trim() || '',
        nicheTags: row[6]?.trim() || '',
        platform: row[9]?.trim() || '',
        subscribers: row[10]?.trim() || '',
        instagramHandle: row[11]?.trim() || ''
      };
    }
  });

  console.log(`Found ${rows.length - 1} emails to regenerate.\n`);

  let success = 0;
  let skipped = 0;
  let failed = 0;

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const company = row[0];
    const ceoName = row[1];
    const status = row[6];

    console.log(`━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`);
    console.log(`[${i}/${rows.length - 1}] ${company}`);

    // Get company info from lookup (discovery data lives here now, not a Trustpilot URL)
    const info = companyLookup[company] || { ceoName, email: row[2], website: '' };

    // Skip if no discovery data or already skipped/failed
    if ((!info.nicheTags && !info.platform) || status?.includes('Skipped') || status?.includes('Failed')) {
      console.log('  Skipping - no discovery data (niche/platform) or previously skipped/failed');
      skipped++;
      continue;
    }

    try {
      // Build creator signal from discovery data (no live scraping)
      console.log(`  Creator data: niche="${info.nicheTags || 'n/a'}" platform="${info.platform || 'n/a'}"`);
      const reviews = buildCreatorSignal(info);

      // Analyze pain points
      console.log('  Analyzing pain points...');
      const painPoints = await analyzeReviewsWithAI(reviews, company);

      // Generate new email (variant A — regenerate doesn't need A/B split)
      console.log('  Generating high-converting email...');
      const newEmail = await generateEmail({
        company,
        ceoName: info.ceoName || ceoName,
        reviews,
        variant: 'A'
      });

      // Update the row
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `'emails'!E${i + 1}:F${i + 1}`,
        valueInputOption: 'RAW',
        requestBody: {
          values: [[painPoints, newEmail]]
        }
      });

      console.log('  ✓ Email regenerated');
      success++;

      // Delay between companies
      if (i < rows.length - 1) {
        await new Promise(r => setTimeout(r, 2000));
      }

    } catch (error) {
      console.error(`  ✗ Error: ${error.message}`);
      failed++;
    }
  }

  console.log('\n═══════════════════════════════════════════════════════════════');
  console.log('                        SUMMARY                                 ');
  console.log('═══════════════════════════════════════════════════════════════');
  console.log(`  Regenerated:  ${success}`);
  console.log(`  Skipped:      ${skipped}`);
  console.log(`  Failed:       ${failed}`);
  console.log('═══════════════════════════════════════════════════════════════');
}

regenerateEmails().catch(console.error);
