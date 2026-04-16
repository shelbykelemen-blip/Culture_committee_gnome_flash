// =============================================================================
// Gnome Flash — Team Spotlight | Google Form → Google Slides Automation
// Culture Committee | Google Apps Script
//
// HOW IT WORKS:
//   1. Someone fills out the "Team Spotlight" Google Form
//   2. This script runs automatically via a form-submit trigger
//   3. It copies your Slides template and fills in the form data
//   4. The new slide is saved to the designated Google Drive folder
//   5. A Slack message is posted with a link to the new slide
//
// SETUP STEPS: See README.md for full instructions.
// =============================================================================


// =============================================================================
// CONFIGURATION — fill these in before deploying
// =============================================================================

const CONFIG = {
  // ID of your Google Slides template
  // → Open the template in Google Slides, copy the ID from the URL:
  //   docs.google.com/presentation/d/THIS_PART_HERE/edit
  TEMPLATE_SLIDE_ID: '1hMTv31DchU6CkiOG_CM7TYG8PJ4DQccn0rIWZH0WJWI',

  // ID of the Google Drive folder where new slides will be saved
  // → Open the destination folder in Drive, copy the ID from the URL:
  //   drive.google.com/drive/folders/THIS_PART_HERE
  DESTINATION_FOLDER_ID: '1qAiI5Ks0dc7LTOgo5Y5eWPS9p2o6fZF_',

  // Slack Incoming Webhook URL
  // → Create one at: api.slack.com → Your Apps → Incoming Webhooks → Add New Webhook
  SLACK_WEBHOOK_URL: 'YOUR_SLACK_WEBHOOK_URL_HERE',

  // Display name shown in the Slack message
  SLACK_BOT_NAME: 'Gnome Flash Bot',

  // Prefix used in the title of each saved Slides file
  SLIDE_NAME_PREFIX: 'Team Spotlight — ',
};


// =============================================================================
// FORM FIELD NAMES
// These match the exact question titles in the "Team Spotlight" Google Form.
// If you rename a question in the form, update it here too (case-sensitive).
// =============================================================================

const FORM_FIELDS = {
  // Required fields — always filled in on the slide
  TEAM_NAME:            'What is your team?',
  FUNCTION_DESCRIPTION: 'In a sentence or two, what does your function do?',

  // Optional fields — if left blank, their shapes are removed from the slide
  TEAM_FOCUS:           "What's your team focused on or excited about in the coming months?",
  RECENT_WIN:           "What's a recent win, milestone, or accomplishment you'd like the company to know about?",
  SHOUTOUTS:            'Any shoutouts? Recognize someone on your team (or a cross-functional partner) who deserves a callout.',
  EXTERNAL_RECOGNITION: "Has your team presented, published, or been recognized externally recently — or is something coming up?",
  FUN_FACT:             "What's something people outside your function probably don't know or would find surprising about what you do?",
  PHOTO:                'Got a team photo to share? (optional)',
};


// =============================================================================
// SLIDE TEMPLATE PLACEHOLDERS
// Add text boxes to your Google Slides template containing these exact strings.
// This script replaces each one with the matching form answer.
//
// REQUIRED placeholders (always present on every slide):
//   {{TEAM_NAME}}             — team name
//   {{FUNCTION_DESCRIPTION}}  — what the team/function does
//   {{SUBMISSION_DATE}}       — auto-filled with the form submission date
//
// OPTIONAL placeholders (shape is DELETED from the slide if the question was left blank):
//   {{TEAM_FOCUS}}            — upcoming priorities
//   {{RECENT_WIN}}            — recent accomplishment
//   {{SHOUTOUTS}}             — callouts / recognition
//   {{EXTERNAL_RECOGNITION}}  — conferences, awards, publications
//   {{FUN_FACT}}              — surprising fact about the team
//
// IMPORTANT for optional placeholders:
//   Each optional answer must live in its own self-contained shape (text box or
//   group) in the template. If the submitter leaves that question blank:
//     1. That shape is removed from the slide.
//     2. The remaining shapes are automatically redistributed to fill the gap —
//        no empty space is left behind.
//   You can include a label like "Recent Win:" inside the same shape; the whole
//   shape disappears cleanly when the answer is empty.
//   Stack all optional shapes in a vertical column in the template. The script
//   uses the top of the topmost shape and the bottom of the bottommost shape
//   as the zone, then spreads survivors evenly within it.
//
// TIP: For the photo, insert a rectangle shape where the image should go,
//      then Format > Alt Text → set Description to {{PHOTO}}.
// =============================================================================

const PLACEHOLDERS = {
  // Required
  TEAM_NAME:            '{{TEAM_NAME}}',
  FUNCTION_DESCRIPTION: '{{FUNCTION_DESCRIPTION}}',
  SUBMISSION_DATE:      '{{SUBMISSION_DATE}}',
  // Optional — shapes containing these are removed when the answer is blank
  TEAM_FOCUS:           '{{TEAM_FOCUS}}',
  RECENT_WIN:           '{{RECENT_WIN}}',
  SHOUTOUTS:            '{{SHOUTOUTS}}',
  EXTERNAL_RECOGNITION: '{{EXTERNAL_RECOGNITION}}',
  FUN_FACT:             '{{FUN_FACT}}',
  // Photo — matched by alt-text on an image/shape element
  PHOTO:                '{{PHOTO}}',
};


// =============================================================================
// MAIN TRIGGER — runs automatically on every form submission
// Wire this up by running setupFormTrigger() once (see below).
// =============================================================================

/**
 * Entry point called by the "On form submit" trigger.
 * @param {Object} e - The Apps Script form-submit event object
 */
function onFormSubmit(e) {
  try {
    const formData = extractFormData(e.response);
    Logger.log('New Team Spotlight submission received.');

    const newFile = createSlideFromTemplate(formData);
    Logger.log('Slide created: ' + newFile.getName());

    sendSlackNotification(newFile, formData);

  } catch (err) {
    Logger.log('ERROR in onFormSubmit: ' + err.toString());
    sendSlackErrorNotification(err);
  }
}


// =============================================================================
// CORE FUNCTIONS
// =============================================================================

/**
 * Pulls all question/answer pairs out of the form response into a plain object.
 * @param {FormResponse} formResponse
 * @returns {Object} key = question title, value = submitted answer
 */
function extractFormData(formResponse) {
  const data = {};

  formResponse.getItemResponses().forEach(function (itemResponse) {
    const question = itemResponse.getItem().getTitle();
    const answer   = itemResponse.getResponse();
    // File upload answers are arrays of Drive file IDs — join them if multiple
    data[question] = Array.isArray(answer) ? answer.join(', ') : (answer || '');
  });

  // Metadata extras
  data['_timestamp'] = Utilities.formatDate(
    formResponse.getTimestamp(),
    Session.getScriptTimeZone(),
    'MMMM d, yyyy'
  );
  data['_email'] = formResponse.getRespondentEmail() || '';

  return data;
}

/**
 * Copies the template, fills in all placeholders, and moves the file to the
 * destination folder. Returns the new Drive File object.
 * @param {Object} formData
 * @returns {File} the newly created Drive file
 */
function createSlideFromTemplate(formData) {
  const templateFile      = DriveApp.getFileById(CONFIG.TEMPLATE_SLIDE_ID);
  const destinationFolder = DriveApp.getFolderById(CONFIG.DESTINATION_FOLDER_ID);
  const slideName         = buildSlideName(formData);

  const newFile      = templateFile.makeCopy(slideName, destinationFolder);
  const presentation = SlidesApp.openById(newFile.getId());

  replaceRequiredFields(presentation, formData);   // always-present fields
  processOptionalFields(presentation, formData);   // remove blanks + repack
  handlePhotoUpload(presentation, formData);       // swap photo placeholder
  presentation.saveAndClose();

  return newFile;
}

/**
 * Builds a human-readable file name for the new slide.
 * Uses the team name as the primary identifier.
 * @param {Object} formData
 * @returns {string}
 */
function buildSlideName(formData) {
  const teamName  = formData[FORM_FIELDS.TEAM_NAME] || 'Unknown Team';
  const dateStamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return CONFIG.SLIDE_NAME_PREFIX + teamName + ' — ' + dateStamp;
}

/**
 * Replaces all required placeholder text with form answers.
 * Required fields are always present — they are never removed from the slide.
 * @param {Presentation} presentation
 * @param {Object} formData
 */
function replaceRequiredFields(presentation, formData) {
  const required = [
    { find: PLACEHOLDERS.TEAM_NAME,            replace: formData[FORM_FIELDS.TEAM_NAME]            || '' },
    { find: PLACEHOLDERS.FUNCTION_DESCRIPTION, replace: formData[FORM_FIELDS.FUNCTION_DESCRIPTION] || '' },
    { find: PLACEHOLDERS.SUBMISSION_DATE,      replace: formData['_timestamp']                          },
    // ↓ Add more required fields here as your form grows
  ];

  required.forEach(function (r) {
    presentation.replaceAllText(r.find, r.replace);
  });

  Logger.log('Required fields replaced.');
}

/**
 * Handles optional fields in three phases so the slide never has empty gaps:
 *
 *  Phase 1 — Scan: Before touching anything, find every shape that contains an
 *             optional placeholder and record its position and size.
 *
 *  Phase 2 — Remove: Delete shapes whose question was left blank.
 *
 *  Phase 3 — Repack: Redistribute the surviving shapes evenly within the
 *             vertical zone that all optional shapes originally occupied,
 *             closing any gaps left by the removed shapes.
 *
 * Template requirement: each optional placeholder must be in its own
 * dedicated shape (text box or group). You can put a label like
 * "Recent Win:" inside the same shape — it all disappears together cleanly.
 * Do NOT put two different optional placeholders in the same shape.
 *
 * @param {Presentation} presentation
 * @param {Object} formData
 */
function processOptionalFields(presentation, formData) {
  const optionalFields = [
    { placeholder: PLACEHOLDERS.TEAM_FOCUS,           value: formData[FORM_FIELDS.TEAM_FOCUS]           || '' },
    { placeholder: PLACEHOLDERS.RECENT_WIN,           value: formData[FORM_FIELDS.RECENT_WIN]           || '' },
    { placeholder: PLACEHOLDERS.SHOUTOUTS,            value: formData[FORM_FIELDS.SHOUTOUTS]            || '' },
    { placeholder: PLACEHOLDERS.EXTERNAL_RECOGNITION, value: formData[FORM_FIELDS.EXTERNAL_RECOGNITION] || '' },
    { placeholder: PLACEHOLDERS.FUN_FACT,             value: formData[FORM_FIELDS.FUN_FACT]             || '' },
    // ↓ Add more optional fields here as your form grows
  ];

  // Work slide-by-slide so position math stays within the same slide
  presentation.getSlides().forEach(function (slide) {

    // ── Phase 1: Scan ────────────────────────────────────────────────────────
    // Map each element to the optional field whose placeholder it contains.
    // We record original top/height now because Phase 2 will delete some.
    const entries = []; // { element, top, height, field, keep }

    slide.getPageElements().forEach(function (element) {
      try {
        const text = element.asShape().getText().asString();
        for (var i = 0; i < optionalFields.length; i++) {
          if (text.indexOf(optionalFields[i].placeholder) !== -1) {
            entries.push({
              element: element,
              top:     element.getTop(),
              height:  element.getHeight(),
              field:   optionalFields[i],
              keep:    optionalFields[i].value !== '',
            });
            break; // one shape = one field
          }
        }
      } catch (e) {
        // No text body (image, line, etc.) — skip
      }
    });

    if (entries.length === 0) return; // no optional sections on this slide

    // Sort by original vertical position so repacking preserves reading order
    entries.sort(function (a, b) { return a.top - b.top; });

    // Record the zone bounds before any removals
    const zoneTop    = entries[0].top;
    const zoneBottom = entries[entries.length - 1].top + entries[entries.length - 1].height;

    // ── Phase 2: Remove ──────────────────────────────────────────────────────
    entries.forEach(function (entry) {
      if (!entry.keep) {
        entry.element.remove();
        Logger.log('Removed blank optional section: ' + entry.field.placeholder);
      }
    });

    // ── Phase 2b: Replace text in surviving shapes ───────────────────────────
    // Do this after removals so replaceAllText only touches the remaining shapes.
    entries.forEach(function (entry) {
      if (entry.keep) {
        presentation.replaceAllText(entry.field.placeholder, entry.field.value);
      }
    });

    // ── Phase 3: Repack ──────────────────────────────────────────────────────
    // Redistribute surviving shapes evenly within the original zone,
    // eliminating any gaps left by the removed shapes.
    const survivors = entries.filter(function (e) { return e.keep; });

    if (survivors.length === 0) return;

    const totalContentHeight = survivors.reduce(function (sum, e) { return sum + e.height; }, 0);
    const totalZoneHeight    = zoneBottom - zoneTop;
    // Spread leftover space evenly between items (not around edges)
    const gap = survivors.length > 1
      ? (totalZoneHeight - totalContentHeight) / (survivors.length - 1)
      : 0;

    var currentY = zoneTop;
    survivors.forEach(function (entry) {
      entry.element.setTop(currentY);
      currentY += entry.height + gap;
    });

    Logger.log(
      'Repacked optional sections: ' + survivors.length + ' kept of ' +
      entries.length + ' total. Gap between items: ' + Math.round(gap) + 'pt.'
    );
  });
}

/**
 * If the form included a photo upload, finds the {{PHOTO}} placeholder shape
 * in the template (matched by alt-text) and swaps it for the actual image.
 *
 * To set up in your template:
 *   Insert a rectangle/image placeholder → Format > Alt Text → set Description to {{PHOTO}}
 *
 * @param {Presentation} presentation
 * @param {Object} formData
 */
function handlePhotoUpload(presentation, formData) {
  const photoAnswer = formData[FORM_FIELDS.PHOTO];
  if (!photoAnswer) {
    Logger.log('No photo provided — skipping image replacement.');
    return;
  }

  // Form file-upload answers are Drive file IDs (sometimes comma-separated)
  const fileId = photoAnswer.split(',')[0].trim();

  try {
    const imageBlob = DriveApp.getFileById(fileId).getBlob();

    presentation.getSlides().forEach(function (slide) {
      slide.getPageElements().forEach(function (element) {
        // Match the placeholder by its alt-text description
        if (getElementAltText(element) === PLACEHOLDERS.PHOTO) {
          const left   = element.getLeft();
          const top    = element.getTop();
          const width  = element.getWidth();
          const height = element.getHeight();
          element.remove();
          slide.insertImage(imageBlob, left, top, width, height);
          Logger.log('Photo inserted into slide.');
        }
      });
    });
  } catch (err) {
    // Non-fatal — log and continue without the photo
    Logger.log('Could not insert photo (file ID: ' + fileId + '): ' + err.toString());
  }
}

/**
 * Helper: safely reads the alt-text description of any page element.
 * Returns an empty string if the element type doesn't support alt-text.
 * @param {PageElement} element
 * @returns {string}
 */
function getElementAltText(element) {
  try {
    return element.getDescription() || '';
  } catch (e) {
    return '';
  }
}


// =============================================================================
// SLACK NOTIFICATION
// =============================================================================

/**
 * Posts a formatted Slack message to the configured webhook.
 * @param {File}   slideFile  Drive file for the newly created slide
 * @param {Object} formData
 */
function sendSlackNotification(slideFile, formData) {
  if (!CONFIG.SLACK_WEBHOOK_URL || CONFIG.SLACK_WEBHOOK_URL === 'YOUR_SLACK_WEBHOOK_URL_HERE') {
    Logger.log('Slack webhook not configured — skipping notification.');
    return;
  }

  const slideUrl      = 'https://docs.google.com/presentation/d/' + slideFile.getId() + '/edit';
  const teamName      = formData[FORM_FIELDS.TEAM_NAME]            || 'Unknown Team';
  const functionBlurb = formData[FORM_FIELDS.FUNCTION_DESCRIPTION] || '';
  // Show a short preview — first sentence or first 120 chars
  const preview = functionBlurb.split(/[.!?]/)[0].trim() || functionBlurb.substring(0, 120).trim();

  const payload = {
    username: CONFIG.SLACK_BOT_NAME,
    icon_emoji: ':gnome:',
    blocks: [
      {
        type: 'header',
        text: { type: 'plain_text', text: ':spotlight: New Team Spotlight Slide Added!', emoji: true },
      },
      {
        type: 'section',
        fields: [
          { type: 'mrkdwn', text: '*Team:*\n' + teamName },
          { type: 'mrkdwn', text: '*What they do:*\n' + (preview || '—') },
        ],
      },
      {
        type: 'section',
        text: { type: 'mrkdwn', text: '<' + slideUrl + '|:link:  Open slide in Google Drive>' },
      },
      { type: 'divider' },
    ],
  };

  const response = UrlFetchApp.fetch(CONFIG.SLACK_WEBHOOK_URL, {
    method:          'post',
    contentType:     'application/json',
    payload:         JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  if (response.getResponseCode() === 200) {
    Logger.log('Slack notification sent.');
  } else {
    Logger.log('Slack notification failed (' + response.getResponseCode() + '): ' + response.getContentText());
  }
}

/**
 * Posts a plain error alert to Slack. Non-fatal — won't throw.
 * @param {Error} err
 */
function sendSlackErrorNotification(err) {
  if (!CONFIG.SLACK_WEBHOOK_URL || CONFIG.SLACK_WEBHOOK_URL === 'YOUR_SLACK_WEBHOOK_URL_HERE') return;

  try {
    UrlFetchApp.fetch(CONFIG.SLACK_WEBHOOK_URL, {
      method:      'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        username: CONFIG.SLACK_BOT_NAME,
        text: ':warning: *Team Spotlight error* — a form submission could not be processed.\n```' + err.toString() + '```',
      }),
      muteHttpExceptions: true,
    });
  } catch (e) {
    Logger.log('Could not send Slack error notification: ' + e.toString());
  }
}


// =============================================================================
// ONE-TIME SETUP UTILITIES
// Run these manually from the Apps Script editor (not as triggers).
// =============================================================================

/**
 * RUN ONCE: Connects this script to your Google Form so onFormSubmit() fires
 * automatically every time someone submits the form.
 *
 * How to find your Form ID:
 *   Open the form in Google Forms → copy the ID from the URL:
 *   docs.google.com/forms/d/THIS_PART_HERE/edit
 */
function setupFormTrigger() {
  const FORM_ID = '136YDrb8SRjMO5okLu5fhaxv0iX44vpb3v839-vAKpgo';

  // Remove any existing form-submit triggers to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT) {
      ScriptApp.deleteTrigger(t);
      Logger.log('Removed existing trigger: ' + t.getUniqueId());
    }
  });

  ScriptApp.newTrigger('onFormSubmit')
    .forForm(FORM_ID)
    .onFormSubmit()
    .create();

  Logger.log('Trigger created. onFormSubmit() will now fire for form: ' + FORM_ID);
}

/**
 * RUN TO TEST: Simulates a form submission with dummy data so you can verify
 * the slide is created and Slack is notified — without submitting the real form.
 */
function testWithSampleData() {
  // To test the "optional field removed when blank" behavior, set any optional
  // field to '' below and verify that shape disappears on the generated slide.
  const sampleData = {
    // Required
    [FORM_FIELDS.TEAM_NAME]:            'Data Engineering',
    [FORM_FIELDS.FUNCTION_DESCRIPTION]: 'We build and maintain the data pipelines that power company-wide reporting and analytics.',
    // Optional — leave any of these as '' to test blank-removal
    [FORM_FIELDS.TEAM_FOCUS]:           'Launching a real-time dashboard for the sales team and migrating our ETL jobs to a new orchestration platform.',
    [FORM_FIELDS.RECENT_WIN]:           'Reduced our nightly pipeline runtime from 4 hours to 45 minutes — a 10× improvement that unblocked the finance team.',
    [FORM_FIELDS.SHOUTOUTS]:            '',  // ← blank: shape should be removed from the slide
    [FORM_FIELDS.EXTERNAL_RECOGNITION]: 'Two of our engineers will be speaking at DataConf in May!',
    [FORM_FIELDS.FUN_FACT]:             "We process over 2 billion events a day — and most people think we just make spreadsheets.",
    '_timestamp':                       'April 14, 2026',
    '_email':                           '',
  };

  Logger.log('Running test with sample data...');
  const newFile = createSlideFromTemplate(sampleData);
  Logger.log('Test slide created: ' + newFile.getName());
  Logger.log('URL: https://docs.google.com/presentation/d/' + newFile.getId() + '/edit');
  sendSlackNotification(newFile, sampleData);
}

/**
 * RUN TO AUDIT: Prints every text element in your template slide to the log,
 * so you can confirm all {{PLACEHOLDER}} strings are present and spelled correctly.
 */
function auditTemplatePlaceholders() {
  if (CONFIG.TEMPLATE_SLIDE_ID === 'YOUR_TEMPLATE_SLIDE_ID_HERE') {
    Logger.log('Set TEMPLATE_SLIDE_ID in CONFIG first.');
    return;
  }

  const presentation = SlidesApp.openById(CONFIG.TEMPLATE_SLIDE_ID);
  Logger.log('Auditing template: ' + presentation.getName());

  presentation.getSlides().forEach(function (slide, idx) {
    Logger.log('--- Slide ' + (idx + 1) + ' ---');
    slide.getPageElements().forEach(function (el) {
      try {
        const text = el.asShape().getText().asString().trim();
        if (text) Logger.log('  Text: ' + text);
      } catch (e) { /* element has no text body */ }

      const alt = getElementAltText(el);
      if (alt) Logger.log('  Alt-text: ' + alt);
    });
  });
}
