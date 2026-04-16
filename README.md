# Gnome Flash — Form-to-Slide Automation

Automatically turns a Google Form submission into a formatted Google Slide and posts a Slack notification with a link.

---

## How it works

1. Someone fills out the **Google Form**
2. Apps Script detects the submission via a trigger
3. The script **copies your Slides template** and fills in all the form answers
4. The new slide is saved to a **specific Google Drive folder**
5. A **Slack message** is posted with the slide title and a direct link

---

## Files in this repo

| File | Purpose |
|---|---|
| `Code.gs` | All the automation logic — edit this to customize |
| `appsscript.json` | Apps Script project manifest (permissions) |
| `.clasp.json` | Config for pushing/pulling code with `clasp` CLI |

---

## One-time setup

### Step 1 — Create your Google Form

1. Go to [Google Forms](https://forms.google.com) and create a new form
2. Add questions for all the information you want on each slide (e.g. Name, Department, Headline, Date, Description, Contact, Photo)
3. Note the **Form ID** from the URL:
   ```
   docs.google.com/forms/d/FORM_ID_IS_HERE/edit
   ```

### Step 2 — Create your Slides template

1. Go to [Google Slides](https://slides.google.com) and design your template slide
2. In every text box where you want form data to appear, type the matching placeholder exactly:

   **Required** — always appear on every slide:

   | Placeholder | Maps to |
   |---|---|
   | `{{TEAM_NAME}}` | "What is your team?" |
   | `{{FUNCTION_DESCRIPTION}}` | "In a sentence or two, what does your function do?" |
   | `{{SUBMISSION_DATE}}` | Auto-filled with the form submission date |

   **Optional** — if the question is left blank, the **entire shape is deleted** from the slide. Each optional answer must be in its own text box in the template (don't mix with required text):

   | Placeholder | Maps to |
   |---|---|
   | `{{TEAM_FOCUS}}` | "What's your team focused on or excited about in the coming months?" |
   | `{{RECENT_WIN}}` | "What's a recent win, milestone, or accomplishment…" |
   | `{{SHOUTOUTS}}` | "Any shoutouts? Recognize someone on your team…" |
   | `{{EXTERNAL_RECOGNITION}}` | "Has your team presented, published, or been recognized externally…" |
   | `{{FUN_FACT}}` | "What's something people outside your function probably don't know…" |
   | `{{PHOTO}}` | Photo file upload *(see below)* |

3. **For the photo placeholder:** Insert a rectangle or image shape where you want the photo to go → **Format > Alt Text** → set *Description* to `{{PHOTO}}`
4. Note the **Template Slide ID** from the URL:
   ```
   docs.google.com/presentation/d/TEMPLATE_SLIDE_ID_IS_HERE/edit
   ```

### Step 3 — Set up the destination Drive folder

1. In Google Drive, create (or navigate to) the folder where finished slides should be saved
2. Note the **Folder ID** from the URL:
   ```
   drive.google.com/drive/folders/FOLDER_ID_IS_HERE
   ```

### Step 4 — Create a Slack Incoming Webhook

1. Go to [api.slack.com/apps](https://api.slack.com/apps) → **Create New App** → *From scratch*
2. Give it a name (e.g. "Gnome Flash Bot") and pick your workspace
3. Under **Features**, click **Incoming Webhooks** → toggle it On
4. Click **Add New Webhook to Workspace** → select the channel to post in → Allow
5. Copy the **Webhook URL** (starts with `https://hooks.slack.com/services/...`)

### Step 5 — Create the Apps Script project

1. Go to [script.google.com](https://script.google.com) → **New project**
2. Delete any placeholder code
3. Copy-paste the contents of `Code.gs` into the editor
4. Click the gear icon (Project Settings) → check **Show "appsscript.json" manifest file in editor**
5. Click on `appsscript.json` in the editor and paste in the contents of `appsscript.json` from this repo
6. Fill in the four values at the top of `Code.gs` inside the `CONFIG` block:

   ```javascript
   const CONFIG = {
     TEMPLATE_SLIDE_ID:    'paste-your-template-id-here',
     DESTINATION_FOLDER_ID: 'paste-your-folder-id-here',
     SLACK_WEBHOOK_URL:    'paste-your-webhook-url-here',
     SLACK_BOT_NAME:       'Gnome Flash Bot',
     SLIDE_NAME_PREFIX:    'Gnome Flash — ',
   };
   ```

7. Update the `FORM_FIELDS` object so each value matches your form's question titles **exactly**

### Step 6 — Connect the form trigger

1. In the Apps Script editor, open the `setupFormTrigger` function
2. Replace `'YOUR_FORM_ID_HERE'` with your actual Form ID
3. Run `setupFormTrigger` (click the play button with that function selected)
4. Authorize the requested permissions when prompted
5. The trigger is now live — every future form submission will fire the automation

---

## Testing before going live

Run **`testWithSampleData()`** from the Apps Script editor.  
It uses dummy data to create a real test slide in your Drive folder and sends a real Slack message, so you can verify everything looks right without submitting the form.

Run **`auditTemplatePlaceholders()`** to print all text elements in your template to the log and confirm the placeholder strings are in place.

---

## Customizing the form fields

To add a new field:

1. Add the question to your Google Form
2. Add a matching entry to `FORM_FIELDS` in `Code.gs`:
   ```javascript
   MY_NEW_FIELD: 'Exact Question Title From Form',
   ```
3. Add a placeholder to `PLACEHOLDERS`:
   ```javascript
   MY_NEW_FIELD: '{{MY_NEW_FIELD}}',
   ```
4. Add a row to `replacePlaceholders()` in `Code.gs`:
   ```javascript
   { find: PLACEHOLDERS.MY_NEW_FIELD, replace: formData[FORM_FIELDS.MY_NEW_FIELD] || '' },
   ```
5. Add `{{MY_NEW_FIELD}}` to the matching spot in your Slides template

---

## Optional: push code from this repo using `clasp`

[`clasp`](https://github.com/google/clasp) lets you edit `Code.gs` locally and push changes to Apps Script.

```bash
npm install -g @google/clasp
clasp login
```

Then fill in the `scriptId` in `.clasp.json` (get it from Apps Script → Project Settings → Script ID) and run:

```bash
clasp push   # upload local changes to Apps Script
clasp pull   # download from Apps Script to local
```

---

## Troubleshooting

| Symptom | Fix |
|---|---|
| Slide created but placeholders not replaced | Check that `FORM_FIELDS` values match form question titles exactly (case-sensitive) |
| No slide created after form submit | Make sure `setupFormTrigger()` was run and the trigger appears in **Triggers** (clock icon in Apps Script) |
| Slack message not sent | Verify `SLACK_WEBHOOK_URL` is set and the Slack app is still active |
| Photo not inserted | Confirm the placeholder shape's Alt Text Description is set to `{{PHOTO}}` and the form question title matches `FORM_FIELDS.PHOTO` |
| Permission denied errors | Re-run `setupFormTrigger()` and accept all OAuth prompts |
