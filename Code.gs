const USE_QUALTRICS = false; // Set to true to enable Qualtrics integration

const CONDITION_DESCRIPTIONS = {
  '1': 'Section A (Baseline): A minimal 2FA push prompt that only asks "Are you trying to sign in?" with Approve/Deny buttons and provides NO device, location, or IP context.',
  '2': 'Section B (Repeated Requests / MFA Fatigue): Multiple back-to-back 2FA push notifications ("4th sign-in request in the last 2 minutes") for the same login attempt.',
  '3': 'Section C (Number Matching): A 2FA number-matching prompt asking the user to confirm code 47 from an unrecognized mobile browser.',
  '4': 'Section D (High-Context Anomaly): A detailed 2FA push prompt warning of a sign-in attempt from Chrome on Windows 11 in an unfamiliar location (Moscow, Russia, IP 185.220.101.4).'
};

const SYSTEM_PROMPT = `
You are a neutral security notice assistant for a university study on 2FA prompts.
You will receive the visual 2FA prompt shown, the participant's initial decision, their decision latency, and their Annoyance/Confusion and Confidence ratings.

Adapt your notice based on these behavioral signals:
- Fast decision (< 4 seconds) + Approve: Treat as potential habituation. Point out the specific red flag or missing context they overlooked.
- Slow decision (> 10 seconds) OR high Annoyance/Confusion (4-5): Treat as uncertainty. Calmly translate the technical cues or repeated requests into simple risk terms.
- Initial Decision = Deny: Briefly note what specific cue in the prompt warranted caution.

Rules:
- Generate ONLY the warning notice (1 to 2 short sentences, under 30 words total).
- Do NOT affirm or praise their decision. Do NOT ask follow-up questions. Do NOT mention their survey numbers/seconds back to them.
- Simply state the suspicious detail(s) and let the user decide.
`;

function doGet(e) {
  if (USE_QUALTRICS) {
    const params = (e && e.parameter) || {};
    const rid = params.rid || '';
    const returnUrl = 'https://utk.co1.qualtrics.com/jfe/form/SV_5dopubFWkIN8cd0?rid=' + encodeURIComponent(rid);
    Logger.log(returnUrl);

    const template = HtmlService.createTemplateFromFile('Index');
    template.rid = rid;
    template.returnUrl = returnUrl;
    return template
      .evaluate()
      .setTitle('COSC 469/569 Class Project — 2FA Prompt Study')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('COSC 469/569 Class Project — 2FA Prompt Study')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

function generateWarning(studyData) {
  const apiKey = (PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '').trim();
  if (!apiKey) {
    throw new Error('GEMINI_API_KEY is missing in Project Settings > Script Properties.');
  }

  const conditionLabel = CONDITION_DESCRIPTIONS[studyData.condition] || ('Section ' + studyData.condition);
  const promptText = `Participant Telemetry & Ratings:
- Prompt Shown: ${conditionLabel}
- Initial Decision: ${studyData.decision}
- Decision Latency: ${studyData.decisionTimeSec} seconds
- Annoyed/Confused (1-5): ${studyData.annoyed}
- Confidence (1-5): ${studyData.confidence}

Generate the adaptive 2FA security warning now.`;

  const payload = {
    system_instruction: { parts: [{ text: SYSTEM_PROMPT }] },
    contents: [{ role: 'user', parts: [{ text: promptText }] }],
    generationConfig: {
      maxOutputTokens: 60
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const models = ['gemini-3.5-flash-lite', 'gemini-3.8-flash'];
  const host = 'generative' + 'language.' + 'googleapis.com';
  let responseCode = 0;
  let responseText = '';

  for (let i = 0; i < models.length; i++) {
    const model = models.at(i);
    const url = 'https://' + host + '/v1beta/models/' + model + ':generateContent';

    const httpResponse = UrlFetchApp.fetch(url, options);
    responseCode = httpResponse.getResponseCode();
    responseText = httpResponse.getContentText();

    if (responseCode === 200) {
      break;
    }
    if (responseCode === 503 || responseCode === 429) {
      Utilities.sleep(1500);
    }
  }

  if (responseCode !== 200) {
    throw new Error('Gemini API Error (' + responseCode + '): ' + responseText);
  }

  const response = JSON.parse(responseText);
  return response.candidates.at(0).content.parts.at(0).text.trim();
}

function saveFinalResponse(studyData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Responses');
  if (USE_QUALTRICS) {
    sheet.appendRow([
      new Date(),
      studyData.rid,
      'Condition ' + studyData.condition,
      studyData.decision,
      studyData.decisionTimeSec,
      studyData.annoyed,
      studyData.confidence,
      studyData.warning,
      studyData.finalDecision
    ]);
  } else {
    sheet.appendRow([
      new Date(),
      'Condition ' + studyData.condition,
      studyData.decision,
      studyData.decisionTimeSec,
      studyData.annoyed,
      studyData.confidence,
      studyData.warning,
      studyData.finalDecision
    ]);
  }
  return true;
}

function generateSimulatedData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Responses');
  if (!sheet) {
    sheet = ss.insertSheet('Responses');
  }

  sheet.clear();
  if (USE_QUALTRICS) {
    sheet.appendRow([
      'Timestamp',
      'RunID',
      'Condition',
      'Initial_Decision',
      'Decision_Time_Sec',
      'Annoyed_Confused',
      'Confidence',
      'Generated_Warning',
      'Final_Decision'
    ]);
  } else {
    sheet.appendRow([
      'Timestamp',
      'Condition',
      'Initial_Decision',
      'Decision_Time_Sec',
      'Annoyed_Confused',
      'Confidence',
      'Generated_Warning',
      'Final_Decision'
    ]);
  }

  const profiles = {
    'Condition 1': {
      initApproveProb: 0.75,
      flipToDenyProb: 0.60,
      timeRange: [1.4, 5.2],
      annoyedRange: [1, 3],
      confRange: [3, 5],
      warning: 'This login request provides no location or device details. Approving a blind prompt grants full account access if you did not just enter your password.'
    },
    'Condition 2': {
      initApproveProb: 0.55,
      flipToDenyProb: 0.75,
      timeRange: [2.1, 8.5],
      annoyedRange: [4, 5],
      confRange: [1, 3],
      warning: 'Receiving multiple rapid sign-in requests is a common MFA bombing tactic. Do not approve just to stop the alerts—tap Deny if you are not actively signing in.'
    },
    'Condition 3': {
      initApproveProb: 0.35,
      flipToDenyProb: 0.80,
      timeRange: [5.5, 14.0],
      annoyedRange: [2, 4],
      confRange: [3, 5],
      warning: 'Only approve number-matching prompts when the exact two-digit code appears on the screen in front of you right now.'
    },
    'Condition 4': {
      initApproveProb: 0.25,
      flipToDenyProb: 0.90,
      timeRange: [3.8, 11.5],
      annoyedRange: [2, 4],
      confRange: [4, 5],
      warning: 'Security Alert: This login originates from Chrome on Windows 11 in Moscow, Russia. Unless you are using a VPN there, tap Deny immediately.'
    }
  };

  const rows = [];
  const baseTime = new Date().getTime() - (80 * 60000);
  let rowIndex = 0;

  Object.keys(profiles).forEach(cond => {
    const p = profiles[cond];
    for (let i = 0; i < 20; i++) {
      const ts = new Date(baseTime + rowIndex * 60000);
      const initApprove = Math.random() < p.initApproveProb;
      const initialDecision = initApprove ? 'Approve' : 'Deny';

      let latency = p.timeRange[0] + Math.random() * (p.timeRange[1] - p.timeRange[0]);
      if (initApprove && cond === 'Condition 1' && Math.random() < 0.5) {
        latency = 1.2 + Math.random() * 1.8;
      }
      latency = Number(latency.toFixed(2));

      const annoyed = Math.floor(p.annoyedRange[0] + Math.random() * (p.annoyedRange[1] - p.annoyedRange[0] + 1));
      const confidence = Math.floor(p.confRange[0] + Math.random() * (p.confRange[1] - p.confRange[0] + 1));

      let finalDecision = initialDecision;
      if (initialDecision === 'Approve' && Math.random() < p.flipToDenyProb) {
        finalDecision = 'Deny';
      }

      if (USE_QUALTRICS) {
        rows.push(['', ts, cond, initialDecision, latency, annoyed, confidence, p.warning, finalDecision]);
      } else {
        rows.push([ts, cond, initialDecision, latency, annoyed, confidence, p.warning, finalDecision]);
      }
      rowIndex++;
    }
  });

  const numCols = USE_QUALTRICS ? 9 : 8;
  sheet.getRange(2, 1, rows.length, numCols).setValues(rows);
  Logger.log('Successfully generated 80 simulated responses (20 per condition).');
}