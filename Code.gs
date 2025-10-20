const SETTINGS_KEY_PREFIX = 'QUIZ_CONTROL_SETTINGS_';
const DEFAULT_SETTINGS = {
  durationMinutes: 0,
  forceRetakeEnabled: false,
  scoreThreshold: 80,
  timerTriggerId: '',
  sessionEndsAt: ''
};

function onOpen(e) {
  let context;
  try {
    context = getActiveFormContext();
  } catch (err) {
    return;
  }
  const ui = context.ui;
  ui.createMenu('Quiz Control')
    .addItem('Configure settings', 'showSettingsSidebar')
    .addItem('Start timed session', 'startTimedSession')
    .addItem('Cancel active timer', 'cancelTimer')
    .addToUi();
  ensureQuizConfiguration(context.form);
  ensureSubmissionTrigger(context.formId, getSettings(context.formId).forceRetakeEnabled);
}

function onInstall(e) {
  onOpen(e);
}

function showSettingsSidebar() {
  const context = getActiveFormContext();
  const template = HtmlService.createTemplateFromFile('TimerConfig');
  template.settings = getSettings(context.formId);
  const html = template.evaluate()
    .setTitle('Quiz Control Settings')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  context.ui.showSidebar(html);
}

function getSettings(formId) {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(SETTINGS_KEY_PREFIX + formId);
  if (!raw) {
    return Object.assign({}, DEFAULT_SETTINGS);
  }
  try {
    const parsed = JSON.parse(raw);
    return Object.assign({}, DEFAULT_SETTINGS, parsed);
  } catch (err) {
    return Object.assign({}, DEFAULT_SETTINGS);
  }
}

function saveSettings(formId, settings) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty(SETTINGS_KEY_PREFIX + formId, JSON.stringify(settings));
}

function updateSettings(data) {
  const context = getActiveFormContext();
  const formId = context.formId;
  const settings = getSettings(formId);
  settings.durationMinutes = Number(data.durationMinutes) || 0;
  settings.forceRetakeEnabled = data.forceRetakeEnabled === true || data.forceRetakeEnabled === 'true';
  settings.scoreThreshold = Number(data.scoreThreshold) || 80;
  saveSettings(formId, settings);
  ensureSubmissionTrigger(formId, settings.forceRetakeEnabled);
  return settings;
}

function startTimedSession() {
  const context = getActiveFormContext();
  const form = context.form;
  const formId = context.formId;
  const settings = getSettings(formId);
  if (!settings.durationMinutes || settings.durationMinutes <= 0) {
    throw new Error('Configure a timer duration before starting a session.');
  }
  cancelTimerInternal(formId, null, true);
  form.setAcceptingResponses(true);
  const endTime = new Date(Date.now() + settings.durationMinutes * 60000);
  const trigger = ScriptApp.newTrigger('closeFormAfterTimer')
    .timeBased()
    .after(settings.durationMinutes * 60000)
    .create();
  settings.timerTriggerId = trigger.getUniqueId();
  settings.sessionEndsAt = endTime.toISOString();
  saveSettings(formId, settings);
  context.ui.alert('Timed session started. The form will close at ' + endTime.toLocaleString());
}

function cancelTimer(silent) {
  const context = getActiveFormContext();
  cancelTimerInternal(context.formId, context.ui, silent);
}

function closeFormAfterTimer(e) {
  const triggerUid = e && e.triggerUid;
  const formId = findFormIdByTimerTrigger(triggerUid);
  if (!formId) {
    return;
  }
  const form = FormApp.openById(formId);
  form.setAcceptingResponses(false);
  cancelTimerInternal(formId, null, true);
  Logger.log('Timed session complete. Closed form "%s".', form.getTitle());
}

function deleteTriggerById(id) {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getUniqueId() === id) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function ensureSubmissionTrigger(formId, shouldExist) {
  const triggers = ScriptApp.getProjectTriggers();
  const matching = triggers.filter(function(trigger) {
    return trigger.getHandlerFunction() === 'handleSubmit' &&
      trigger.getTriggerSource() === ScriptApp.TriggerSource.FORMS &&
      trigger.getTriggerSourceId() === formId;
  });
  if (shouldExist) {
    if (!matching.length) {
      ScriptApp.newTrigger('handleSubmit')
        .forForm(FormApp.openById(formId))
        .onFormSubmit()
        .create();
    }
  } else {
    matching.forEach(function(trigger) {
      ScriptApp.deleteTrigger(trigger);
    });
  }
}

function ensureQuizConfiguration(form) {
  if (!form.isQuiz()) {
    form.setIsQuiz(true);
  }
  if (!form.collectsEmail()) {
    form.setCollectEmail(true);
  }
}

function handleSubmit(e) {
  const form = e && e.source ? e.source : null;
  const formId = form ? form.getId() : null;
  if (!formId) {
    return;
  }
  const settings = getSettings(formId);
  if (!settings.forceRetakeEnabled) {
    return;
  }
  const response = e && e.response;
  if (!response) {
    return;
  }
  const percent = calculatePercentScore(response);
  if (percent === null) {
    return;
  }
  const threshold = settings.scoreThreshold || 80;
  if (percent >= threshold) {
    return;
  }
  const respondentEmail = response.getRespondentEmail();
  const formInstance = form || FormApp.openById(formId);
  const retakeMessage = 'Your score was ' + percent.toFixed(2) + '% which is below the required ' + threshold + '%. Please retake the quiz using the link below.';
  if (respondentEmail) {
    MailApp.sendEmail({
      to: respondentEmail,
      subject: 'Please retake the quiz',
      htmlBody: retakeMessage + '<br><br><a href="' + formInstance.getPublishedUrl() + '">Retake quiz</a>'
    });
  }
  try {
    response.deleteResponse();
  } catch (err) {
    Logger.log('Unable to delete response: ' + err);
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function calculatePercentScore(response) {
  const totalScore = response.getScore();
  const gradableResponses = response.getGradableItemResponses();
  let earned = 0;
  let possible = 0;
  gradableResponses.forEach(function(itemResponse) {
    const possiblePoints = itemResponse.getPossiblePoints();
    if (typeof possiblePoints === 'number' && possiblePoints > 0) {
      possible += possiblePoints;
      const earnedPoints = itemResponse.getScore();
      if (typeof earnedPoints === 'number') {
        earned += earnedPoints;
      }
    }
  });
  if (possible > 0) {
    return (earned / possible) * 100;
  }
  if (typeof totalScore === 'number' && totalScore >= 0) {
    return totalScore;
  }
  return null;
}

function getActiveFormContext() {
  const form = FormApp.getActiveForm();
  if (!form) {
    throw new Error('Open a Google Form to use this add-on.');
  }
  return {
    form: form,
    formId: form.getId(),
    ui: FormApp.getUi()
  };
}

function cancelTimerInternal(formId, ui, silent) {
  const settings = getSettings(formId);
  if (settings.timerTriggerId) {
    deleteTriggerById(settings.timerTriggerId);
    settings.timerTriggerId = '';
    settings.sessionEndsAt = '';
    saveSettings(formId, settings);
    if (ui && !silent) {
      ui.alert('Active timer cancelled.');
    }
  } else if (ui && !silent) {
    ui.alert('No active timer to cancel.');
  }
}

function findFormIdByTimerTrigger(triggerUid) {
  if (!triggerUid) {
    return null;
  }
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  for (const key in all) {
    if (!Object.prototype.hasOwnProperty.call(all, key)) {
      continue;
    }
    if (key.indexOf(SETTINGS_KEY_PREFIX) !== 0) {
      continue;
    }
    try {
      const formId = key.substring(SETTINGS_KEY_PREFIX.length);
      const parsed = JSON.parse(all[key]);
      if (parsed.timerTriggerId === triggerUid) {
        return formId;
      }
    } catch (err) {
      continue;
    }
  }
  return null;
}
