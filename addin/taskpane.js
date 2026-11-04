/* global Office, Word */
import { convertSelection, createAutoConverter, readActiveParagraph } from './src/word.js';
import { measureCaret } from './src/diagnostics.js';

const AUTO_SETTING_KEY = 'chem-vzorce.auto';
const MIN_EVENT_API_VERSION = '1.6';
const SHORTCUT_MAC = '⌘⇧⌥H';
const SHORTCUT_OTHER = 'Ctrl+Alt+Shift+H';

const readAutoSetting = () => {
  try {
    return localStorage.getItem(AUTO_SETTING_KEY) !== 'off';
  } catch {
    return true;
  }
};

const saveAutoSetting = (isEnabled) => {
  try {
    localStorage.setItem(AUTO_SETTING_KEY, isEnabled ? 'on' : 'off');
  } catch (error) {
    console.error('Nastavení se nepodařilo uložit', error);
  }
};

const setStatus = (message, kind = 'info') => {
  const status = document.getElementById('status');
  if (!status) return;
  status.textContent = message;
  status.dataset.kind = kind;
};

const reportError = (error) => {
  console.error('Úprava vzorců selhala', error);
  setStatus('Vzorce se nepodařilo upravit. Zkus to prosím znovu.', 'error');
};

const reportConverted = (count) => {
  setStatus(count > 0 ? `Upraveno vzorců: ${count}` : 'Nenašel jsem žádný vzorec k úpravě.');
};

const convertNow = async () => {
  try {
    await Word.run(async (context) => reportConverted(await convertSelection(context)));
  } catch (error) {
    reportError(error);
  }
};

const converter = createAutoConverter({
  runWord: (task) => Word.run(task),
  readActive: () => Word.run(readActiveParagraph),
  onError: reportError,
  onConverted: reportConverted,
});

let eventHandlers = [];

const startAutoConversion = async () => {
  await Word.run(async (context) => {
    const onChange = (event) => converter.handleEvent(event);
    eventHandlers = [
      context.document.onParagraphChanged.add(onChange),
      context.document.onParagraphAdded.add(onChange),
    ];
    await context.sync();
  });
};

const stopAutoConversion = async () => {
  converter.stop();
  if (eventHandlers.length === 0) return;
  const handlers = eventHandlers;
  eventHandlers = [];
  await Word.run(handlers[0].context, async (context) => {
    handlers.forEach((handler) => handler.remove());
    await context.sync();
  });
};

const setAutoConversion = async (isEnabled) => {
  try {
    await (isEnabled ? startAutoConversion() : stopAutoConversion());
    saveAutoSetting(isEnabled);
    setStatus(isEnabled ? 'Automatická úprava je zapnutá.' : 'Automatická úprava je vypnutá.');
  } catch (error) {
    console.error('Automatickou úpravu se nepodařilo přepnout', error);
    setStatus('Automatickou úpravu se nepodařilo přepnout.', 'error');
  }
};

/** Loads the add-in with each document, so automatic conversion needs no extra click. */
const loadWithDocuments = async () => {
  try {
    await Office.addin.setStartupBehavior(Office.StartupBehavior.load);
  } catch (error) {
    console.error('Automatické načítání doplňku není k dispozici', error);
  }
};

Office.onReady(async () => {
  Office.actions.associate('ConvertFormulas', async (event) => {
    await convertNow();
    event?.completed?.();
  });

  const isMac = Office.context.platform === Office.PlatformType.Mac;
  document.getElementById('shortcut-hint').textContent = isMac ? SHORTCUT_MAC : SHORTCUT_OTHER;
  document.getElementById('convert-button').addEventListener('click', convertNow);
  document.getElementById('diag-button').addEventListener('click', async () => {
    const output = document.getElementById('diag-output');
    output.textContent = 'Měřím…';
    try {
      output.textContent = await measureCaret();
    } catch (error) {
      output.textContent = `Měření selhalo: ${error?.code ?? ''} ${error?.message ?? error}`;
    }
  });

  const toggle = document.getElementById('auto-toggle');
  const canRunAuto = Office.context.requirements.isSetSupported('WordApi', MIN_EVENT_API_VERSION);
  if (!canRunAuto) {
    toggle.disabled = true;
    setStatus('Tato verze Wordu automatickou úpravu nepodporuje. Zkratka funguje dál.');
    return;
  }

  toggle.checked = readAutoSetting();
  toggle.addEventListener('change', () => setAutoConversion(toggle.checked));
  if (toggle.checked) await setAutoConversion(true);
  await loadWithDocuments();
});
