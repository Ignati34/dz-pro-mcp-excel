const $ = (selector) => document.querySelector(selector);
const messages = $('#messages');
const input = $('#message-input');
const sendButton = $('#send-button');
const statusPill = $('#service-status');
let sessionId = crypto.randomUUID();

function element(tag, className, value) {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (value !== undefined) node.textContent = value;
  return node;
}

function addTrace(row, traces) {
  if (!traces?.length) return;
  const details = element('details', 'trace');
  details.append(element('summary', '', `Вызовы MCP · ${traces.length}`));
  const list = element('div', 'trace-list');
  for (const trace of traces) {
    const item = element('div', 'trace-item');
    const head = element('div', 'trace-head');
    head.append(element('code', '', trace.name || 'неизвестный инструмент'));
    head.append(element('span', `trace-state ${trace.status === 'success' ? '' : 'error'}`, trace.status === 'success' ? 'Успех' : 'Ошибка'));
    item.append(head, element('div', 'trace-label', 'Параметры'), element('pre', '', JSON.stringify(trace.arguments ?? {}, null, 2)));
    item.append(element('div', 'trace-label', 'Результат / ошибка'));
    item.append(element('pre', '', typeof trace.result === 'string' ? trace.result : JSON.stringify(trace.result, null, 2)));
    list.append(item);
  }
  details.append(list);
  row.append(details);
}

function addMessage(role, text, traces = [], pending = null) {
  const row = element('article', `message-row ${role}`);
  row.append(element('div', 'bubble', text), element('div', 'message-meta', role === 'user' ? 'Вы' : 'Помощник'));
  addTrace(row, traces);
  if (pending) {
    const actions = element('div', 'pending-actions');
    const confirmButton = element('button', 'button button-confirm', 'Подтвердить изменение');
    const cancelButton = element('button', 'button button-light', 'Отмена');
    confirmButton.addEventListener('click', () => finishAction(pending.action_id, true, confirmButton, cancelButton));
    cancelButton.addEventListener('click', () => finishAction(pending.action_id, false, confirmButton, cancelButton));
    actions.append(confirmButton, cancelButton);
    row.append(actions);
  }
  messages.append(row);
  messages.scrollTop = messages.scrollHeight;
}

async function refreshRequests() {
  const state = $('#requests-state');
  try {
    const response = await fetch('/api/requests');
    const data = await response.json();
    $('#requests-body').replaceChildren();
    if (data.error) {
      state.textContent = `Ошибка Excel MCP: ${data.error}`;
      state.className = 'request-state error';
      return;
    }
    state.textContent = `${data.requests.length} заявок · получено через read_range`;
    state.className = 'request-state';
    for (const request of data.requests) {
      const tr = document.createElement('tr');
      [request.id, request.client, request.problem, request.status].forEach((value) => tr.append(element('td', '', String(value ?? ''))));
      $('#requests-body').append(tr);
    }
  } catch {
    state.textContent = 'Не удалось получить данные Excel MCP.';
    state.className = 'request-state error';
  }
}

async function checkHealth() {
  try {
    const response = await fetch('/api/health');
    const data = await response.json();
    const online = data.mcp === 'available' && data.compatible;
    statusPill.className = `status ${online ? 'ok' : 'bad'}`;
    statusPill.lastElementChild.textContent = online ? `Excel MCP доступен · ${data.tool_count} инструментов` : `Excel MCP недоступен${data.error ? `: ${data.error}` : ''}`;
  } catch {
    statusPill.className = 'status bad';
    statusPill.lastElementChild.textContent = 'Нет связи с приложением';
  }
}

async function finishAction(actionId, approve, confirmButton, cancelButton) {
  confirmButton.disabled = cancelButton.disabled = true;
  try {
    const response = await fetch(`/api/requests/${approve ? 'confirm' : 'cancel'}`, {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ session_id: sessionId, action_id: actionId }),
    });
    const data = await response.json();
    const message = !approve ? data.message : data.ok
      ? data.already_completed ? 'Изменение уже было выполнено ранее; повторной записи не было.' : `Статус заявки №${data.request.id} изменён на «${data.request.status}», проверено повторным чтением MCP.`
      : data.message || 'Изменение не подтверждено.';
    addMessage('assistant', message, data.tool_calls || []);
    confirmButton.closest('.pending-actions').remove();
    await refreshRequests();
  } catch {
    addMessage('assistant', 'Не удалось завершить действие. Обновите таблицу и проверьте состояние заявки.');
  }
}

async function send(text) {
  const message = text.trim();
  if (!message || sendButton.disabled) return;
  addMessage('user', message);
  input.value = '';
  input.style.height = 'auto';
  sendButton.disabled = true;
  sendButton.textContent = '…';
  const waiting = element('article', 'message-row assistant');
  waiting.append(element('div', 'bubble', 'Обращаюсь к OpenAI и Excel MCP…'));
  messages.append(waiting);
  try {
    const response = await fetch('/api/chat', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ session_id: sessionId, message }),
    });
    const data = await response.json();
    waiting.remove();
    addMessage('assistant', data.reply || 'Пустой ответ приложения.', data.tool_calls || [], data.pending || null);
    if ((data.tool_calls || []).some((trace) => trace.status === 'error')) {
      statusPill.className = 'status bad';
      statusPill.lastElementChild.textContent = 'Ошибка инструмента MCP';
    }
    if ((data.tool_calls || []).some((trace) => trace.status === 'success' && trace.name === 'write_range')) await refreshRequests();
  } catch {
    waiting.remove();
    addMessage('assistant', 'Не удалось подключиться к приложению. Проверьте, что FastAPI запущен.');
  } finally {
    sendButton.disabled = false;
    sendButton.textContent = '↑';
    input.focus();
    checkHealth();
  }
}

$('#chat-form').addEventListener('submit', (event) => { event.preventDefault(); send(input.value); });
input.addEventListener('input', () => { input.style.height = 'auto'; input.style.height = `${Math.min(input.scrollHeight, 120)}px`; });
input.addEventListener('keydown', (event) => { if (event.key === 'Enter' && !event.shiftKey) { event.preventDefault(); $('#chat-form').requestSubmit(); } });
document.querySelectorAll('.suggestion').forEach((button) => button.addEventListener('click', () => send(button.textContent)));
$('#refresh-requests').addEventListener('click', refreshRequests);
$('#new-chat').addEventListener('click', async () => {
  const oldId = sessionId;
  sessionId = crypto.randomUUID();
  await fetch('/api/new-chat', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ session_id: oldId }) });
  messages.replaceChildren(element('div', 'welcome', 'Новый диалог. Напишите команду для работы с заявками.'));
  input.focus();
});

checkHealth();
refreshRequests();
