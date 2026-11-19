let sessionId = null;
let presentation = null;
const messagesEl = document.getElementById('messages');
const filesEl = document.getElementById('files');
const form = document.getElementById('form');
const input = document.getElementById('input');
const sendBtn = document.getElementById('send');

function addMsg(role, text, tools, validation, saved) {
  const div = document.createElement('div');
  div.className = 'msg ' + role;
  div.textContent = text;
  if (validation && validation.issues && validation.issues.length) {
    const v = document.createElement('div');
    v.className = 'tools';
    v.textContent = (validation.ok ? 'проверено судьёй' : 'судья: ' + validation.issues.join('; '));
    div.appendChild(v);
  }
  if (tools && tools.length) {
    const det = document.createElement('details');
    det.className = 'tools';
    det.innerHTML = '<summary> вызовы инструментов: ' + tools.length + '</summary>';
    const pre = document.createElement('pre');
    pre.textContent = tools.join('\n');
    det.appendChild(pre);
    div.appendChild(det);
  }
  messagesEl.appendChild(div);
  messagesEl.scrollTop = messagesEl.scrollHeight;
  if (saved && saved.presentation) renderSlidesPreview(div, saved.presentation);
  if (saved && saved.html) renderHtmlPreview(div, saved.html);
  return div;
}

function esc(s) { const d = document.createElement('div'); d.textContent = s || ''; return d.innerHTML; }

async function renderSlidesPreview(container, name) {
  const r = await fetch('/api/presentation/' + encodeURIComponent(name));
  if (!r.ok) return;
  const data = await r.json();
  const accent = '#' + ((data.theme && data.theme.accent) || '2E5E4E');
  const wrap = document.createElement('div');
  wrap.className = 'pv-wrap';
  // настоящие рендеры слайдов (если готовы) вместо CSS-схемы — они показывают
  // реальный дизайн шаблона, как в PowerPoint
  const pv = await fetch('/api/presentation/' + encodeURIComponent(name) + '/previews')
    .then(x => x.json()).catch(() => ({ready: false}));
  if (pv.ready) {
    for (let i = 1; i <= pv.count; i++) {
      const img = document.createElement('img');
      img.className = 'pv-shot';
      img.loading = 'lazy';
      img.src = '/previews/' + encodeURIComponent(name) + '/p' + i + '.png';
      wrap.appendChild(img);
    }
    const src = document.createElement('span');
    src.className = 'pv-src';
    src.textContent = 'рендер из файла .pptx';
    wrap.appendChild(src);
  } else {
    data.slides.forEach(s => {
      const d = document.createElement('div');
      if (s.kind === 'title' || s.kind === 'closing') {
        d.className = 'pv-slide pv-title-slide';
        d.style.background = '#1A2E28';
        d.innerHTML = `<h2>${esc(s.title)}</h2><p>${esc(s.subtitle || '')}</p>`;
      } else if (s.kind === 'section') {
        d.className = 'pv-slide pv-section';
        d.style.background = accent;
        d.innerHTML = `<h2>${esc(s.title)}</h2>`;
      } else {
        const dense = (s.bullets || []).length > 4;
        const text = `<h3 style="color:${accent}">${esc(s.title)}</h3>
          <ul>${(s.bullets || []).map(b => `<li>${esc(b)}</li>`).join('')}</ul>`;
        d.className = 'pv-slide pv-content' + (dense ? ' dense' : '');
        d.innerHTML = `<div class="bar" style="background:${accent}"></div>` + (s.image
          ? `<div class="pv-flex"><div class="txt">${text}</div><img src="/assets/${encodeURIComponent(s.image)}"></div>`
          : text);
      }
      wrap.appendChild(d);
    });
  }
  container.appendChild(wrap);
  const btn = document.createElement('a');
  btn.className = 'big-edit-btn';
  btn.href = '/editor/' + encodeURIComponent(name);
  btn.textContent = 'Открыть в редакторе';
  container.appendChild(btn);
  const pdf = document.createElement('a');
  pdf.className = 'pdf-btn';
  pdf.href = '#';
  pdf.textContent = 'Скачать PDF';
  pdf.onclick = async (e) => {
    e.preventDefault();
    pdf.textContent = 'Готовлю PDF…';
    try {
      const r = await fetch('/api/presentation/' + encodeURIComponent(name) + '/pdf', { method: 'POST' });
      const res = await r.json();
      if (res.ok) { window.open('/files/' + encodeURIComponent(name) + '.pdf', '_blank'); pdf.textContent = 'Скачать PDF'; }
      else { pdf.textContent = ' ' + (res.detail || 'ошибка'); }
    } catch (err) { pdf.textContent = ' ' + err; }
  };
  container.appendChild(pdf);
}

function renderHtmlPreview(container, name) {
  const f = document.createElement('iframe');
  f.className = 'landing-pv';
  f.src = '/files/' + encodeURIComponent(name) + '.html';
  container.appendChild(f);
}

async function refreshFiles() {
  const r = await fetch('/api/files');
  const data = await r.json();
  filesEl.innerHTML = '';
  data.files.forEach(name => {
    const a = document.createElement('a');
    a.className = 'file-link';
    a.href = '/files/' + encodeURIComponent(name);
    a.textContent = name;
    a.target = '_blank';
    filesEl.appendChild(a);
    if (name.endsWith('.pptx')) {
      const base = name.replace(/\.pptx$/, '');
      const edit = document.createElement('a');
      edit.className = 'edit-btn';
      edit.href = '/editor/' + encodeURIComponent(base);
      edit.textContent = 'Редактировать';
      filesEl.appendChild(edit);
      const chat = document.createElement('a');
      chat.className = 'chat-btn';
      chat.textContent = 'чат';
      chat.onclick = () => goChat(base);
      filesEl.appendChild(chat);
    }
  });
}

function setPresentation(name) {
  presentation = name;
  const bar = document.getElementById('focusbar');
  if (!name) { bar.className = ''; bar.innerHTML = ''; return; }
  bar.className = 'on';
  bar.innerHTML = ' Правим презентацию: <b>' + esc(name) + '</b>' +
    '<a href="/editor/' + encodeURIComponent(name) + '">открыть в редакторе</a>';
  input.placeholder = 'Например: поменяй заголовок второго слайда на «…» или добавь слайд про цены';
}

async function goChat(name) {
  const r = await fetch('/api/session/for-presentation', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({name}),
  });
  const data = await r.json();
  location.href = '/?session=' + encodeURIComponent(data.id);
}

// --- восстановление чата по ссылке ?session=<id> (кнопка « чат» в панели) ---
const restoreId = new URLSearchParams(location.search).get('session');
if (restoreId) restoreSession(restoreId);

async function restoreSession(id) {
  try {
    const r = await fetch('/api/session/' + encodeURIComponent(id));
    if (!r.ok) return;
    const data = await r.json();
    sessionId = data.id;
    setPresentation(data.presentation);
    for (const m of data.messages) {
      if (m.role === 'system' || m.role === 'tool') continue;
      if (m.role === 'user') addMsg('user', m.content);
      else if (m.role === 'assistant' && m.content) addMsg('agent', m.content);
    }
    // предпросмотр последнего сохранённого результата этого чата
    for (let i = data.messages.length - 1; i >= 0; i--) {
      const m = data.messages[i];
      if (m.role !== 'assistant' || !m.tool_calls) continue;
      let presName = null, htmlName = null;
      for (const call of m.tool_calls) {
        try {
          const args = JSON.parse(call.function.arguments || '{}');
          if (call.function.name === 'save_presentation') presName = args.name;
          if (call.function.name === 'save_html') htmlName = args.name;
        } catch (_) {}
      }
      if (presName || htmlName) {
        // сервер сохраняет под слагом (пробелы/знаки -> _), зеркало _safe_name
        const slug = s => String(s).replace(/[^\p{L}\p{N}_-]/gu, '_');
        const anchor = addMsg('agent', '— последний результат этого чата:');
        if (presName) renderSlidesPreview(anchor, slug(presName));
        if (htmlName) renderHtmlPreview(anchor, slug(htmlName));
        break;
      }
    }
  } catch (err) {
    console.error('не удалось восстановить сессию:', err);
  }
}

form.addEventListener('submit', async (e) => {
  e.preventDefault();
  const text = input.value.trim();
  if (!text) return;
  input.value = '';
  sendBtn.disabled = true;
  addMsg('user', text);
  const thinking = addMsg('agent', 'Думаю…');
  try {
    const r = await fetch('/api/chat', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({message: text, session_id: sessionId,
        template: document.getElementById('template').value,
        template_images: document.getElementById('no-template-photos').checked ? 'remove' : '',
        output_format: document.getElementById('outformat').value}),
    });
    const data = await r.json();
    sessionId = data.session_id;
    if (data.presentation) setPresentation(data.presentation);
    thinking.remove();
    refreshTemplates();
    if (data.error) addMsg('agent', 'Ошибка: ' + data.error, data.tools);
    else addMsg('agent', data.answer, data.tools, data.validation, data.saved);
    refreshFiles();
  } catch (err) {
    thinking.textContent = 'Ошибка соединения: ' + err;
  }
  sendBtn.disabled = false;
  input.focus();
});

refreshFiles();

// ---------- выбор шаблона презентаций ----------
async function refreshTemplates() {
  const [list, pref] = await Promise.all([
    fetch('/api/templates').then(r => r.json()),
    fetch('/api/templates/preferred').then(r => r.json()).catch(() => ({preferred: ''})),
  ]);
  const sel = document.getElementById('template');
  const current = sel.value || pref.preferred;
  sel.innerHTML = '<option value="">Авто — самый свежий загруженный</option>' +
    list.files.map(f => '<option value="' + esc(f) + '">' + esc(f) + '</option>').join('');
  sel.value = current;
}
document.getElementById('template').addEventListener('change', async e => {
  await fetch('/api/templates/preferred', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({name: e.target.value}),
  });
});
refreshTemplates();

// ---------- загрузка файлов: кнопки, drag&drop, вставка из буфера ----------
document.getElementById('up-img').addEventListener('change', e => uploadInput(e));
document.getElementById('up-tpl').addEventListener('change', e => uploadInput(e));

async function uploadInput(e) {
  const files = Array.from(e.target.files || []);
  e.target.value = '';
  for (const file of files) await uploadOne(file);
}

function kindOf(file) {
  const ext = (file.name.split('.').pop() || '').toLowerCase();
  if (ext === 'pptx') return { url: '/api/templates', label: 'шаблон' };
  if (['png', 'jpg', 'jpeg', 'webp', 'gif', 'svg', 'bmp'].includes(ext)) return { url: '/api/assets', label: 'изображение' };
  return null;
}

async function uploadOne(file) {
  const kind = kindOf(file);
  if (!kind) {
    note(`Файл «${file.name}» не принят: поддерживаются картинки (png/jpg/webp/gif/svg) и шаблоны .pptx`, true);
    return;
  }
  const pill = note(` Загружаю «${file.name}» как ${kind.label}…`);
  const fd = new FormData();
  fd.append('file', file);
  try {
    const r = await fetch(kind.url, { method: 'POST', body: fd });
    const res = await r.json();
    if (res.ok) {
      pill.textContent = ` «${res.name}» — ${kind.label} загружен. Агент видит файл через ${kind.url === '/api/assets' ? 'list_images' : 'list_templates'}.`;
    } else {
      pill.textContent = ` «${file.name}» не загружен: ${res.detail || 'ошибка сервера'}`;
      pill.classList.add('err');
    }
  } catch (err) {
    pill.textContent = ` «${file.name}» не загружен: ${err}`;
    pill.classList.add('err');
  }
  refreshFiles();
  refreshTemplates();
}

function note(text, isError) {
  const div = document.createElement('div');
  div.className = 'msg agent upload' + (isError ? ' err' : '');
  div.textContent = text;
  messagesEl.appendChild(div);
  messagesEl.scrollTop = messagesEl.scrollHeight;
  return div;
}

// --- перетаскивание файлов в любое место окна ---
const dropzone = document.getElementById('dropzone');
let dragDepth = 0;
const hasFiles = e => Array.from((e.dataTransfer && e.dataTransfer.types) || []).includes('Files');

window.addEventListener('dragenter', e => {
  if (!hasFiles(e)) return;
  e.preventDefault();
  dragDepth++;
  dropzone.classList.add('on');
});
window.addEventListener('dragover', e => {
  if (!hasFiles(e)) return;
  e.preventDefault();
  e.dataTransfer.dropEffect = 'copy';
});
window.addEventListener('dragleave', e => {
  if (!hasFiles(e)) return;
  if (--dragDepth <= 0) { dragDepth = 0; dropzone.classList.remove('on'); }
});
window.addEventListener('drop', async e => {
  if (!hasFiles(e)) return;
  e.preventDefault();
  dragDepth = 0;
  dropzone.classList.remove('on');
  for (const file of Array.from(e.dataTransfer.files)) await uploadOne(file);
});

// --- вставка скриншота из буфера (Ctrl+V) ---
document.addEventListener('paste', async e => {
  const items = Array.from((e.clipboardData && e.clipboardData.items) || []);
  const images = items.filter(i => i.type.startsWith('image/'));
  if (!images.length) return;
  e.preventDefault();
  for (const [index, item] of images.entries()) {
    const blob = item.getAsFile();
    if (!blob) continue;
    const ext = (blob.type.split('/')[1] || 'png').replace('jpeg', 'jpg');
    const stamp = new Date().toISOString().slice(11, 19).replace(/:/g, '-');
    await uploadOne(new File([blob], `вставка-${stamp}${index ? '-' + index : ''}.${ext}`, { type: blob.type }));
  }
});

// ---------- список чатов: открыть / удалить ----------
async function refreshChats() {
  const r = await fetch('/api/sessions');
  const data = await r.json();
  const box = document.getElementById('chats');
  box.innerHTML = '';
  data.sessions.forEach(s => {
    const row = document.createElement('div');
    row.className = 'chat-item';
    const a = document.createElement('a');
    a.className = 'title';
    a.href = '/?session=' + encodeURIComponent(s.id);
    a.textContent = s.presentation || (s.preview || 'чат без сообщений');
    const del = document.createElement('button');
    del.className = 'del';
    del.title = 'Удалить чат';
    del.textContent = '×';
    del.onclick = async () => {
      if (!confirm('Удалить этот чат? История пропадёт (файлы презентаций останутся).')) return;
      await fetch('/api/session/' + encodeURIComponent(s.id), { method: 'DELETE' });
      if (sessionId === s.id) location.href = '/';
      else refreshChats();
    };
    row.appendChild(a);
    row.appendChild(del);
    box.appendChild(row);
  });
  if (!data.sessions.length) box.innerHTML = '<div class="hint">Пока пусто — чаты появятся после первого сообщения.</div>';
}
refreshChats();


// ---------- картинки пользователя: список + удаление ----------
async function refreshAssets() {
  const r = await fetch('/api/assets');
  const data = await r.json();
  const box = document.getElementById('assets');
  box.innerHTML = '';
  data.files.forEach(name => {
    const row = document.createElement('div');
    row.className = 'asset-item';
    const img = document.createElement('img');
    img.src = '/assets/' + encodeURIComponent(name);
    const span = document.createElement('span');
    span.textContent = name;
    const del = document.createElement('button');
    del.className = 'del';
    del.title = 'Удалить картинку';
    del.textContent = '×';
    del.onclick = async () => {
      if (!confirm('Удалить картинку «' + name + '»? Из слайдов она не пропадёт — только из активов.')) return;
      await fetch('/api/assets/' + encodeURIComponent(name), { method: 'DELETE' });
      refreshAssets();
    };
    row.appendChild(img); row.appendChild(span); row.appendChild(del);
    box.appendChild(row);
  });
  if (!data.files.length) box.innerHTML = '<div class="hint">Картинок нет.</div>';
}
refreshAssets();
