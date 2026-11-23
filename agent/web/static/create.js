const form = document.querySelector('#creator');
const fields = document.querySelector('#inputs');
const brief = document.querySelector('#brief');
const template = document.querySelector('#template');
const drop = document.querySelector('#drop');
const remove = document.querySelector('#remove');
const statusLine = document.querySelector('#status');
const result = document.querySelector('#result');
const viewer = document.querySelector('#viewer');
const slideImage = document.querySelector('#slide-image');
const previous = document.querySelector('#previous');
const next = document.querySelector('#next');
const previewNote = document.querySelector('#preview-note');
const retryPreview = document.querySelector('#retry-preview');
const useSample = document.querySelector('#use-sample');
let activeDeck = null;

function showSlide(index) {
  const deck = activeDeck;
  if (!deck || index < 1 || index > deck.pages) return;
  deck.index = index;
  const label = 'Слайд ' + index + ' из ' + deck.pages;
  document.querySelector('#position').textContent = label;
  previous.disabled = index === 1;
  next.disabled = index === deck.pages;
  slideImage.alt = label;
  slideImage.hidden = false;
  previewNote.textContent = '';
  retryPreview.hidden = true;
  slideImage.src = '/previews/' + encodeURIComponent(deck.name) + '/p' + index + '.png?v=' + encodeURIComponent(deck.version);
}

previous.onclick = () => showSlide(activeDeck.index - 1);
next.onclick = () => showSlide(activeDeck.index + 1);
viewer.addEventListener('keydown', event => {
  if (event.key === 'ArrowLeft' || event.key === 'ArrowRight') {
    event.preventDefault();
    showSlide(activeDeck.index + (event.key === 'ArrowLeft' ? -1 : 1));
  }
});
slideImage.onerror = () => {
  slideImage.hidden = true;
  previewNote.textContent = 'Не удалось загрузить изображение слайда. Файл PowerPoint доступен ниже.';
  retryPreview.hidden = false;
};

async function loadPreview(deck) {
  viewer.hidden = true;
  retryPreview.hidden = true;
  previewNote.textContent = 'Открываем слайды…';
  try {
    const response = await fetch('/api/presentation/' + encodeURIComponent(deck.name) + '/previews', {cache: 'no-store'});
    if (!response.ok) throw new Error();
    const info = await response.json();
    if (activeDeck !== deck) return;
    if (!info.ready || !info.count) throw new Error();
    deck.pages = info.count;
    deck.version = info.version;
    viewer.hidden = false;
    showSlide(Math.min(deck.index, deck.pages));
  } catch {
    if (activeDeck !== deck) return;
    previewNote.textContent = 'Предпросмотр пока недоступен. Можно скачать PowerPoint или открыть редактор.';
    retryPreview.hidden = false;
  }
}
retryPreview.onclick = () => loadPreview(activeDeck);
useSample.onclick = () => {
  brief.value = 'Объясни начинающим, как планировать неделю. Сделай 4 слайда с конкретным примером и коротким заданием в конце.';
  useSample.disabled = true;
  brief.setCustomValidity('');
  brief.focus();
};

function report(message, error = false) {
  statusLine.textContent = message;
  statusLine.classList.toggle('error', error);
}

function updateFile() {
  const file = template.files[0];
  document.querySelector('#filename').textContent = file ? file.name : 'Выберите или перетащите файл';
  remove.hidden = !file;
  const invalid = file && !/\.(pptx|pdf)$/i.test(file.name);
  template.setCustomValidity(invalid ? 'Выберите файл PowerPoint (.pptx) или PDF.' : '');
  report(template.validationMessage, !!invalid);
}

template.addEventListener('change', updateFile);
drop.addEventListener('dragover', () => drop.classList.add('over'));
for (const event of ['drop', 'dragleave']) drop.addEventListener(event, () => drop.classList.remove('over'));
remove.addEventListener('click', () => {
  template.value = '';
  updateFile();
  template.focus();
});
brief.addEventListener('input', () => {
  brief.setCustomValidity('');
  useSample.disabled = !!brief.value.trim();
});

const images = document.querySelector('#images');
const imageDrop = document.querySelector('#image-drop');
const imageList = document.querySelector('#attached-images');
const attachedImages = new Map();
const fontFile = document.querySelector('#font');
const removeFont = document.querySelector('#remove-font');

function updateFont() {
  const file = fontFile.files[0];
  document.querySelector('#font-name').textContent = file?.name || 'Выбрать свой шрифт';
  removeFont.hidden = !file;
  fontFile.setCustomValidity(file && (!/\.ttf$/i.test(file.name) || file.size > 16 * 1024 * 1024)
    ? 'Выберите статический файл TTF размером до 16 МБ.' : '');
  report(fontFile.validationMessage, !!fontFile.validationMessage);
}
fontFile.addEventListener('change', updateFont);
removeFont.onclick = () => {
  fontFile.value = '';
  updateFont();
  fontFile.focus();
};

function attachImage(file) {
  if (!/\.(png|jpe?g|webp|gif|bmp|svg)$/i.test(file.name)) {
    report('Файл «' + file.name + '» не добавлен. Выберите изображение.', true);
    return;
  }
  const item = document.createElement('li');
  const thumb = document.createElement('img');
  thumb.alt = '';
  thumb.src = URL.createObjectURL(file);
  const label = document.createElement('span');
  label.textContent = file.name;
  const button = document.createElement('button');
  button.type = 'button';
  button.textContent = '×';
  button.setAttribute('aria-label', 'Убрать ' + file.name);
  button.onclick = () => {
    attachedImages.delete(item);
    URL.revokeObjectURL(thumb.src);
    item.remove();
    images.focus();
  };
  item.append(thumb, label, button);
  imageList.append(item);
  attachedImages.set(item, {file, savedName: null});
}

async function uploadImages() {
  const names = [];
  for (const attachment of attachedImages.values()) {
    if (!attachment.savedName) {
      const body = new FormData();
      body.append('file', attachment.file);
      body.append('slide_attachment', 'true');
      const response = await fetch('/api/assets', {method: 'POST', body});
      const saved = await response.json();
      if (!response.ok || !saved.ok) throw new Error(saved.detail || 'Не удалось загрузить ' + attachment.file.name);
      attachment.savedName = saved.name;
    }
    names.push(attachment.savedName);
  }
  return names;
}
images.addEventListener('change', () => {
  report('');
  for (const file of images.files) attachImage(file);
  images.value = '';
});
imageDrop.addEventListener('dragover', () => imageDrop.classList.add('over'));
for (const event of ['drop', 'dragleave']) imageDrop.addEventListener(event, () => imageDrop.classList.remove('over'));

form.addEventListener('submit', async event => {
  event.preventDefault();
  if (fields.disabled) return;
  if (!brief.value.trim()) {
    brief.setCustomValidity('Введите тему или исходный текст.');
    brief.reportValidity();
    return;
  }
  const data = new FormData(form);
  if (!template.files.length) data.delete('template');
  fields.disabled = true;
  form.setAttribute('aria-busy', 'true');
  useSample.disabled = true;
  report('Готовим слайды и проверяем содержание. Это может занять несколько минут.');
  try {
    if (fontFile.files.length) {
      report('Проверяем шрифт…');
      const font = await uploadFont(fontFile.files[0]);
      data.set('font_id', font.id);
      report('Готовим слайды со шрифтом ' + font.family + '…');
    }
    if (attachedImages.size) {
      report('Загружаем изображения…');
      const names = await uploadImages();
      for (const name of names) data.append('image_names', name);
      data.set('brief', brief.value.trim() + '\n\nРазмести на слайдах прикреплённые изображения. '
        + 'Используй эти точные имена файлов в поле image:\n' + names.join('\n'));
      report('Готовим слайды с вашими изображениями…');
    }
    const response = await fetch(form.action, {method: 'POST', body: data});
    const answer = await response.json().catch(() => ({}));
    if (!response.ok || !answer.file) throw new Error(typeof answer.detail === 'string' ? answer.detail : 'Не удалось собрать презентацию.');
    const download = document.querySelector('#download');
    download.href = answer.file;
    download.download = (answer.title || 'Презентация') + '.pptx';
    document.querySelector('#result-title').textContent = answer.title;
    document.querySelector('#count').textContent = 'Слайдов: ' + answer.slides;
    document.querySelector('#edit').href = '/editor/' + encodeURIComponent(answer.name);
    document.querySelector('#chat').href = '/?session=' + encodeURIComponent(answer.session_id);
    const issues = document.querySelector('#issues');
    issues.replaceChildren();
    for (const issue of answer.validation?.issues || []) {
      const item = document.createElement('li');
      item.textContent = issue;
      issues.append(item);
    }
    issues.hidden = !issues.childElementCount;
    report(!issues.hidden ? 'Черновик готов. Проверьте замечания перед скачиванием.' :
      answer.validation?.raw ? 'Файл готов. Автоматическая проверка содержания недоступна.' : 'Презентация готова.');
    result.hidden = false;
    document.querySelector('#sample').hidden = true;
    activeDeck = {name: answer.name, pages: 0, index: 1, version: ''};
    loadPreview(activeDeck);
    document.querySelector('#result-title').focus({preventScroll: true});
    result.scrollIntoView({block: 'nearest'});
  } catch (error) {
    report(error instanceof TypeError ? 'Не удалось связаться с сайтом. Проверьте соединение и попробуйте снова.' : error.message, true);
  } finally {
    fields.disabled = false;
    useSample.disabled = !!brief.value.trim();
    form.removeAttribute('aria-busy');
  }
});
fields.disabled = false;
