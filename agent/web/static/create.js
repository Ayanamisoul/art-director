const form = document.querySelector('#creator');
const fields = document.querySelector('#inputs');
const brief = document.querySelector('#brief');
const template = document.querySelector('#template');
const drop = document.querySelector('#drop');
const remove = document.querySelector('#remove');
const statusLine = document.querySelector('#status');
const result = document.querySelector('#result');

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
brief.addEventListener('input', () => brief.setCustomValidity(''));

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
  result.hidden = true;
  report('Готовим слайды и проверяем содержание. Это может занять несколько минут.');
  try {
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
    download.focus();
  } catch (error) {
    report(error instanceof TypeError ? 'Не удалось связаться с сайтом. Проверьте соединение и попробуйте снова.' : error.message, true);
  } finally {
    fields.disabled = false;
    form.removeAttribute('aria-busy');
  }
});
