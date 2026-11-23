async function uploadFont(file) {
  const body = new FormData();
  body.append('file', file);
  const response = await fetch('/api/fonts', {method: 'POST', body});
  const font = await response.json().catch(() => ({}));
  if (!response.ok || !font.id) throw new Error(response.status === 404
    ? 'Для загрузки шрифта запустите обновлённую версию сервиса.'
    : font.detail || 'Не удалось загрузить шрифт.');
  return font;
}
