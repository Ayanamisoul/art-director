"""Блок валидации: проверяет сгенерированный HTML кодом, а не доверием к модели.

Принцип «код vs промпт»: всё, что можно проверить детерминированно,
проверяем здесь. Модель получает список замечаний и может исправить результат.
"""
import re

# Эти домены не считаем «внешними ресурсами» (нейтральные пространства имён)
ALLOWED_URL_PREFIXES = ("http://www.w3.org", "https://www.w3.org")


def _visible_text(html: str) -> str:
    """Текст страницы без тегов, стилей и скриптов — для проверок языка."""
    without_blocks = re.sub(r"<(style|script|template)\b[\s\S]*?</\1>", " ", html, flags=re.IGNORECASE)
    return re.sub(r"<[^>]+>", " ", without_blocks)


def validate_html(html: str) -> list[str]:
    """Возвращает список замечаний. Пустой список = всё хорошо."""
    issues: list[str] = []
    text = html.strip()

    if not text:
        return ["пустой документ"]
    if "```" in text:
        issues.append("код обёрнут в markdown (```) — нужен чистый HTML")
    if not re.search(r"<html[\s>]", text, re.IGNORECASE):
        issues.append("нет тега <html> — документ неполный")
    if "</html>" not in text.lower():
        issues.append("нет закрывающего </html> — документ обрезан")

    urls = re.findall(r"https?://[^\s\"'<>)]+", text)
    external = [u for u in urls if not u.startswith(ALLOWED_URL_PREFIXES)]
    if external:
        issues.append("внешние ресурсы запрещены: " + ", ".join(external[:3]))

    if not re.search(r"<h1[\s>]", text, re.IGNORECASE):
        issues.append("нет заголовка <h1> — на странице должен быть главный заголовок")

    if not re.search(r"(контакт|тел\.|телефон|mailto:|@)", text, re.IGNORECASE):
        issues.append("нет блока контактов")

    # Смешение кириллицы и латиницы в одном слове видимого текста («маккиацchio»)
    for word in set(re.findall(r"[A-Za-zА-Яа-яЁё]{2,}", _visible_text(text))):
        if re.search(r"[А-Яа-яЁё]", word) and re.search(r"[A-Za-z]", word):
            issues.append(f"смешение кириллицы и латиницы в слове «{word}»")

    return issues
