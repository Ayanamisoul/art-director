"""Именованные подразделы в материалах презентации и лендинга."""


def read_subsections(parent):
    groups = parent.get("subsections", [])
    if not isinstance(groups, list):
        raise ValueError("В subsections нужен список подразделов.")
    for index, group in enumerate(groups, 1):
        if not isinstance(group, dict) or not isinstance(group.get("title"), str) or not group["title"].strip():
            raise ValueError(f"У подраздела {index} должен быть заголовок.")
        text = group.get("text", "")
        items = group.get("items", [])
        if not isinstance(text, str) or not isinstance(items, list) or any(not isinstance(i, str) for i in items):
            raise ValueError(f"В подразделе «{group['title']}» нужны текст и список строк в items.")
        if group.get("subsections"):
            raise ValueError(f"Вложите детали подраздела «{group['title']}» в items.")
        yield group["title"].strip(), text.strip(), [i.strip() for i in items if i.strip()]
