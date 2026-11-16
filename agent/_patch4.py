"""Одноразовый патч: rich-фичи в _emit_slides. Запуск: py _patch4.py"""
import pathlib

p = pathlib.Path("pptx_builder.py")
src = p.read_text(encoding="utf-8")

# 1) stats var в content-ветке _emit_slides
old = '''            bullets = slide_data.get("bullets") or []
            image_name = slide_data.get("image")'''
new = '''            bullets = slide_data.get("bullets") or []
            stats = slide_data.get("stats") or []
            image_name = slide_data.get("image")'''
assert src.count(old) == 1, "stats var"
src = src.replace(old, new, 1)

# 2) custom-ветка: буллеты с подпунктами + рамка/линейка/статы
loop_marker = "                for i, bullet in enumerate(bullets):\n"
first = src.index(loop_marker)                      # use_template branch
second = src.index(loop_marker, first + 10)         # custom branch
i0 = second
i1 = src.index('run.font.color.rgb = _rgb(t["text"])', i0) \
    + len('run.font.color.rgb = _rgb(t["text"])\n')
new_block = '''                if stats:
                    body.height = int(SLIDE_H * 0.5)
                _fill_bullets_box(body, bullets, 22 if len(bullets) <= 4 else 18)
'''
src = src[:i0] + new_block + src[i1:]

# 3) рамка/линейка/статы: рамку вокруг body и линейку под заголовком
old_body = '''                text_width = Inches(6.4) if has_image else SLIDE_W - Inches(1.8)
                body = slide.shapes.add_textbox(Inches(0.9), Inches(1.7),
                                                text_width, SLIDE_H - Inches(2.2))
                frame = body.text_frame
                frame.word_wrap = True
'''
new_body = '''                text_width = Inches(6.4) if has_image else SLIDE_W - Inches(1.8)
                body = slide.shapes.add_textbox(Inches(0.9), Inches(1.7),
                                                text_width, SLIDE_H - Inches(2.2))
                frame = body.text_frame
                frame.word_wrap = True
                if sd.get("frame"):
                    _add_frame(slide, prs,
                               Inches(0.3), Inches(0.35), SLIDE_W - Inches(0.6), SLIDE_H - Inches(0.7),
                               t["accent"])
                if stats:
                    body.height = Inches(2.9)
'''
assert src.count(old_body) == 1, "body block"
src = src.replace(old_body, new_body, 1)

old_title = '''                add_text(slide, Inches(0.7), Inches(0.5), SLIDE_W - Inches(1.4), Inches(1),
                         title, 32, t["accent"], bold=True)
                text_width = Inches(6.4)'''
new_title = '''                title_box = add_text(slide, Inches(0.7), Inches(0.5), SLIDE_W - Inches(1.4), Inches(1),
                         title, 32, t["accent"], bold=True)
                if sd.get("underline", True):
                    _add_title_underline(slide, title_box, t["accent"])
                text_width = Inches(6.4)'''
assert src.count(old_title) == 1, "title block"
src = src.replace(old_title, new_title, 1)

p.write_text(src, encoding="utf-8")
print("patch4 done")
