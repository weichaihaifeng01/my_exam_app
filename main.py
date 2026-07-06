# -*- coding: utf-8 -*-
import flet as ft
import json
import os
from typing import List, Dict, Optional, Tuple, Any
from datetime import datetime
import shutil

try:
    import openpyxl
    EXCEL_SUPPORT = True
except ImportError:
    EXCEL_SUPPORT = False

# ---------- 默认题库 ----------
DEFAULT_QUESTIONS = [
    {"question": "Python中用于定义函数的关键字是？", "options": ["func", "def", "function", "define"], "answer": ["B"], "type": "single"},
    {"question": "以下哪些是Python的内置数据类型？（多选）", "options": ["列表", "字典", "数组", "元组"], "answer": ["A", "B", "D"], "type": "multiple"},
    {"question": "2 + 3 * 4 的结果是？", "options": ["20", "14", "24", "9"], "answer": ["B"], "type": "single"},
    {"question": "以下哪些是有效的变量名？（多选）", "options": ["_var", "2var", "var_name", "var-name"], "answer": ["A", "C"], "type": "multiple"},
    {"question": "Python中如何获取列表长度？", "options": ["len()", "length()", "size()", "count()"], "answer": ["A"], "type": "single"}
]

WRONG_SET_FILE = "wrong_set.json"
QUESTIONS_FILE = "questions.json"
PROGRESS_FILE = "progress.json"
BACKUP_DIR = "backups"

# ---------- 数据管理 ----------
class DataManager:
    def __init__(self):
        self.questions: List[Dict] = DEFAULT_QUESTIONS.copy()
        self.wrong_set: Dict[int, int] = {}
        self._ensure_backup_dir()
        self.load_wrong_set()
        self.load_questions()

    def _ensure_backup_dir(self):
        if not os.path.exists(BACKUP_DIR):
            os.makedirs(BACKUP_DIR)

    def load_wrong_set(self):
        if os.path.exists(WRONG_SET_FILE):
            try:
                with open(WRONG_SET_FILE, 'r', encoding='utf-8') as f:
                    self.wrong_set = json.load(f)
                    self.wrong_set = {int(k): v for k, v in self.wrong_set.items()}
            except:
                self.wrong_set = {}

    def save_wrong_set(self):
        try:
            with open(WRONG_SET_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.wrong_set, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存错题集失败: {e}")

    def load_questions(self):
        if os.path.exists(QUESTIONS_FILE):
            try:
                with open(QUESTIONS_FILE, 'r', encoding='utf-8') as f:
                    self.questions = json.load(f)
            except:
                self.questions = DEFAULT_QUESTIONS.copy()

    def save_questions(self):
        try:
            if os.path.exists(QUESTIONS_FILE):
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                backup_path = os.path.join(BACKUP_DIR, f"questions_backup_{timestamp}.json")
                shutil.copy2(QUESTIONS_FILE, backup_path)
            with open(QUESTIONS_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.questions, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存题库失败: {e}")

    def _parse_answer_string(self, answer_str: str, valid_letters: List[str]) -> List[str]:
        s = answer_str.replace(' ', '').replace('\u3000', '')
        if not s:
            return []
        if any(sep in s for sep in ('、', ',', '，', ';', '；')):
            parts = s.replace(',', '、').replace('，', '、').replace(';', '、').replace('；', '、').split('、')
            letters = [p.strip().upper() for p in parts if p.strip()]
        else:
            letters = [c.upper() for c in s if c.isalpha()]
        return [l for l in letters if l in valid_letters]

    def _normalize_answer_to_letters(self, answer_data: Any, options: List[str], q_index: int) -> Tuple[Optional[List[str]], Optional[str]]:
        valid_letters = [chr(ord('A') + i) for i in range(len(options))]
        option_letter_map = {opt.strip(): valid_letters[i] for i, opt in enumerate(options)}
        if isinstance(answer_data, str):
            letters = self._parse_answer_string(answer_data, valid_letters)
            if not letters:
                return None, f"答案字符串 '{answer_data}' 无法解析"
            return sorted(list(set(letters))), None
        elif isinstance(answer_data, list):
            letters = []
            for item in answer_data:
                item = str(item).strip()
                if not item:
                    continue
                if len(item) == 1 and item.isalpha() and item.upper() == item:
                    letter = item.upper()
                    if letter in valid_letters:
                        letters.append(letter)
                    else:
                        return None, f"答案字母 '{letter}' 超出范围"
                else:
                    if item in option_letter_map:
                        letters.append(option_letter_map[item])
                    else:
                        return None, f"答案 '{item}' 不匹配任何选项"
            if not letters:
                return None, "答案为空"
            return sorted(list(set(letters))), None
        else:
            return None, "答案格式错误"

    def import_from_json(self, filepath: str) -> Tuple[bool, str]:
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                new_questions = json.load(f)
            if not isinstance(new_questions, list):
                raise ValueError("JSON 根元素必须是数组")
            cleaned = []
            for i, q in enumerate(new_questions):
                if not isinstance(q, dict):
                    continue
                required = ("question", "options", "answer", "type")
                if not all(k in q for k in required):
                    raise ValueError(f"第 {i+1} 题缺少必要字段")
                if q['type'] not in ("single", "multiple"):
                    raise ValueError(f"第 {i+1} 题 type 必须为 'single' 或 'multiple'")
                options = [opt.strip() for opt in q['options'] if opt and str(opt).strip()]
                if len(options) < 2:
                    raise ValueError(f"第 {i+1} 题选项至少2个")
                answer_letters, err = self._normalize_answer_to_letters(q['answer'], options, i+1)
                if err:
                    raise ValueError(f"第 {i+1} 题答案错误：{err}")
                cleaned.append({
                    "question": q['question'].strip(),
                    "options": options,
                    "answer": answer_letters,
                    "type": q['type']
                })
            if not cleaned:
                raise ValueError("未解析到有效题目")
            self.questions = cleaned
            self.save_questions()
            self.wrong_set.clear()
            self.save_wrong_set()
            return True, f"✅ JSON 导入成功，共 {len(cleaned)} 题"
        except Exception as e:
            return False, f"❌ 导入失败：{str(e)}"

    def import_from_excel(self, filepath: str) -> Tuple[bool, str]:
        if not EXCEL_SUPPORT:
            return False, "❌ 缺少 openpyxl，请执行 pip install openpyxl"
        try:
            wb = openpyxl.load_workbook(filepath)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
            if len(rows) < 2:
                return False, "❌ Excel 至少需要标题行和一行数据"
            headers = [str(h).strip() if h else "" for h in rows[0]]
            question_col = answer_col = type_col = None
            option_cols = {}
            for i, h in enumerate(headers):
                if h == "题目":
                    question_col = i
                elif h == "答案":
                    answer_col = i
                elif h == "类型":
                    type_col = i
                elif h.startswith("选项") and len(h) > 2:
                    letter = h[2:]
                    if letter.isalpha() and len(letter) == 1:
                        option_cols[letter.upper()] = i
            missing = []
            if question_col is None: missing.append("题目")
            if answer_col is None: missing.append("答案")
            if type_col is None: missing.append("类型")
            if not option_cols: missing.append("至少一个选项列（如“选项A”）")
            if missing:
                return False, f"❌ Excel 缺少列：{', '.join(missing)}"
            sorted_letters = sorted(option_cols.keys())
            option_indices = [option_cols[l] for l in sorted_letters]
            new_questions = []
            error_msgs = []
            for row_idx, row in enumerate(rows[1:], start=2):
                if not row or all(c is None for c in row):
                    continue
                try:
                    question = str(row[question_col]).strip()
                    options = []
                    for idx in option_indices:
                        val = row[idx]
                        if val is not None:
                            opt = str(val).strip()
                            if opt:
                                options.append(opt)
                    if len(options) < 2:
                        error_msgs.append(f"第 {row_idx} 行：选项不足")
                        continue
                    answer_raw = str(row[answer_col]).strip()
                    valid_letters = [chr(ord('A')+i) for i in range(len(options))]
                    answer_letters = self._parse_answer_string(answer_raw, valid_letters)
                    if not answer_letters:
                        error_msgs.append(f"第 {row_idx} 行：答案格式错误")
                        continue
                    qtype_raw = str(row[type_col]).strip()
                    if qtype_raw in ("单选", "single"):
                        qtype = "single"
                    elif qtype_raw in ("多选", "multiple"):
                        qtype = "multiple"
                    else:
                        error_msgs.append(f"第 {row_idx} 行：类型必须为“单选”或“多选”")
                        continue
                    new_questions.append({
                        "question": question,
                        "options": options,
                        "answer": sorted(list(set(answer_letters))),
                        "type": qtype
                    })
                except Exception as e:
                    error_msgs.append(f"第 {row_idx} 行：解析出错 - {str(e)}")
            if not new_questions:
                err_detail = "\n".join(error_msgs) if error_msgs else "未读取到有效题目"
                return False, f"❌ Excel 导入失败：{err_detail}"
            self.questions = new_questions
            self.save_questions()
            self.wrong_set.clear()
            self.save_wrong_set()
            msg = f"✅ Excel 导入成功，共 {len(new_questions)} 题"
            if error_msgs:
                msg += f"\n⚠️ 以下行被跳过：\n" + "\n".join(error_msgs)
            return True, msg
        except Exception as e:
            return False, f"❌ Excel 读取失败：{str(e)}"

    def update_wrong_set(self, index: int, is_correct: bool):
        if is_correct:
            if index in self.wrong_set:
                self.wrong_set[index] += 1
                if self.wrong_set[index] >= 2:
                    del self.wrong_set[index]
        else:
            self.wrong_set[index] = 0
        self.save_wrong_set()

    def get_statistics(self) -> Dict[str, Any]:
        return {'total': len(self.questions), 'wrong': len(self.wrong_set), 'wrong_indices': list(self.wrong_set.keys())}


# ---------- Flet 主应用 ----------
def main(page: ft.Page):
    # 页面设置 —— 移动端竖屏适配
    page.title = "📚 考题练习软件"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.padding = 15
    page.scroll = ft.ScrollMode.AUTO
    page.window_width = 400   # 桌面预览竖屏宽度
    page.window_height = 700  # 桌面预览竖屏高度
    page.adaptive = True

    data = DataManager()

    # ---------- 应用状态 ----------
    class AppState:
        def __init__(self):
            self.question_order = list(range(len(data.questions)))
            self.order_pos = 0
            self.question_states = {}
            self.current_index = None
            self.submitted = False
            self.score = 0
            self.review_mode = False
            self.review_queue = []
            self.review_total = 0
            self.review_correct = 0
            self.selected_letters = set()
    state = AppState()

    # ---------- UI 控件 ----------
    title_text = ft.Text("📚 考题练习软件", size=24, weight=ft.FontWeight.BOLD)
    progress_text = ft.Text("", size=16)
    score_text = ft.Text("🏆 得分: 0", size=16, color=ft.Colors.GREEN)
    top_row = ft.Row(
        [title_text, progress_text, score_text],
        alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
        vertical_alignment=ft.CrossAxisAlignment.CENTER
    )

    type_label = ft.Text("", size=14, weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE)
    question_label = ft.Text("", size=19, weight=ft.FontWeight.W_500, selectable=True)
    result_label = ft.Text("", size=17, weight=ft.FontWeight.BOLD)
    options_column = ft.Column(spacing=12)

    submit_btn = ft.ElevatedButton("✅ 提交答案", on_click=None, disabled=True, height=50, style=ft.ButtonStyle(text_style=ft.TextStyle(size=16)))
    prev_btn = ft.OutlinedButton("⏪ 上一题", on_click=None, disabled=True, height=50, style=ft.ButtonStyle(text_style=ft.TextStyle(size=14)))
    next_btn = ft.OutlinedButton("⏭️ 下一题", on_click=None, disabled=True, height=50, style=ft.ButtonStyle(text_style=ft.TextStyle(size=14)))
    import_btn = ft.OutlinedButton("📥 导入题库", on_click=None, height=45, style=ft.ButtonStyle(text_style=ft.TextStyle(size=13)))
    view_wrong_btn = ft.OutlinedButton("📋 查看错题", on_click=None, height=45, style=ft.ButtonStyle(text_style=ft.TextStyle(size=13)))
    review_btn = ft.OutlinedButton("🔄 复习错题", on_click=None, height=45, style=ft.ButtonStyle(text_style=ft.TextStyle(size=13)))
    exit_review_btn = ft.OutlinedButton("🚪 退出复习", on_click=None, disabled=True, height=45, style=ft.ButtonStyle(text_style=ft.TextStyle(size=13)))
    clear_wrong_btn = ft.OutlinedButton("🗑️ 清空错题", on_click=None, height=45, style=ft.ButtonStyle(text_style=ft.TextStyle(size=13)))

    button_row1 = ft.Row([prev_btn, submit_btn, next_btn], spacing=10, alignment=ft.MainAxisAlignment.CENTER)
    button_row2 = ft.Row(
        [import_btn, view_wrong_btn, review_btn, exit_review_btn, clear_wrong_btn],
        spacing=8,
        wrap=True,
        alignment=ft.MainAxisAlignment.CENTER
    )

    file_picker = ft.FilePicker()
    page.overlay.append(file_picker)

    # ---------- 辅助函数 ----------
    def close_dialog():
        if page.dialog:
            page.dialog.open = False
            page.update()

    def show_snack(msg: str):
        page.snack_bar = ft.SnackBar(ft.Text(msg))
        page.snack_bar.open = True
        page.update()

    def reset_all():
        state.review_mode = False
        state.score = 0
        state.order_pos = 0
        state.question_states = {}
        state.review_queue = []
        state.review_total = 0
        state.review_correct = 0
        state.question_order = list(range(len(data.questions)))
        score_text.value = "🏆 得分: 0"
        review_btn.disabled = False
        exit_review_btn.disabled = True
        title_text.value = "📚 考题练习软件"
        save_progress()
        load_question()

    def update_progress():
        total = len(state.question_order)
        if total == 0:
            progress_text.value = "📭 无题目"
            return
        pos = state.order_pos + 1
        if state.review_mode:
            progress_text.value = f"🔄 [复习] {pos}/{total}"
        else:
            progress_text.value = f"📖 {pos}/{total}"
        page.update()

    def save_progress():
        progress = {
            'mode': 'review' if state.review_mode else 'normal',
            'score': state.score,
            'order_pos': state.order_pos,
            'question_states': state.question_states,
            'review_queue': state.review_queue if state.review_mode else [],
            'review_total': state.review_total,
            'review_correct': state.review_correct
        }
        try:
            with open(PROGRESS_FILE, 'w', encoding='utf-8') as f:
                json.dump(progress, f, ensure_ascii=False, indent=2)
        except:
            pass

    def load_progress():
        if os.path.exists(PROGRESS_FILE):
            try:
                with open(PROGRESS_FILE, 'r', encoding='utf-8') as f:
                    prog = json.load(f)
                state.review_mode = prog.get('mode') == 'review'
                state.score = prog.get('score', 0)
                state.order_pos = prog.get('order_pos', 0)
                state.question_states = {int(k): v for k, v in prog.get('question_states', {}).items()}
                if state.review_mode:
                    state.review_queue = prog.get('review_queue', [])
                    state.review_total = prog.get('review_total', 0)
                    state.review_correct = prog.get('review_correct', 0)
                    review_btn.disabled = True
                    exit_review_btn.disabled = False
                    title_text.value = "📚 考题练习软件 [🔄 复习模式]"
                    state.question_order = sorted(data.wrong_set.keys())
                else:
                    state.question_order = list(range(len(data.questions)))
                    review_btn.disabled = False
                    exit_review_btn.disabled = True
                score_text.value = f"🏆 得分: {state.score}"
                if state.order_pos >= len(state.question_order) and len(state.question_order) > 0:
                    state.order_pos = len(state.question_order) - 1
                update_progress()
            except:
                pass

    def show_result(s: dict):
        q = data.questions[state.current_index]
        correct_letters = set(q.get('answer', []))
        user_letters = set(s.get('selected', []))
        is_correct = (user_letters == correct_letters)
        if is_correct:
            result_label.value = "✅ 回答正确！"
            result_label.color = ft.Colors.GREEN
        else:
            correct_display = ''.join(sorted(correct_letters))
            result_label.value = f"❌ 错误，正确答案：{correct_display}"
            result_label.color = ft.Colors.RED

        for idx, chk in enumerate(options_column.controls):
            letter = chr(ord('A') + idx)
            opt_text = q['options'][idx]
            is_corr = letter in correct_letters
            is_wrong = letter in user_letters and letter not in correct_letters
            if is_corr:
                chk.label = f"✅ {letter}. {opt_text}"
                chk.value = True
                chk.fill_color = ft.Colors.GREEN
            elif is_wrong:
                chk.label = f"❌ {letter}. {opt_text}"
                chk.value = True
                chk.fill_color = ft.Colors.RED
            else:
                chk.label = f"○ {letter}. {opt_text}"
                chk.value = False
            chk.disabled = True
        page.update()

    def on_option_change(chk: ft.Checkbox):
        if state.submitted:
            return
        q = data.questions[state.current_index]
        qtype = q.get('type', 'single')
        if qtype == 'single':
            for c in options_column.controls:
                if c != chk:
                    c.value = False
            chk.value = not chk.value
        else:
            chk.value = not chk.value
        state.selected_letters = {c.data for c in options_column.controls if c.value}
        page.update()

    # ---------- 核心功能 ----------
    def load_question():
        if not state.question_order:
            question_label.value = "暂无题目，请导入题库"
            options_column.controls.clear()
            submit_btn.disabled = True
            next_btn.disabled = True
            prev_btn.disabled = True
            result_label.value = ""
            page.update()
            return

        state.current_index = state.question_order[state.order_pos]
        q = data.questions[state.current_index]
        state.submitted = False
        state.selected_letters.clear()

        qtype = q.get('type', 'single')
        type_label.value = "🔵 【单选题】" if qtype == 'single' else "🟣 【多选题】"
        question_label.value = q.get('question', '')
        result_label.value = ""
        result_label.color = None

        options_column.controls.clear()
        letters = [chr(ord('A') + i) for i in range(len(q.get('options', [])))]
        for i, opt in enumerate(q.get('options', [])):
            letter = letters[i]
            chk = ft.Checkbox(
                label=f"○ {letter}. {opt}",
                value=False,
                data=letter,
                on_change=on_option_change,
                fill_color=ft.Colors.BLUE_200,
                label_style=ft.TextStyle(size=16),
                adaptive=True
            )
            options_column.controls.append(chk)

        if state.current_index in state.question_states:
            s = state.question_states[state.current_index]
            if s.get('submitted', False):
                state.submitted = True
                show_result(s)
                submit_btn.disabled = True
                next_btn.disabled = False
                page.update()
                return

        submit_btn.disabled = False
        next_btn.disabled = True
        for chk in options_column.controls:
            chk.disabled = False
            chk.fill_color = ft.Colors.BLUE_200

        prev_btn.disabled = (state.order_pos == 0)
        update_progress()
        page.update()

    def submit_answer(e):
        if state.submitted:
            return
        if not state.selected_letters:
            show_snack("请至少选择一个选项！")
            return

        q = data.questions[state.current_index]
        correct_letters = set(q.get('answer', []))
        user_letters = set(state.selected_letters)
        is_correct = (user_letters == correct_letters)

        if is_correct:
            state.score += 1
            score_text.value = f"🏆 得分: {state.score}"

        state.question_states[state.current_index] = {
            'submitted': True,
            'selected': list(user_letters),
            'correct': is_correct
        }
        data.update_wrong_set(state.current_index, is_correct)

        state.submitted = True
        show_result(state.question_states[state.current_index])
        submit_btn.disabled = True
        next_btn.disabled = False
        save_progress()
        page.update()

    def next_question(e):
        if not state.submitted:
            show_snack("请先提交答案")
            return
        if state.order_pos + 1 < len(state.question_order):
            state.order_pos += 1
            load_question()
        else:
            if state.review_mode:
                exit_review(e)
            else:
                def reset_dlg(e):
                    page.dialog.open = False
                    reset_all()
                    page.update()
                def close_dlg(e):
                    page.dialog.open = False
                    page.update()
                dlg = ft.AlertDialog(
                    title=ft.Text("🎉 全部完成！"),
                    content=ft.Text("所有题目已答完，是否重新开始？"),
                    actions=[
                        ft.TextButton("是", on_click=reset_dlg),
                        ft.TextButton("否", on_click=close_dlg)
                    ]
                )
                page.dialog = dlg
                dlg.open = True
                page.update()

    def prev_question(e):
        if state.order_pos > 0:
            state.order_pos -= 1
            load_question()

    def start_review(e):
        if not data.wrong_set:
            show_snack("🎉 当前没有错题！")
            return
        state.review_mode = True
        state.review_queue = sorted(data.wrong_set.keys())
        state.review_total = len(state.review_queue)
        state.review_correct = 0
        review_btn.disabled = True
        exit_review_btn.disabled = False
        title_text.value = "📚 考题练习软件 [🔄 复习模式]"
        state.question_order = state.review_queue
        state.order_pos = 0
        load_question()
        save_progress()

    def exit_review(e):
        if not state.review_mode:
            return
        if state.review_total > 0:
            msg = f"📊 复习结束！\n✅ 答对：{state.review_correct} 题\n❌ 答错：{state.review_total - state.review_correct} 题"
            show_snack(msg)
        state.review_mode = False
        state.review_queue = []
        state.review_total = 0
        state.review_correct = 0
        review_btn.disabled = False
        exit_review_btn.disabled = True
        title_text.value = "📚 考题练习软件"
        state.question_order = list(range(len(data.questions)))
        state.order_pos = 0
        load_question()
        save_progress()

    def clear_wrong_set(e):
        if not data.wrong_set:
            show_snack("错题集已为空")
            return
        def confirm_clear(e):
            page.dialog.open = False
            data.wrong_set.clear()
            data.save_wrong_set()
            if state.review_mode:
                exit_review(e)
            state.question_order = list(range(len(data.questions)))
            state.order_pos = 0
            load_question()
            save_progress()
            show_snack("错题集已清空")
            page.update()
        def cancel_clear(e):
            page.dialog.open = False
            page.update()
        dlg = ft.AlertDialog(
            title=ft.Text("🗑️ 确认清空"),
            content=ft.Text("确定要清空所有错题记录吗？"),
            actions=[
                ft.TextButton("确定", on_click=confirm_clear),
                ft.TextButton("取消", on_click=cancel_clear)
            ]
        )
        page.dialog = dlg
        dlg.open = True
        page.update()

    def view_wrong_list(e):
        wrong_indices = list(data.wrong_set.keys())
        if not wrong_indices:
            show_snack("🎉 当前没有错题")
            return
        content_col = ft.Column(spacing=8, scroll=ft.ScrollMode.AUTO)
        for idx in wrong_indices:
            q = data.questions[idx]
            correct_letters = set(q.get('answer', []))
            correct_display = ''.join(sorted(correct_letters))
            content_col.controls.append(ft.Text(f"【第 {idx+1} 题】{q['question']}", weight=ft.FontWeight.BOLD))
            for j, opt in enumerate(q['options']):
                letter = chr(ord('A') + j)
                mark = "✅" if letter in correct_letters else "○"
                content_col.controls.append(ft.Text(f"  {mark} {letter}. {opt}"))
            content_col.controls.append(ft.Text(f"🎯 正确答案：{correct_display}"))
            content_col.controls.append(ft.Divider(height=2))
        dlg = ft.AlertDialog(
            title=ft.Text(f"📋 错题列表（共 {len(wrong_indices)} 题）"),
            content=ft.Container(content_col, width=400, height=400),
            actions=[ft.TextButton("关闭", on_click=lambda _: close_dialog())]
        )
        page.dialog = dlg
        dlg.open = True
        page.update()

    # ---------- 导入功能 ----------
    def import_questions(e):
        file_picker.pick_files(
            allow_multiple=False,
            file_type=ft.FilePickerFileType.ALL,
            on_result=on_file_picked
        )

    def on_file_picked(result: ft.FilePickerResultEvent):
        if not result or not result.files:
            return
        filepath = result.files[0].path
        if not filepath:
            return
        if filepath.endswith('.json'):
            success, msg = data.import_from_json(filepath)
        elif filepath.endswith('.xlsx'):
            success, msg = data.import_from_excel(filepath)
        else:
            show_snack("不支持的文件格式，请选择 .xlsx 或 .json")
            return
        if success:
            show_snack(msg)
            reset_all()
        else:
            show_snack(msg)

    # ---------- 绑定回调 ----------
    submit_btn.on_click = submit_answer
    prev_btn.on_click = prev_question
    next_btn.on_click = next_question
    import_btn.on_click = import_questions
    view_wrong_btn.on_click = view_wrong_list
    review_btn.on_click = start_review
    exit_review_btn.on_click = exit_review
    clear_wrong_btn.on_click = clear_wrong_set

    # ---------- 组装界面 ----------
    page.add(
        top_row,
        ft.Divider(height=8),
        type_label,
        question_label,
        ft.Divider(height=8),
        result_label,
        options_column,
        ft.Divider(height=8),
        button_row1,
        ft.Divider(height=5),
        button_row2,
    )

    load_progress()
    load_question()

    def on_close(e):
        save_progress()
    page.on_close = on_close


if __name__ == "__main__":
    ft.app(target=main)