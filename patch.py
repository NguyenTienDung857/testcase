import re

with open('fit_images_column_k.py', 'r', encoding='utf-8') as f:
    text = f.read()

text = re.sub(
    r'@dataclass\nclass DependencyStatus:.*?missing: list\[str\]\n',
    '',
    text,
    flags=re.DOTALL
)

print("Check deps removal:", re.subn(r'def check_dependencies\(\).*?installed\.append.*?\n', '', text, flags=re.DOTALL)[1])
text = re.sub(r'def check_dependencies\(\) -> DependencyStatus:.*?return DependencyStatus.*?\n\n', '', text, flags=re.DOTALL)

print("on_check_deps removal:", re.subn(r'    def on_check_dependencies\(self\) -> None:.*?Da du thu vien.*?\n\n', '', text, flags=re.DOTALL)[1])
text = re.sub(r'    def on_check_dependencies\(self\) -> None:.*?Da du thu vien.*?\n\n', '', text, flags=re.DOTALL)

print("btn instantiation removal:", re.subn(r'        self\.check_deps_button = self\._build_button\(.*?\)\s*self\.check_deps_button\.grid.*?\n\n', '', text, flags=re.DOTALL)[1])
text = re.sub(r'        self\.check_deps_button = self\._build_button\(.*?\)\s*self\.check_deps_button\.grid.*?\n\n', '', text, flags=re.DOTALL)

text = re.sub(r'\s*self\.check_deps_button\.configure\(state=state\)\n', '\n', text)

text = re.sub(r'EXCEL IMAGE FITTER', 'NEXUS GRID', text)
text = re.sub(r'Quan ly va xu ly anh trong Excel', 'Cyberpunk Operation Branch', text)

text = re.sub(
    r'        nav_items = \[\n\s*"Huong dan nhanh",\n\s*"1\) Them mot hoac nhieu file Excel",\n\s*"2\) Chon cot can xu ly cho tung file",\n\s*"3\) Dung Run cho tung dong hoac RUN ALL",\n\s*\]',
    '''        nav_items = [
            "DARK MARKETPLACE",
            "» Data Sync Node",
            "» Neural Link Auth",
            "» Encrypted Vault",
            "» Satellite Uplink",
        ]''',
    text,
    flags=re.DOTALL
)

text = re.sub(r'BANG DIEU KHIEN XU LY EXCEL', 'COMMAND ZONE // PREMIUM', text)
text = re.sub(r'Chon file, dat cot cho tung file, theo doi log va chay xu ly nhanh\.', 'Visual overhaul // Million-dollar app experience active.', text)

nav_end_idx = text.find('        main = tk.Frame(shell, bg=UI_COLORS["bg"])')
if nav_end_idx != -1:
    memo_code = '''        theme_card = self._create_glass_card(sidebar, padx=12, pady=12)
        theme_card.pack(fill="x", padx=10, pady=(0, 10))
        tk.Label(
            theme_card,
            text="Theme memo:\\nOld one was red,\\nnow emerald green",
            bg=UI_COLORS["panel"],
            fg=UI_COLORS["accent"],
            font=("Segoe UI", 9, "italic"),
            anchor="w",
            justify="left",
        ).grid(row=0, column=0, sticky="w")
        
'''
    text = text[:nav_end_idx] + memo_code + text[nav_end_idx:]

with open('fit_images_column_k.py', 'w', encoding='utf-8') as f:
    f.write(text)
