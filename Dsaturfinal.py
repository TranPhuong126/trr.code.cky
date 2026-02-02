import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from collections import defaultdict
import heapq
import os
from datetime import datetime, timedelta

# Cố gắng import để vẽ đồ thị
try:
    import networkx as nx
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    HAS_GRAPH = True
except Exception:
    HAS_GRAPH = False


class ExamSchedulerPro:
    def __init__(self, root):
        self.root = root
        self.root.title("Xếp Lịch Thi Thông Minh - DSatur Pro v2.1")
        # Khởi tạo kích thước an toàn (người dùng có thể resize)
        self.root.geometry("1200x800")

        # Dữ liệu
        self.data = None
        self.subjects = []
        self.student_subjects = defaultdict(set)
        self.subject_students = defaultdict(set)
        self.conflict_graph = defaultdict(set)
        self.schedule = {}
        self.schedule_by_day = {}  # Lưu lịch theo ngày
        self.max_exams_per_day = 2
        self.start_date = datetime.now()  # Ngày bắt đầu thi

        # Style hiện đại
        self.colors = {
            'bg': '#f0f2f5',
            'card': '#ffffff',
            'primary': '#4361ee',
            'success': '#4cc9f0',
            'warning': '#f72585',
            'danger': '#d90429',
            'dark': '#2b2d42',
            'light': '#edf2f4'
        }

        self.setup_styles()
        self.create_ui()

    def setup_styles(self):
        style = ttk.Style()
        try:
            style.theme_use('clam')
        except Exception:
            pass
        style.configure("TButton", padding=6, font=('Segoe UI', 10, 'bold'))
        style.configure("Treeview", background="white", fieldbackground="white", rowheight=26)
        style.map('Treeview', background=[('selected', self.colors['primary'])])

    def create_ui(self):
        # Header
        header = tk.Frame(self.root, bg=self.colors['primary'], height=60)
        header.pack(fill='x')
        header.pack_propagate(False)
        tk.Label(header, text="XẾP LỊCH THI THÔNG MINH - DSATUR PRO", font=('Segoe UI', 16, 'bold'),
                 fg='white', bg=self.colors['primary']).pack(pady=12)

        main = tk.PanedWindow(self.root, orient=tk.HORIZONTAL, sashrelief=tk.RAISED, bg=self.colors['bg'])
        main.pack(fill='both', expand=True, padx=10, pady=10)

        # === SIDEBAR TRÁI ===
        left = tk.Frame(main, bg=self.colors['card'], width=360, relief='flat')
        main.add(left)

        # Upload
        upload_frame = tk.LabelFrame(left, text="NHẬP DỮ LIỆU", bg=self.colors['card'], fg=self.colors['dark'], font=('Segoe UI', 11, 'bold'))
        upload_frame.pack(fill='x', padx=12, pady=8)
        tk.Button(upload_frame, text="CHỌN FILE EXCEL", command=self.load_file,
                  bg=self.colors['primary'], fg='white', font=('Segoe UI', 10, 'bold'),
                  relief='flat', padx=10, pady=8, cursor='hand2').pack(pady=8)
        self.file_label = tk.Label(upload_frame, text="Chưa chọn file...", bg=self.colors['card'], fg='gray', wraplength=320)
        self.file_label.pack(pady=4)

        # Cài đặt
        setting_frame = tk.LabelFrame(left, text="CÀI ĐẶT", bg=self.colors['card'], fg=self.colors['dark'], font=('Segoe UI', 11, 'bold'))
        setting_frame.pack(fill='x', padx=12, pady=8)

        # Số ca tối đa mỗi ngày
        tk.Label(setting_frame, text="Số ca tối đa mỗi ngày:", bg=self.colors['card'], font=('Segoe UI', 10)).pack(anchor='w', padx=8, pady=5)
        self.max_var = tk.IntVar(value=2)
        tk.Spinbox(setting_frame, from_=1, to=10, textvariable=self.max_var, width=6, font=('Segoe UI', 10)).pack(anchor='w', padx=8, pady=(0,6))

        # Ngày bắt đầu thi
        tk.Label(setting_frame, text="Ngày bắt đầu thi (dd/mm/yyyy):", bg=self.colors['card'], font=('Segoe UI', 10)).pack(anchor='w', padx=8, pady=5)
        date_frame = tk.Frame(setting_frame, bg=self.colors['card'])
        date_frame.pack(anchor='w', padx=8, pady=(0,6))

        self.day_var = tk.StringVar(value=str(datetime.now().day))
        self.month_var = tk.StringVar(value=str(datetime.now().month))
        self.year_var = tk.StringVar(value=str(datetime.now().year))

        tk.Spinbox(date_frame, from_=1, to=31, textvariable=self.day_var, width=4, font=('Segoe UI', 9)).pack(side='left', padx=2)
        tk.Label(date_frame, text="/", bg=self.colors['card']).pack(side='left')
        tk.Spinbox(date_frame, from_=1, to=12, textvariable=self.month_var, width=4, font=('Segoe UI', 9)).pack(side='left', padx=2)
        tk.Label(date_frame, text="/", bg=self.colors['card']).pack(side='left')
        tk.Spinbox(date_frame, from_=2024, to=2035, textvariable=self.year_var, width=6, font=('Segoe UI', 9)).pack(side='left', padx=2)

        # Nút chạy
        tk.Button(left, text="CHẠY DSATUR", command=self.run_dsatur,
                  bg=self.colors['success'], fg='white', font=('Segoe UI', 12, 'bold'),
                  relief='flat', padx=10, pady=10, cursor='hand2').pack(pady=18, padx=12, fill='x')

        # Thống kê
        stats_frame = tk.LabelFrame(left, text="THỐNG KÊ", bg=self.colors['card'], fg=self.colors['dark'], font=('Segoe UI', 11, 'bold'))
        stats_frame.pack(fill='both', expand=True, padx=12, pady=8)
        self.stats_text = tk.Text(stats_frame, height=12, bg=self.colors['light'], relief='flat', font=('Consolas', 10))
        self.stats_text.pack(fill='both', padx=8, pady=8)

        # === PHẦN PHẢI ===
        right = tk.Frame(main, bg=self.colors['card'])
        main.add(right)

        notebook = ttk.Notebook(right)
        notebook.pack(fill='both', expand=True, padx=8, pady=8)

        # Tab 1: Lịch thi theo ngày (MỚI)
        tab1 = tk.Frame(notebook, bg='white')
        notebook.add(tab1, text='📅 Lịch Thi Theo Ngày')
        self.tree_day = ttk.Treeview(tab1, columns=('Ngày', 'Ca', 'Môn', 'SV'), show='headings', height=20)
        self.tree_day.heading('Ngày', text='Ngày Thi')
        self.tree_day.heading('Ca', text='Ca')
        self.tree_day.heading('Môn', text='Môn Học')
        self.tree_day.heading('SV', text='Số SV')
        self.tree_day.column('Ngày', width=140, anchor='center')
        self.tree_day.column('Ca', width=80, anchor='center')
        self.tree_day.column('Môn', width=420)
        self.tree_day.column('SV', width=80, anchor='center')
        self.tree_day.pack(side='left', fill='both', expand=True, padx=8, pady=8)
        scroll_day = ttk.Scrollbar(tab1, orient='vertical', command=self.tree_day.yview)
        scroll_day.pack(side='right', fill='y')
        self.tree_day.configure(yscrollcommand=scroll_day.set)

        # Tab 2: Lịch thi theo ca
        tab2 = tk.Frame(notebook, bg='white')
        notebook.add(tab2, text='🎯 Lịch Thi Theo Ca')
        self.tree_schedule = ttk.Treeview(tab2, columns=('Ca', 'Môn', 'SV'), show='headings', height=20)
        self.tree_schedule.heading('Ca', text='Ca Thi')
        self.tree_schedule.heading('Môn', text='Môn Học')
        self.tree_schedule.heading('SV', text='Số SV')
        self.tree_schedule.column('Ca', width=80, anchor='center')
        self.tree_schedule.column('Môn', width=540)
        self.tree_schedule.column('SV', width=80, anchor='center')
        self.tree_schedule.pack(side='left', fill='both', expand=True, padx=8, pady=8)
        scroll_schedule = ttk.Scrollbar(tab2, orient='vertical', command=self.tree_schedule.yview)
        scroll_schedule.pack(side='right', fill='y')
        self.tree_schedule.configure(yscrollcommand=scroll_schedule.set)

        # Tab 3: Lịch SV
        tab3 = tk.Frame(notebook, bg='white')
        notebook.add(tab3, text='👨‍🎓 Lịch Sinh Viên')
        search_frame = tk.Frame(tab3, bg='white')
        search_frame.pack(fill='x', padx=8, pady=6)
        tk.Label(search_frame, text="Tìm:", bg='white', font=('Segoe UI', 10)).pack(side='left')
        self.search_var = tk.StringVar()
        tk.Entry(search_frame, textvariable=self.search_var, width=40, font=('Segoe UI', 10)).pack(side='left', padx=6)
        self.search_var.trace('w', self.filter_students)
        self.tree_student = ttk.Treeview(tab3, columns=('MSSV', 'Tên', 'Ngày', 'Ca', 'Môn'), show='headings')
        self.tree_student.heading('MSSV', text='MSSV')
        self.tree_student.heading('Tên', text='Họ Tên')
        self.tree_student.heading('Ngày', text='Ngày Thi')
        self.tree_student.heading('Ca', text='Ca')
        self.tree_student.heading('Môn', text='Môn Học')
        self.tree_student.column('MSSV', width=100)
        self.tree_student.column('Tên', width=200)
        self.tree_student.column('Ngày', width=120)
        self.tree_student.column('Ca', width=80)
        self.tree_student.column('Môn', width=340)
        scroll_student = ttk.Scrollbar(tab3, orient='vertical', command=self.tree_student.yview)
        self.tree_student.pack(side='left', fill='both', expand=True, padx=8, pady=8)
        scroll_student.pack(side='right', fill='y')
        self.tree_student.configure(yscrollcommand=scroll_student.set)

        # Tab 4: Đồ thị
        tab4 = tk.Frame(notebook, bg='white')
        notebook.add(tab4, text='📊 Đồ Thị Xung Đột')
        self.graph_canvas = tk.Canvas(tab4, bg='white')
        self.graph_canvas.pack(fill='both', expand=True, padx=8, pady=8)
        if not HAS_GRAPH:
            tk.Label(tab4, text="Cài networkx + matplotlib để xem đồ thị!", fg='red', font=('Segoe UI', 12)).pack(pady=50)

        # Tab 5: Export & Cảnh báo
        tab5 = tk.Frame(notebook, bg='white')
        notebook.add(tab5, text='💾 Export & Kiểm Tra')
        tk.Button(tab5, text="XUẤT LỊCH THI EXCEL", command=self.export_all,
                  bg=self.colors['warning'], fg='white', font=('Segoe UI', 11, 'bold'), pady=8).pack(pady=16)
        self.warning_text = tk.Text(tab5, height=12, bg='#fff5f5', fg='red', font=('Segoe UI', 10))
        self.warning_text.pack(fill='both', expand=True, padx=8, pady=8)

    def load_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not path:
            return

        try:
            all_dfs = []
            excel = pd.ExcelFile(path, engine='openpyxl')

            for sheet in excel.sheet_names:
                try:
                    # Đọc toàn bộ sheet như string
                    df = pd.read_excel(excel, sheet_name=sheet, header=None, dtype=str, engine='openpyxl')
                    df = df.fillna('')

                    # Tìm dòng header (chứa "Mã SV" hoặc "MSSV")
                    header_row = None
                    for idx in range(min(5, len(df))):  # Chỉ tìm trong 5 dòng đầu
                        row_text = ' '.join(df.iloc[idx].astype(str).str.lower().tolist())
                        if 'mã sv' in row_text or 'mssv' in row_text or 'ma sv' in row_text:
                            header_row = idx
                            break

                    if header_row is None:
                        # Nếu không tìm thấy header, thử đọc sheet như 1 cột danh sách MSSV
                        # Giả sử sheet là danh sách MSSV dưới header mặc định
                        df2 = pd.read_excel(excel, sheet_name=sheet, dtype=str, engine='openpyxl')
                        if df2.shape[1] >= 1:
                            col0 = df2.columns[0]
                            tmp = df2[[col0]].dropna()
                            tmp.columns = ['MaSV']
                            tmp['HoTen'] = 'N/A'
                            tmp['ChuongTrinh'] = sheet
                            all_dfs.append(tmp)
                            continue
                        else:
                            continue

                    # Lấy tên môn học (dòng đầu tiên hoặc tên sheet)
                    subject_name = sheet
                    if header_row > 0:
                        first_cell = str(df.iloc[0, 0]).strip()
                        if len(first_cell) > 0:
                            subject_name = first_cell

                    # Đặt header
                    df.columns = df.iloc[header_row]
                    df = df.iloc[header_row + 1:].reset_index(drop=True)

                    # Tìm cột Mã SV và Họ Tên
                    masv_col = None
                    hoten_col = None

                    for col in df.columns:
                        col_str = str(col).lower().strip()
                        if 'mã sv' in col_str or 'mssv' in col_str or 'ma sv' in col_str or 'mssv' in col_str:
                            masv_col = col
                        if 'họ' in col_str and 'tên' in col_str:
                            hoten_col = col
                        elif 'tên' in col_str and hoten_col is None:
                            hoten_col = col

                    if masv_col is None:
                        # Bỏ qua sheet nếu không có cột mã
                        continue

                    # Lọc dữ liệu
                    if hoten_col:
                        df_clean = df[[masv_col, hoten_col]].copy()
                        df_clean.columns = ['MaSV', 'HoTen']
                    else:
                        df_clean = df[[masv_col]].copy()
                        df_clean.columns = ['MaSV']
                        df_clean['HoTen'] = 'N/A'

                    df_clean['MaSV'] = df_clean['MaSV'].astype(str).str.strip()
                    df_clean = df_clean.loc[df_clean['MaSV'].str.len() > 0].copy()
                    mask_numeric = df_clean['MaSV'].str.match(r'^\d+$', na=False)
                    df_clean = df_clean.loc[mask_numeric].copy()

                    if len(df_clean) > 0:
                        df_clean['ChuongTrinh'] = subject_name
                        all_dfs.append(df_clean)

                except Exception as e:
                    print(f"Lỗi đọc sheet {sheet}: {e}")
                    continue

            if not all_dfs:
                messagebox.showerror("Lỗi", "Không tìm thấy dữ liệu hợp lệ!\n\nKiểm tra:\n- File có cột 'Mã SV'\n- Có ít nhất 1 sinh viên")
                return

            self.data = pd.concat(all_dfs, ignore_index=True)
            self.data.drop_duplicates(subset=['MaSV', 'ChuongTrinh'], inplace=True)

            self.file_label.config(
                text=f"ĐÃ TẢI: {os.path.basename(path)}\n{len(self.data)} dòng • {self.data['ChuongTrinh'].nunique()} môn",
                fg='green'
            )

            messagebox.showinfo("Thành công",
                                f"Đã tải thành công!\n\n• {len(self.data):,} bản ghi\n• {len(excel.sheet_names)} sheet\n• {self.data['MaSV'].nunique()} sinh viên\n• {self.data['ChuongTrinh'].nunique()} môn học")

            self.process_data()

        except Exception as e:
            messagebox.showerror("Lỗi đọc file", f"Chi tiết lỗi:\n{str(e)}")

    def process_data(self):
        self.subjects = sorted(self.data['ChuongTrinh'].unique().tolist())
        self.student_subjects.clear()
        self.subject_students.clear()
        self.conflict_graph.clear()

        for _, row in self.data.iterrows():
            sid = str(row['MaSV']).strip()
            subj = row['ChuongTrinh']
            self.student_subjects[sid].add(subj)
            self.subject_students[subj].add(sid)

        # Xây đồ thị xung đột
        for subs in self.student_subjects.values():
            subs = list(subs)
            for i in range(len(subs)):
                for j in range(i+1, len(subs)):
                    a = subs[i]
                    b = subs[j]
                    self.conflict_graph[a].add(b)
                    self.conflict_graph[b].add(a)

        self.update_stats()

    def update_stats(self):
        text = f"TỔNG QUAN DỮ LIỆU\n"
        text += f"{'='*40}\n"
        text += f"Sinh viên: {len(self.student_subjects):,}\n"
        text += f"Môn học: {len(self.subjects):,}\n"
        text += f"Xung đột cạnh: {sum(len(v) for v in self.conflict_graph.values())//2:,}\n"

        if self.schedule:
            total_slots = max(self.schedule.values())
            total_days = (total_slots + self.max_exams_per_day - 1) // self.max_exams_per_day
            text += f"\n{'='*40}\n"
            text += f"LỊCH THI\n"
            text += f"{'='*40}\n"
            text += f"Tổng ca thi: {total_slots}\n"
            text += f"Ca/ngày: {self.max_exams_per_day}\n"
            text += f"Tổng số ngày: {total_days}\n"

        self.stats_text.delete(1.0, 'end')
        self.stats_text.insert('end', text)

    def run_dsatur(self):
        if self.data is None or self.data.empty or len(self.subjects) == 0:
            messagebox.showwarning("Cảnh báo", "Chưa tải dữ liệu!")
            return

        # Lấy ngày bắt đầu
        try:
            self.start_date = datetime(int(self.year_var.get()), int(self.month_var.get()), int(self.day_var.get()))
        except Exception:
            messagebox.showerror("Lỗi", "Ngày tháng không hợp lệ!")
            return

        self.max_exams_per_day = int(self.max_var.get())
        self.schedule.clear()
        self.schedule_by_day.clear()

        # DSatur algorithm
        degree = {s: len(self.conflict_graph[s]) for s in self.subjects}
        saturation = {s: 0 for s in self.subjects}
        color_of = {}

        # Build initial heap: use (-saturation, -degree, subject) so we pop highest sat then highest degree
        heap = [(-saturation[s], -degree[s], s) for s in self.subjects]
        heapq.heapify(heap)
        colored = set()

        while heap:
            _, _, subj = heapq.heappop(heap)
            if subj in colored:
                continue
            # choose smallest color not used by neighbors
            used = {color_of.get(n) for n in self.conflict_graph[subj] if n in color_of}
            c = 1
            while c in used:
                c += 1
            color_of[subj] = c
            colored.add(subj)

            # update neighbors' saturation and push back
            for nei in self.conflict_graph[subj]:
                if nei not in colored:
                    # recompute saturation as number of distinct colors in neighbors
                    neigh_colors = {color_of.get(n) for n in self.conflict_graph[nei] if n in color_of}
                    saturation[nei] = len(neigh_colors)
                    heapq.heappush(heap, (-saturation[nei], -degree[nei], nei))

        self.schedule = color_of

        # Tính toán lịch theo ngày
        self.calculate_schedule_by_day()

        self.display_results()
        self.check_conflicts()
        self.draw_graph()

        total_days = (max(color_of.values()) + self.max_exams_per_day - 1) // self.max_exams_per_day
        messagebox.showinfo("HOÀN THÀNH",
                            f"Đã xếp lịch thành công!\n\n"
                            f"• Tổng ca thi: {max(color_of.values())}\n"
                            f"• Số ca/ngày: {self.max_exams_per_day}\n"
                            f"• Tổng số ngày thi: {total_days}")

    def calculate_schedule_by_day(self):
        """Tính toán lịch thi theo ngày dựa trên số ca tối đa mỗi ngày"""
        self.schedule_by_day.clear()

        for subject, slot in self.schedule.items():
            # Tính ngày thi (slot 1,2,3 = ngày 1, slot 4,5,6 = ngày 2,...)
            day_index = (slot - 1) // self.max_exams_per_day
            session_in_day = ((slot - 1) % self.max_exams_per_day) + 1

            exam_date = self.start_date + timedelta(days=day_index)
            date_str = exam_date.strftime("%d/%m/%Y")

            if date_str not in self.schedule_by_day:
                self.schedule_by_day[date_str] = {}

            # In trường hợp nhiều môn rơi vào cùng ca trong ngày (hiếm nếu slot mapping trùng), sắp xếp bằng slot
            self.schedule_by_day[date_str][session_in_day] = {
                'subject': subject,
                'students': len(self.subject_students[subject]),
                'slot': slot
            }

    def display_results(self):
        # Xóa dữ liệu cũ
        for tree in [self.tree_day, self.tree_schedule, self.tree_student]:
            for i in tree.get_children():
                tree.delete(i)

        # Tab 1: Lịch theo ngày
        for date in sorted(self.schedule_by_day.keys(), key=lambda x: datetime.strptime(x, "%d/%m/%Y")):
            sessions = self.schedule_by_day[date]
            for session in sorted(sessions.keys()):
                info = sessions[session]
                self.tree_day.insert('', 'end', values=(
                    date,
                    f'Ca {session}',
                    info['subject'],
                    info['students']
                ))

        # Tab 2: Lịch theo ca
        ca_dict = defaultdict(list)
        for subj, ca in self.schedule.items():
            ca_dict[ca].append((subj, len(self.subject_students[subj])))
        for ca in sorted(ca_dict):
            for subj, count in sorted(ca_dict[ca], key=lambda x: -x[1]):
                self.tree_schedule.insert('', 'end', values=(f'Ca {ca}', subj, count))

        # Tab 3: Lịch SV
        for sid, subs in self.student_subjects.items():
            name_df = self.data.loc[self.data['MaSV'] == sid, 'HoTen']
            name = name_df.iloc[0] if len(name_df) > 0 else "N/A"

            for sub in sorted(subs):
                slot = self.schedule.get(sub, 0)
                if slot == 0:
                    date_str = ""
                    session_in_day = ""
                else:
                    day_index = (slot - 1) // self.max_exams_per_day
                    session_in_day = ((slot - 1) % self.max_exams_per_day) + 1
                    exam_date = self.start_date + timedelta(days=day_index)
                    date_str = exam_date.strftime("%d/%m/%Y")

                self.tree_student.insert('', 'end', values=(sid, name, date_str, f'Ca {session_in_day}', sub))

        self.update_stats()

    def filter_students(self, *args):
        search = self.search_var.get().lower()
        for i in self.tree_student.get_children():
            self.tree_student.delete(i)

        for sid, subs in self.student_subjects.items():
            name_df = self.data.loc[self.data['MaSV'] == sid, 'HoTen']
            name = name_df.iloc[0] if len(name_df) > 0 else "N/A"

            if search in sid.lower() or search in name.lower():
                for sub in sorted(subs):
                    slot = self.schedule.get(sub, 0)
                    if slot == 0:
                        date_str = ""
                        session_in_day = ""
                    else:
                        day_index = (slot - 1) // self.max_exams_per_day
                        session_in_day = ((slot - 1) % self.max_exams_per_day) + 1
                        exam_date = self.start_date + timedelta(days=day_index)
                        date_str = exam_date.strftime("%d/%m/%Y")

                    self.tree_student.insert('', 'end', values=(sid, name, date_str, f'Ca {session_in_day}', sub))

    def check_conflicts(self):
        self.warning_text.delete(1.0, 'end')
        conflicts = []

        for sid, subs in self.student_subjects.items():
            cas = [self.schedule.get(s) for s in subs]
            if len([c for c in cas if c is not None]) != len(set([c for c in cas if c is not None])):
                name_df = self.data.loc[self.data['MaSV'] == sid, 'HoTen']
                name = name_df.iloc[0] if len(name_df) > 0 else "N/A"
                conflicts.append(f"TRÙNG: {sid} - {name}")

        if conflicts:
            self.warning_text.insert('1.0', "CÓ LỖI TRÙNG CA!\n" + "\n".join(conflicts[:200]))
            self.warning_text.config(fg='red')
        else:
            self.warning_text.insert(
                '1.0',
                "HOÀN HẢO! Không có sinh viên nào bị trùng ca thi\n\n"
                "✓ Tất cả sinh viên đều có lịch thi hợp lệ\n"
                "✓ Không có xung đột thời gian"
            )

    def draw_graph(self):
        # Vẽ đồ thị xung đột nếu thư viện có sẵn
        self.graph_canvas.delete('all')
        if not HAS_GRAPH:
            return

        try:
            G = nx.Graph()
            for subj in self.subjects:
                G.add_node(subj)
            for a, neighs in self.conflict_graph.items():
                for b in neighs:
                    if a != b:
                        G.add_edge(a, b)

            plt.clf()
            fig = plt.figure(figsize=(8, 6))
            ax = fig.add_subplot(111)
            ax.axis('off')

            # position
            pos = nx.spring_layout(G, seed=42)

            # color nodes by assigned slot (if any)
            node_colors = []
            max_slot = max(self.schedule.values()) if self.schedule else 1
            for n in G.nodes():
                slot = self.schedule.get(n, 0)
                node_colors.append(slot if slot > 0 else 0)

            nx.draw_networkx_nodes(G, pos, node_size=300, cmap=plt.cm.tab20, node_color=node_colors)
            nx.draw_networkx_edges(G, pos, alpha=0.4)
            nx.draw_networkx_labels(G, pos, font_size=8)

            # embed to tkinter
            canvas = FigureCanvasTkAgg(fig, master=self.graph_canvas)
            canvas.draw()
            canvas.get_tk_widget().pack(fill='both', expand=True)
        except Exception as e:
            print("Lỗi vẽ đồ thị:", e)

    def export_all(self):
        """Xuất file Excel: Lịch theo ngày (top-down), Lịch theo ca, Lịch sinh viên"""
        if not self.schedule:
            messagebox.showwarning("Chú ý", "Chưa có lịch để xuất. Vui lòng chạy DSatur trước.")
            return

        # đảm bảo schedule_by_day được cập nhật theo max_exams_per_day hiện tại
        self.calculate_schedule_by_day()

        # Chọn file lưu
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel files", "*.xlsx *.xls")],
                                            title="Lưu file lịch thi")
        if not path:
            return

        # --- Lịch theo ngày: sắp xếp ngày ↑, trong ngày ca ↑ ---
        day_rows = []
        sorted_dates = sorted(self.schedule_by_day.keys(),
                              key=lambda x: datetime.strptime(x, "%d/%m/%Y"))
        for date in sorted_dates:
            sessions = self.schedule_by_day[date]
            for session in sorted(sessions.keys()):
                info = sessions[session]
                day_rows.append({
                    'Ngày': date,
                    'Ca trong ngày': f'Ca {session}',
                    'Ca toàn bộ (DSatur)': info['slot'],
                    'Môn': info['subject'],
                    'Số SV': info['students']
                })
        df_day = pd.DataFrame(day_rows)

        # --- Lịch theo ca toàn bộ ---
        ca_rows = []
        ca_dict = defaultdict(list)
        for subj, ca in self.schedule.items():
            ca_dict[ca].append((subj, len(self.subject_students[subj])))
        for ca in sorted(ca_dict.keys()):
            for subj, count in sorted(ca_dict[ca], key=lambda x: -x[1]):
                ca_rows.append({
                    'Ca toàn bộ': f'Ca {ca}',
                    'Môn': subj,
                    'Số SV': count
                })
        df_ca = pd.DataFrame(ca_rows)

        # --- Lịch chi tiết theo sinh viên ---
        stu_rows = []
        for sid, subs in self.student_subjects.items():
            name_df = self.data.loc[self.data['MaSV'] == sid, 'HoTen']
            name = name_df.iloc[0] if len(name_df) > 0 else "N/A"
            for sub in sorted(subs):
                slot = self.schedule.get(sub, 0)
                if slot == 0:
                    date_str = ""
                    ca_in_day = ""
                    ca_full = ""
                else:
                    day_index = (slot - 1) // self.max_exams_per_day
                    session_in_day = ((slot - 1) % self.max_exams_per_day) + 1
                    exam_date = self.start_date + timedelta(days=day_index)
                    date_str = exam_date.strftime("%d/%m/%Y")
                    ca_in_day = f'Ca {session_in_day}'
                    ca_full = f'Ca {slot}'
                stu_rows.append({
                    'MSSV': sid,
                    'Họ Tên': name,
                    'Môn': sub,
                    'Ngày Thi': date_str,
                    'Ca trong ngày': ca_in_day,
                    'Ca toàn bộ': ca_full
                })
        df_stu = pd.DataFrame(stu_rows)

        # --- Sheet tóm tắt ---
        summary = {
            'Tổng sinh viên': [len(self.student_subjects)],
            'Tổng môn': [len(self.subjects)],
            'Tổng ca (toàn bộ)': [max(self.schedule.values())],
            'Số ca/ngày (cấu hình)': [self.max_exams_per_day],
            'Ngày bắt đầu': [self.start_date.strftime("%d/%m/%Y")]
        }
        df_sum = pd.DataFrame(summary)

        # Ghi ra Excel
        try:
            with pd.ExcelWriter(path, engine='openpyxl') as writer:
                df_day.to_excel(writer, sheet_name='Lich_Theo_Ngay', index=False)
                df_ca.to_excel(writer, sheet_name='Lich_Theo_Ca', index=False)
                df_stu.to_excel(writer, sheet_name='Lich_SinhVien', index=False)
                df_sum.to_excel(writer, sheet_name='ThongTin_TomTat', index=False)

            messagebox.showinfo("Xuất thành công",
                                f"Đã xuất lịch thi ra file:\n{os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("Lỗi xuất file",
                                 f"Không thể xuất file Excel.\nChi tiết: {str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExamSchedulerPro(root)
    root.mainloop()
