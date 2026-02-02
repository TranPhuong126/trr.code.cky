"""
frontend.py - Giao diện Tkinter
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import os

# Import backend
from backend import ExamSchedulerBackend

# Cố gắng import để vẽ đồ thị
try:
    import networkx as nx
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    HAS_GRAPH = True
except Exception:
    HAS_GRAPH = False


class ExamSchedulerGUI:
    """Giao diện người dùng cho hệ thống xếp lịch thi"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Xếp Lịch Thi Thông Minh - DSatur Pro v2.2")
        self.root.geometry("1200x800")
        
        # Backend
        self.backend = ExamSchedulerBackend()
        
        # Style
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
        """Thiết lập styles cho UI"""
        style = ttk.Style()
        try:
            style.theme_use('clam')
        except Exception:
            pass
        style.configure("TButton", padding=6, font=('Segoe UI', 10, 'bold'))
        style.configure("Treeview", background="white", fieldbackground="white", rowheight=26)
        style.map('Treeview', background=[('selected', self.colors['primary'])])
    
    def create_ui(self):
        """Tạo giao diện người dùng"""
        # Header
        header = tk.Frame(self.root, bg=self.colors['primary'], height=60)
        header.pack(fill='x')
        header.pack_propagate(False)
        tk.Label(header, text="XẾP LỊCH THI THÔNG MINH - DSATUR PRO", 
                font=('Segoe UI', 16, 'bold'),
                fg='white', bg=self.colors['primary']).pack(pady=12)
        
        main = tk.PanedWindow(self.root, orient=tk.HORIZONTAL, 
                            sashrelief=tk.RAISED, bg=self.colors['bg'])
        main.pack(fill='both', expand=True, padx=10, pady=10)
        
        # === SIDEBAR TRÁI ===
        self.create_sidebar(main)
        
        # === PHẦN PHẢI ===
        self.create_main_panel(main)
    
    def create_sidebar(self, parent):
        """Tạo sidebar bên trái"""
        left = tk.Frame(parent, bg=self.colors['card'], width=360, relief='flat')
        parent.add(left)
        
        # Upload
        upload_frame = tk.LabelFrame(left, text="NHẬP DỮ LIỆU", 
                                    bg=self.colors['card'], 
                                    fg=self.colors['dark'], 
                                    font=('Segoe UI', 11, 'bold'))
        upload_frame.pack(fill='x', padx=12, pady=8)
        
        tk.Button(upload_frame, text="CHỌN FILE EXCEL", command=self.load_file,
                 bg=self.colors['primary'], fg='white', 
                 font=('Segoe UI', 10, 'bold'),
                 relief='flat', padx=10, pady=8, 
                 cursor='hand2').pack(pady=8)
        
        self.file_label = tk.Label(upload_frame, text="Chưa chọn file...", 
                                   bg=self.colors['card'], fg='gray', 
                                   wraplength=320)
        self.file_label.pack(pady=4)
        
        # Cài đặt
        setting_frame = tk.LabelFrame(left, text="CÀI ĐẶT", 
                                     bg=self.colors['card'], 
                                     fg=self.colors['dark'], 
                                     font=('Segoe UI', 11, 'bold'))
        setting_frame.pack(fill='x', padx=12, pady=8)
        
        # Số ca tối đa mỗi ngày
        tk.Label(setting_frame, text="Số ca tối đa mỗi ngày:", 
                bg=self.colors['card'], 
                font=('Segoe UI', 10)).pack(anchor='w', padx=8, pady=5)
        self.max_var = tk.IntVar(value=3)
        tk.Spinbox(setting_frame, from_=1, to=10, textvariable=self.max_var, 
                  width=6, font=('Segoe UI', 10)).pack(anchor='w', padx=8, pady=(0,6))
        
        # Ngày bắt đầu thi
        tk.Label(setting_frame, text="Ngày bắt đầu thi (dd/mm/yyyy):", 
                bg=self.colors['card'], 
                font=('Segoe UI', 10)).pack(anchor='w', padx=8, pady=5)
        date_frame = tk.Frame(setting_frame, bg=self.colors['card'])
        date_frame.pack(anchor='w', padx=8, pady=(0,6))
        
        self.day_var = tk.StringVar(value=str(datetime.now().day))
        self.month_var = tk.StringVar(value=str(datetime.now().month))
        self.year_var = tk.StringVar(value=str(datetime.now().year))
        
        tk.Spinbox(date_frame, from_=1, to=31, textvariable=self.day_var, 
                  width=4, font=('Segoe UI', 9)).pack(side='left', padx=2)
        tk.Label(date_frame, text="/", bg=self.colors['card']).pack(side='left')
        tk.Spinbox(date_frame, from_=1, to=12, textvariable=self.month_var, 
                  width=4, font=('Segoe UI', 9)).pack(side='left', padx=2)
        tk.Label(date_frame, text="/", bg=self.colors['card']).pack(side='left')
        tk.Spinbox(date_frame, from_=2024, to=2035, textvariable=self.year_var, 
                  width=6, font=('Segoe UI', 9)).pack(side='left', padx=2)
        
        # Nút chạy
        tk.Button(left, text="CHẠY DSATUR", command=self.run_dsatur,
                 bg=self.colors['success'], fg='white', 
                 font=('Segoe UI', 12, 'bold'),
                 relief='flat', padx=10, pady=10, 
                 cursor='hand2').pack(pady=18, padx=12, fill='x')
        
        # Thống kê
        stats_frame = tk.LabelFrame(left, text="THỐNG KÊ", 
                                   bg=self.colors['card'], 
                                   fg=self.colors['dark'], 
                                   font=('Segoe UI', 11, 'bold'))
        stats_frame.pack(fill='both', expand=True, padx=12, pady=8)
        
        self.stats_text = tk.Text(stats_frame, height=12, 
                                 bg=self.colors['light'], 
                                 relief='flat', 
                                 font=('Consolas', 10))
        self.stats_text.pack(fill='both', padx=8, pady=8)
    
    def create_main_panel(self, parent):
        """Tạo panel chính bên phải"""
        right = tk.Frame(parent, bg=self.colors['card'])
        parent.add(right)
        
        notebook = ttk.Notebook(right)
        notebook.pack(fill='both', expand=True, padx=8, pady=8)
        
        # Tab 1: Lịch thi theo ngày
        self.create_tab_by_day(notebook)
        
        # Tab 2: Lịch thi theo ca
        self.create_tab_by_slot(notebook)
        
        # Tab 3: Lịch sinh viên
        self.create_tab_student(notebook)
        
        # Tab 4: Đồ thị
        self.create_tab_graph(notebook)
        
        # Tab 5: Export & Kiểm tra
        self.create_tab_export(notebook)
    
    def create_tab_by_day(self, notebook):
        """Tab lịch thi theo ngày"""
        tab1 = tk.Frame(notebook, bg='white')
        notebook.add(tab1, text='📅 Lịch Thi Theo Ngày')
        
        self.tree_day = ttk.Treeview(tab1, 
                                    columns=('Ngày', 'Ca', 'Môn', 'SV'), 
                                    show='headings', height=20)
        self.tree_day.heading('Ngày', text='Ngày Thi')
        self.tree_day.heading('Ca', text='Ca')
        self.tree_day.heading('Môn', text='Môn Học')
        self.tree_day.heading('SV', text='Số SV')
        
        self.tree_day.column('Ngày', width=140, anchor='center')
        self.tree_day.column('Ca', width=80, anchor='center')
        self.tree_day.column('Môn', width=420)
        self.tree_day.column('SV', width=80, anchor='center')
        
        scroll_day = ttk.Scrollbar(tab1, orient='vertical', 
                                  command=self.tree_day.yview)
        self.tree_day.configure(yscrollcommand=scroll_day.set)
        
        self.tree_day.pack(side='left', fill='both', expand=True, padx=8, pady=8)
        scroll_day.pack(side='right', fill='y')
    
    def create_tab_by_slot(self, notebook):
        """Tab lịch thi theo ca"""
        tab2 = tk.Frame(notebook, bg='white')
        notebook.add(tab2, text='🎯 Lịch Thi Theo Ca')
        
        self.tree_schedule = ttk.Treeview(tab2, 
                                         columns=('Ca', 'Môn', 'SV'), 
                                         show='headings', height=20)
        self.tree_schedule.heading('Ca', text='Ca Thi')
        self.tree_schedule.heading('Môn', text='Môn Học')
        self.tree_schedule.heading('SV', text='Số SV')
        
        self.tree_schedule.column('Ca', width=80, anchor='center')
        self.tree_schedule.column('Môn', width=540)
        self.tree_schedule.column('SV', width=80, anchor='center')
        
        scroll_schedule = ttk.Scrollbar(tab2, orient='vertical', 
                                       command=self.tree_schedule.yview)
        self.tree_schedule.configure(yscrollcommand=scroll_schedule.set)
        
        self.tree_schedule.pack(side='left', fill='both', expand=True, 
                               padx=8, pady=8)
        scroll_schedule.pack(side='right', fill='y')
    
    def create_tab_student(self, notebook):
        """Tab lịch sinh viên"""
        tab3 = tk.Frame(notebook, bg='white')
        notebook.add(tab3, text='👨‍🎓 Lịch Sinh Viên')
        
        # Search
        search_frame = tk.Frame(tab3, bg='white')
        search_frame.pack(fill='x', padx=8, pady=6)
        
        tk.Label(search_frame, text="Tìm:", bg='white', 
                font=('Segoe UI', 10)).pack(side='left')
        self.search_var = tk.StringVar()
        tk.Entry(search_frame, textvariable=self.search_var, 
                width=40, font=('Segoe UI', 10)).pack(side='left', padx=6)
        self.search_var.trace('w', self.filter_students)
        
        # Treeview
        self.tree_student = ttk.Treeview(tab3, 
                                        columns=('MSSV', 'Tên', 'Ngày', 'Ca', 'Môn'), 
                                        show='headings')
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
        
        scroll_student = ttk.Scrollbar(tab3, orient='vertical', 
                                      command=self.tree_student.yview)
        self.tree_student.configure(yscrollcommand=scroll_student.set)
        
        self.tree_student.pack(side='left', fill='both', expand=True, 
                              padx=8, pady=8)
        scroll_student.pack(side='right', fill='y')
    
    def create_tab_graph(self, notebook):
        """Tab đồ thị xung đột"""
        tab4 = tk.Frame(notebook, bg='white')
        notebook.add(tab4, text='📊 Đồ Thị Xung Đột')
        
        self.graph_canvas = tk.Canvas(tab4, bg='white')
        self.graph_canvas.pack(fill='both', expand=True, padx=8, pady=8)
        
        if not HAS_GRAPH:
            tk.Label(tab4, 
                    text="Cài networkx + matplotlib để xem đồ thị!", 
                    fg='red', 
                    font=('Segoe UI', 12)).pack(pady=50)
    
    def create_tab_export(self, notebook):
        """Tab export và kiểm tra"""
        tab5 = tk.Frame(notebook, bg='white')
        notebook.add(tab5, text='💾 Export & Kiểm Tra')
        
        tk.Button(tab5, text="XUẤT LỊCH THI EXCEL", 
                 command=self.export_excel,
                 bg=self.colors['warning'], fg='white', 
                 font=('Segoe UI', 11, 'bold'), 
                 pady=8).pack(pady=16)
        
        self.warning_text = tk.Text(tab5, height=12, 
                                   bg='#fff5f5', fg='red', 
                                   font=('Segoe UI', 10))
        self.warning_text.pack(fill='both', expand=True, padx=8, pady=8)
    
    # === EVENT HANDLERS ===
    
    def load_file(self):
        """Xử lý tải file Excel"""
        filepath = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not filepath:
            return
        
        success, message, stats = self.backend.load_excel_file(filepath)
        
        if success:
            self.file_label.config(
                text=f"ĐÃ TẢI: {os.path.basename(filepath)}\n"
                     f"{stats['records']} dòng • {stats['subjects']} môn",
                fg='green'
            )
            
            messagebox.showinfo(
                "Thành công",
                f"Đã tải thành công!\n\n"
                f"• {stats['records']:,} bản ghi\n"
                f"• {stats['sheets']} sheet\n"
                f"• {stats['students']} sinh viên\n"
                f"• {stats['subjects']} môn học"
            )
            
            self.update_stats()
        else:
            messagebox.showerror("Lỗi", message)
    
    def run_dsatur(self):
        """Chạy thuật toán DSatur"""
        # Lấy cấu hình
        try:
            start_date = datetime(
                int(self.year_var.get()), 
                int(self.month_var.get()), 
                int(self.day_var.get())
            )
        except Exception:
            messagebox.showerror("Lỗi", "Ngày tháng không hợp lệ!")
            return
        
        max_exams = int(self.max_var.get())
        
        # Chạy backend
        success, message, total_slots, total_days = self.backend.run_dsatur(
            max_exams_per_day=max_exams,
            start_date=start_date
        )
        
        if success:
            self.display_results()
            self.check_conflicts()
            self.draw_graph()
            self.update_stats()
            
            messagebox.showinfo(
                "HOÀN THÀNH",
                f"Đã xếp lịch thành công!\n\n"
                f"• Tổng ca thi: {total_slots}\n"
                f"• Số ca/ngày: {max_exams}\n"
                f"• Tổng số ngày thi: {total_days}"
            )
        else:
            messagebox.showwarning("Cảnh báo", message)
    
    def display_results(self):
        """Hiển thị kết quả lên UI"""
        # Xóa dữ liệu cũ
        for tree in [self.tree_day, self.tree_schedule, self.tree_student]:
            for item in tree.get_children():
                tree.delete(item)
        
        # Tab 1: Lịch theo ngày
        for item in self.backend.get_schedule_by_day():
            self.tree_day.insert('', 'end', values=(
                item['date'],
                f"Ca {item['session']}",
                item['subject'],
                item['students']
            ))
        
        # Tab 2: Lịch theo ca
        for item in self.backend.get_schedule_by_slot():
            self.tree_schedule.insert('', 'end', values=(
                f"Ca {item['slot']}",
                item['subject'],
                item['students']
            ))
        
        # Tab 3: Lịch sinh viên
        for item in self.backend.get_student_schedule():
            self.tree_student.insert('', 'end', values=(
                item['mssv'],
                item['name'],
                item['date'],
                f"Ca {item['session']}" if item['session'] > 0 else "",
                item['subject']
            ))
    
    def filter_students(self, *args):
        """Lọc sinh viên theo search"""
        search = self.search_var.get()
        
        # Xóa dữ liệu cũ
        for item in self.tree_student.get_children():
            self.tree_student.delete(item)
        
        # Hiển thị dữ liệu đã lọc
        for item in self.backend.get_student_schedule(search_term=search):
            self.tree_student.insert('', 'end', values=(
                item['mssv'],
                item['name'],
                item['date'],
                f"Ca {item['session']}" if item['session'] > 0 else "",
                item['subject']
            ))
    
    def check_conflicts(self):
        """Kiểm tra vi phạm"""
        has_conflicts, conflicts = self.backend.check_conflicts()
        
        self.warning_text.delete(1.0, 'end')
        
        if has_conflicts:
            text = "CÓ LỖI TRÙNG CA!\n\n"
            for conf in conflicts[:200]:
                text += f"TRÙNG: {conf['mssv']} - {conf['name']}\n"
            self.warning_text.insert('1.0', text)
            self.warning_text.config(fg='red')
        else:
            self.warning_text.insert(
                '1.0',
                "HOÀN HẢO! Không có sinh viên nào bị trùng ca thi\n\n"
                "✓ Tất cả sinh viên đều có lịch thi hợp lệ\n"
                "✓ Không có xung đột thời gian"
            )
            self.warning_text.config(fg='green')
    
    def update_stats(self):
        """Cập nhật thống kê"""
        stats = self.backend.get_statistics()
        
        text = f"TỔNG QUAN DỮ LIỆU\n"
        text += f"{'='*40}\n"
        text += f"Sinh viên: {stats['students']:,}\n"
        text += f"Môn học: {stats['subjects']:,}\n"
        text += f"Xung đột cạnh: {stats['conflicts']:,}\n"
        
        if stats['schedule_exists']:
            text += f"\n{'='*40}\n"
            text += f"LỊCH THI\n"
            text += f"{'='*40}\n"
            text += f"Tổng ca thi: {stats['total_slots']}\n"
            text += f"Ca/ngày: {stats['slots_per_day']}\n"
            text += f"Tổng số ngày: {stats['total_days']}\n"
        
        self.stats_text.delete(1.0, 'end')
        self.stats_text.insert('end', text)
    
    def draw_graph(self):
        """Vẽ đồ thị xung đột"""
        self.graph_canvas.delete('all')
        if not HAS_GRAPH:
            return
        
        try:
            graph_data = self.backend.get_graph_data()
            
            G = nx.Graph()
            for node in graph_data['nodes']:
                G.add_node(node['id'])
            for edge in graph_data['edges']:
                G.add_edge(edge['source'], edge['target'])
            
            plt.clf()
            fig = plt.figure(figsize=(8, 6))
            ax = fig.add_subplot(111)
            ax.axis('off')
            
            pos = nx.spring_layout(G, seed=42)
            
            # Color nodes by slot
            node_colors = [node['color'] for node in graph_data['nodes']]
            
            nx.draw_networkx_nodes(G, pos, node_size=300, 
                                  cmap=plt.cm.tab20, 
                                  node_color=node_colors)
            nx.draw_networkx_edges(G, pos, alpha=0.4)
            nx.draw_networkx_labels(G, pos, font_size=8)
            
            # Embed to tkinter
            canvas = FigureCanvasTkAgg(fig, master=self.graph_canvas)
            canvas.draw()
            canvas.get_tk_widget().pack(fill='both', expand=True)
        except Exception as e:
            print(f"Lỗi vẽ đồ thị: {e}")
    
    def export_excel(self):
        """Xuất file Excel"""
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx *.xls")],
            title="Lưu file lịch thi"
        )
        if not filepath:
            return
        
        success, message = self.backend.export_to_excel(filepath)
        
        if success:
            messagebox.showinfo(
                "Xuất thành công",
                f"Đã xuất lịch thi ra file:\n{os.path.basename(filepath)}"
            )
        else:
            messagebox.showerror("Lỗi xuất file", message)


def main():
    """Chạy ứng dụng"""
    root = tk.Tk()
    app = ExamSchedulerGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()