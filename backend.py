"""
backend.py - Xử lý logic và thuật toán DSatur
"""
import pandas as pd
from collections import defaultdict
import heapq
from datetime import datetime, timedelta


class ExamSchedulerBackend:
    """Backend xử lý thuật toán DSatur và quản lý dữ liệu"""
    
    def __init__(self):
        # Dữ liệu
        self.data = None
        self.subjects = []
        self.student_subjects = defaultdict(set)  # {MSSV: {mon1, mon2, ...}}
        self.subject_students = defaultdict(set)  # {mon: {sv1, sv2, ...}}
        self.conflict_graph = defaultdict(set)    # {mon: {mon_xung_dot, ...}}
        self.schedule = {}                         # {mon: ca_thi}
        self.schedule_by_day = {}                  # {ngay: {ca: {subject, students, slot}}}
        
        # Cấu hình
        self.max_exams_per_day = 2
        self.start_date = datetime.now()
    
    def load_excel_file(self, filepath):
        """
        Đọc file Excel chứa danh sách lớp học phần
        Returns: (success: bool, message: str, stats: dict)
        """
        try:
            all_dfs = []
            excel = pd.ExcelFile(filepath, engine='openpyxl')
            
            for sheet in excel.sheet_names:
                try:
                    # Đọc toàn bộ sheet như string
                    df = pd.read_excel(excel, sheet_name=sheet, header=None, 
                                      dtype=str, engine='openpyxl')
                    df = df.fillna('')
                    
                    # Tìm dòng header (chứa "Mã SV" hoặc "MSSV")
                    header_row = None
                    for idx in range(min(5, len(df))):
                        row_text = ' '.join(df.iloc[idx].astype(str).str.lower().tolist())
                        if 'mã sv' in row_text or 'mssv' in row_text or 'ma sv' in row_text:
                            header_row = idx
                            break
                    
                    if header_row is None:
                        # Thử đọc sheet như 1 cột danh sách MSSV
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
                        if 'mã sv' in col_str or 'mssv' in col_str or 'ma sv' in col_str:
                            masv_col = col
                        if 'họ' in col_str and 'tên' in col_str:
                            hoten_col = col
                        elif 'tên' in col_str and hoten_col is None:
                            hoten_col = col
                    
                    if masv_col is None:
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
                return False, "Không tìm thấy dữ liệu hợp lệ!", None
            
            self.data = pd.concat(all_dfs, ignore_index=True)
            self.data.drop_duplicates(subset=['MaSV', 'ChuongTrinh'], inplace=True)
            
            stats = {
                'records': len(self.data),
                'sheets': len(excel.sheet_names),
                'students': self.data['MaSV'].nunique(),
                'subjects': self.data['ChuongTrinh'].nunique()
            }
            
            # Process data
            self.process_data()
            
            return True, "Tải file thành công!", stats
            
        except Exception as e:
            return False, f"Lỗi đọc file: {str(e)}", None
    
    def process_data(self):
        """Xử lý dữ liệu và xây dựng đồ thị xung đột"""
        self.subjects = sorted(self.data['ChuongTrinh'].unique().tolist())
        self.student_subjects.clear()
        self.subject_students.clear()
        self.conflict_graph.clear()
        
        # Build mappings
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
    
    def run_dsatur(self, max_exams_per_day=3, start_date=None):
        """
        Chạy thuật toán DSatur để xếp lịch thi
        Returns: (success: bool, message: str, total_slots: int, total_days: int)
        """
        if self.data is None or len(self.subjects) == 0:
            return False, "Chưa tải dữ liệu!", 0, 0
        
        self.max_exams_per_day = max_exams_per_day
        if start_date:
            self.start_date = start_date
        
        self.schedule.clear()
        self.schedule_by_day.clear()
        
        # DSatur algorithm
        degree = {s: len(self.conflict_graph[s]) for s in self.subjects}
        saturation = {s: 0 for s in self.subjects}
        color_of = {}
        
        # Build initial heap: (-saturation, -degree, subject)
        heap = [(-saturation[s], -degree[s], s) for s in self.subjects]
        heapq.heapify(heap)
        colored = set()
        
        while heap:
            _, _, subj = heapq.heappop(heap)
            if subj in colored:
                continue
            
            # Choose smallest color not used by neighbors
            used = {color_of.get(n) for n in self.conflict_graph[subj] if n in color_of}
            c = 1
            while c in used:
                c += 1
            color_of[subj] = c
            colored.add(subj)
            
            # Update neighbors' saturation and push back
            for nei in self.conflict_graph[subj]:
                if nei not in colored:
                    neigh_colors = {color_of.get(n) for n in self.conflict_graph[nei] 
                                   if n in color_of}
                    saturation[nei] = len(neigh_colors)
                    heapq.heappush(heap, (-saturation[nei], -degree[nei], nei))
        
        self.schedule = color_of
        
        # Tính toán lịch theo ngày
        self.calculate_schedule_by_day()
        
        total_slots = max(color_of.values()) if color_of else 0
        total_days = (total_slots + self.max_exams_per_day - 1) // self.max_exams_per_day
        
        return True, "Xếp lịch thành công!", total_slots, total_days
    
    def calculate_schedule_by_day(self):
        """Tính toán lịch thi theo ngày dựa trên số ca tối đa mỗi ngày"""
        self.schedule_by_day.clear()
        
        for subject, slot in self.schedule.items():
            # Tính ngày thi
            day_index = (slot - 1) // self.max_exams_per_day
            session_in_day = ((slot - 1) % self.max_exams_per_day) + 1
            
            exam_date = self.start_date + timedelta(days=day_index)
            date_str = exam_date.strftime("%d/%m/%Y")
            
            if date_str not in self.schedule_by_day:
                self.schedule_by_day[date_str] = {}
            
            self.schedule_by_day[date_str][session_in_day] = {
                'subject': subject,
                'students': len(self.subject_students[subject]),
                'slot': slot
            }
    
    def check_conflicts(self):
        """
        Kiểm tra vi phạm ràng buộc cứng (trùng ca thi)
        Returns: (has_conflicts: bool, conflicts: list)
        """
        conflicts = []
        
        for sid, subs in self.student_subjects.items():
            cas = [self.schedule.get(s) for s in subs]
            unique_cas = set([c for c in cas if c is not None])
            
            if len([c for c in cas if c is not None]) != len(unique_cas):
                name_df = self.data.loc[self.data['MaSV'] == sid, 'HoTen']
                name = name_df.iloc[0] if len(name_df) > 0 else "N/A"
                conflicts.append({
                    'mssv': sid,
                    'name': name,
                    'cas': list(unique_cas)
                })
        
        return len(conflicts) > 0, conflicts
    
    def get_statistics(self):
        """Lấy thống kê hệ thống"""
        stats = {
            'students': len(self.student_subjects),
            'subjects': len(self.subjects),
            'conflicts': sum(len(v) for v in self.conflict_graph.values()) // 2,
            'schedule_exists': len(self.schedule) > 0
        }
        
        if self.schedule:
            total_slots = max(self.schedule.values())
            total_days = (total_slots + self.max_exams_per_day - 1) // self.max_exams_per_day
            stats.update({
                'total_slots': total_slots,
                'total_days': total_days,
                'slots_per_day': self.max_exams_per_day
            })
        
        return stats
    
    def get_schedule_by_day(self):
        """Lấy lịch thi theo ngày (sorted)"""
        result = []
        sorted_dates = sorted(self.schedule_by_day.keys(),
                            key=lambda x: datetime.strptime(x, "%d/%m/%Y"))
        
        for date in sorted_dates:
            sessions = self.schedule_by_day[date]
            for session in sorted(sessions.keys()):
                info = sessions[session]
                result.append({
                    'date': date,
                    'session': session,
                    'subject': info['subject'],
                    'students': info['students'],
                    'slot': info['slot']
                })
        
        return result
    
    def get_schedule_by_slot(self):
        """Lấy lịch thi theo ca"""
        result = []
        ca_dict = defaultdict(list)
        
        for subj, ca in self.schedule.items():
            ca_dict[ca].append((subj, len(self.subject_students[subj])))
        
        for ca in sorted(ca_dict.keys()):
            for subj, count in sorted(ca_dict[ca], key=lambda x: -x[1]):
                result.append({
                    'slot': ca,
                    'subject': subj,
                    'students': count
                })
        
        return result
    
    def get_student_schedule(self, search_term=None):
        """Lấy lịch thi của sinh viên"""
        result = []
        
        for sid, subs in self.student_subjects.items():
            name_df = self.data.loc[self.data['MaSV'] == sid, 'HoTen']
            name = name_df.iloc[0] if len(name_df) > 0 else "N/A"
            
            # Filter by search term
            if search_term:
                if search_term.lower() not in sid.lower() and \
                   search_term.lower() not in name.lower():
                    continue
            
            for sub in sorted(subs):
                slot = self.schedule.get(sub, 0)
                if slot == 0:
                    date_str = ""
                    session_in_day = 0
                else:
                    day_index = (slot - 1) // self.max_exams_per_day
                    session_in_day = ((slot - 1) % self.max_exams_per_day) + 1
                    exam_date = self.start_date + timedelta(days=day_index)
                    date_str = exam_date.strftime("%d/%m/%Y")
                
                result.append({
                    'mssv': sid,
                    'name': name,
                    'date': date_str,
                    'session': session_in_day,
                    'subject': sub,
                    'slot': slot
                })
        
        return result
    
    def export_to_excel(self, filepath):
        """Xuất lịch thi ra file Excel"""
        if not self.schedule:
            return False, "Chưa có lịch để xuất!"
        
        try:
            # Lịch theo ngày
            day_rows = []
            for item in self.get_schedule_by_day():
                day_rows.append({
                    'Ngày': item['date'],
                    'Ca trong ngày': f"Ca {item['session']}",
                    'Ca toàn bộ (DSatur)': item['slot'],
                    'Môn': item['subject'],
                    'Số SV': item['students']
                })
            df_day = pd.DataFrame(day_rows)
            
            # Lịch theo ca
            ca_rows = []
            for item in self.get_schedule_by_slot():
                ca_rows.append({
                    'Ca toàn bộ': f"Ca {item['slot']}",
                    'Môn': item['subject'],
                    'Số SV': item['students']
                })
            df_ca = pd.DataFrame(ca_rows)
            
            # Lịch sinh viên
            stu_rows = []
            for item in self.get_student_schedule():
                stu_rows.append({
                    'MSSV': item['mssv'],
                    'Họ Tên': item['name'],
                    'Môn': item['subject'],
                    'Ngày Thi': item['date'],
                    'Ca trong ngày': f"Ca {item['session']}" if item['session'] > 0 else "",
                    'Ca toàn bộ': f"Ca {item['slot']}" if item['slot'] > 0 else ""
                })
            df_stu = pd.DataFrame(stu_rows)
            
            # Thống kê
            stats = self.get_statistics()
            summary = {
                'Tổng sinh viên': [stats['students']],
                'Tổng môn': [stats['subjects']],
                'Tổng ca (toàn bộ)': [stats.get('total_slots', 0)],
                'Số ca/ngày (cấu hình)': [self.max_exams_per_day],
                'Ngày bắt đầu': [self.start_date.strftime("%d/%m/%Y")]
            }
            df_sum = pd.DataFrame(summary)
            
            # Ghi ra Excel
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                df_day.to_excel(writer, sheet_name='Lich_Theo_Ngay', index=False)
                df_ca.to_excel(writer, sheet_name='Lich_Theo_Ca', index=False)
                df_stu.to_excel(writer, sheet_name='Lich_SinhVien', index=False)
                df_sum.to_excel(writer, sheet_name='ThongTin_TomTat', index=False)
            
            return True, "Xuất file thành công!"
            
        except Exception as e:
            return False, f"Lỗi xuất file: {str(e)}"
    
    def get_graph_data(self):
        """Lấy dữ liệu đồ thị để vẽ"""
        nodes = []
        edges = []
        
        for subject in self.subjects:
            nodes.append({
                'id': subject,
                'color': self.schedule.get(subject, 0)
            })
        
        processed = set()
        for a, neighs in self.conflict_graph.items():
            for b in neighs:
                edge_key = tuple(sorted([a, b]))
                if edge_key not in processed:
                    edges.append({'source': a, 'target': b})
                    processed.add(edge_key)
        
        return {'nodes': nodes, 'edges': edges}