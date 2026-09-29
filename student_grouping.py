"""
Student Grouping and Stratification System (Stateful Auto-Match Edition)
"""

import os
import csv
import random
import collections
from dataclasses import dataclass
from typing import List, Dict, Set, Any, Optional

try:
    import pandas as pd
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    HAS_EXCEL_LIBS = True
except ImportError:
    HAS_EXCEL_LIBS = False

# Config
GROUPS_PER_BATCH_MAP = {1: 53, 2: 53, 3: 53} 
MAX_STUDENTS_PER_GROUP = 55

FACULTIES_LIST: List[str] = ["FTE", "FRI", "FIF", "FEB", "FKS", "FIK", "FIT", "FK"]
# nama depan yang dilewati waktu cek nama kembar dalam satu kelompok
EXCLUDED_FIRST_NAMES: Set[str] = {"mohammad", "muhammad", "muhamad", "mohamad", "mochamad", "mochammad", "muh", "moch"}


# Models
@dataclass
class Student:
    nim: str
    name: str
    email: str
    phone: str
    major: str
    faculty: str
    gender: str
    allocated_batch: int = 0
    allocated_group: int = 0
    is_new: bool = True

    @property
    def resolved_faculty(self) -> str:
        # samain nama panjang / singkatan fakultas jadi kode singkat
        if not self.faculty:
            return "UNKNOWN"
            
        fac_text = str(self.faculty).upper().strip()
        # "FK" harus paling bawah, kalau nggak "FKS" ikut ketangkep sebagai "FK"
        text_map = {
            "FAKULTAS TEKNIK ELEKTRO": "FTE", "FTE": "FTE",
            "FAKULTAS REKAYASA INDUSTRI": "FRI", "FRI": "FRI",
            "FAKULTAS INFORMATIKA": "FIF", "FIF": "FIF",
            "FAKULTAS EKONOMI & BISNIS": "FEB", "FEB": "FEB",
            "FAKULTAS KOMUNIKASI & ILMU SOSIAL": "FKS", "FKS": "FKS",
            "FAKULTAS INDUSTRI KREATIF": "FIK", "FIK": "FIK",
            "FAKULTAS ILMU TERAPAN": "FIT", "FIT": "FIT",
            "FAKULTAS KEDOKTERAN": "FK", "FK": "FK"
        }
        for key, val in text_map.items():
            if key in fac_text:
                return val
        
        return fac_text

    @property
    def clean_first_name(self) -> str:
        # kalau nama depan Muhammad dkk, pakai kata kedua
        if not self.name:
            return ""
        words = self.name.strip().split()
        if not words:
            return ""
        first_word = words[0].lower()
        if first_word in EXCLUDED_FIRST_NAMES and len(words) > 1:
            return words[1].lower()
        return first_word

    def is_valid(self) -> bool:
        return bool(
            self.name and isinstance(self.name, str) and
            self.nim and isinstance(self.nim, str)
        )


class Group:
    def __init__(self, group_id: int, batch_id: int):
        self.group_id: int = group_id
        self.batch_id: int = batch_id
        self.students: List[Student] = []
        self.first_names: Set[str] = set()
        self.major_counts: Dict[str, int] = collections.defaultdict(int)
        self.faculty_counts: Dict[str, int] = collections.defaultdict(int)
        self.gender_counts: Dict[str, int] = collections.defaultdict(int)

    def add_student(self, student: Student) -> None:
        self.students.append(student)
        if student.name:
            self.first_names.add(student.clean_first_name)
        self.major_counts[student.major.strip().upper()] += 1
        self.faculty_counts[student.resolved_faculty] += 1
        self.gender_counts[student.gender.strip().upper()] += 1

    @property
    def size(self) -> int:
        return len(self.students)

    def calculate_penalty(self, chunk: List[Student]) -> float:
        # makin kecil makin cocok chunk ini masuk ke kelompok ini
        c_major = chunk[0].major.strip().upper()
        c_faculty = chunk[0].resolved_faculty
        c_gender = chunk[0].gender.strip().upper()
        c_names = {s.clean_first_name for s in chunk if s.name}
        current_major_count = self.major_counts.get(c_major, 0)

        size_penalty = self.size * 57.0
        major_penalty = self.major_counts.get(c_major, 0) * 7.0
        gender_penalty = self.gender_counts.get(c_gender, 0) * 11.0
        faculty_penalty = self.faculty_counts.get(c_faculty, 0) * 21.0
        name_conflict_penalty = 175.0 if len(c_names.intersection(self.first_names)) > 0 else 0.0
        # chunk 1 orang dari prodi yang belum ada di kelompok ini dihindari
        if len(chunk) == 1 and current_major_count == 0:
            major_penalty += 300.0
        # cuma bikin kelompok penuh kalah, bukan larangan mutlak
        if self.size + len(chunk) > MAX_STUDENTS_PER_GROUP:
            size_penalty += 999999.0

        return size_penalty + major_penalty + gender_penalty + faculty_penalty + name_conflict_penalty


# Data ingestion
class DataIngestionService:
    @staticmethod
    def load_students(file_path: str) -> List[Student]:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File input tidak ditemukan: {file_path}")
            
        ext = os.path.splitext(file_path)[1].lower()
        return DataIngestionService._load_from_excel(file_path) if ext in ['.xlsx', '.xls'] else DataIngestionService._load_from_csv(file_path)

    @staticmethod
    def _load_from_excel(file_path: str) -> List[Student]:
        if not HAS_EXCEL_LIBS:
            raise ImportError("Library 'pandas' & 'openpyxl' diperlukan untuk membaca Excel.")
        
        students = []
        excel_dict = pd.read_excel(file_path, sheet_name=None, header=None)
        
        for _, df_raw in excel_dict.items():
            if df_raw.empty:
                continue
            
            # baris header nggak selalu di paling atas
            header_idx = DataIngestionService._find_header_row(df_raw)
            if header_idx is not None:
                headers = [str(v).strip().lower() for v in df_raw.iloc[header_idx]]
                df_clean = df_raw.iloc[header_idx + 1:].copy()
                df_clean.columns = headers
            else:
                df_clean = df_raw.copy()
                df_clean.columns = [str(c).strip().lower() for c in df_clean.columns]

            for _, row in df_clean.iterrows():
                student = DataIngestionService._row_to_student(row)
                if student.is_valid():
                    students.append(student)
        return students

    @staticmethod
    def _load_from_csv(file_path: str) -> List[Student]:
        students = []
        with open(file_path, mode='r', encoding='utf-8') as f:
            lines = f.readlines()
        
        # delimiter dan baris header dideteksi dari 30 baris pertama
        header_idx = 0
        delimiter = ';' if any(';' in line for line in lines[:10]) else ','
        for idx, line in enumerate(lines[:30]):
            tokens = [t.strip().lower() for t in line.split(delimiter)]
            if 'nim' in tokens and 'nama' in tokens:
                header_idx = idx
                break
        
        reader = csv.DictReader(lines[header_idx:], delimiter=delimiter)
        for row in reader:
            norm_row = {k.strip().lower(): v.strip() for k, v in row.items() if k}
            
            def parse_int(val):
                try: return int(float(val)) if val else 0
                except: return 0

            student = Student(
                name=norm_row.get("nama", ""),
                nim=norm_row.get("nim", ""),
                email=norm_row.get("email", ""),
                phone=norm_row.get("no telp", norm_row.get("phone", "")),
                major=norm_row.get("prodi", "").upper(),
                faculty=norm_row.get("fakultas", "").upper(),
                gender=norm_row.get("gender", norm_row.get("jenis kelamin", "")).upper(),
                allocated_batch=parse_int(norm_row.get("batch", norm_row.get("allocated_batch", 0))),
                allocated_group=parse_int(norm_row.get("kelompok", norm_row.get("allocated_group", 0)))
            )
            if student.is_valid():
                students.append(student)
        return students

    @staticmethod
    def _find_header_row(df: 'pd.DataFrame') -> Optional[int]:
        required = {'nim', 'nama'}
        for idx, row in df.head(30).iterrows():
            row_vals = [str(v).strip().lower() for v in row.values if pd.notna(v)]
            if required.issubset(set(row_vals)) and ({'gender', 'jenis kelamin'} & set(row_vals)):
                return idx
        return None

    @staticmethod
    def _row_to_student(row: 'pd.Series') -> Student:
        def parse_str(val):
            if pd.isna(val): return ""
            if isinstance(val, float):
                return str(int(val)) if val.is_integer() else str(val)
            return str(val).strip()

        def parse_int(val):
            try: return int(float(val)) if pd.notna(val) and str(val).strip() != "" else 0
            except: return 0

        get_val = lambda keys: next((row.get(k) for k in keys if k in row), "")

        # nama header di sini harus sama dengan yang ada di Excel
        return Student(
            name=parse_str(get_val(["nama"])),
            nim=parse_str(get_val(["nim"])),
            email=parse_str(get_val(["email"])),
            phone=parse_str(get_val(["no telp", "no hp", "phone"])),
            major=parse_str(get_val(["program studi", "prodi"])),
            faculty=parse_str(get_val(["fakultas"])),
            gender=parse_str(get_val(["gender", "jenis kelamin"])),
            allocated_batch=parse_int(get_val(["batch", "allocated_batch"])),
            allocated_group=parse_int(get_val(["kelompok", "allocated_group", "group"]))
        )


# Grouping
class StudentGroupingSystem:
    def __init__(self, groups_map: Dict[int, int] = None):
        self.groups_map = groups_map or GROUPS_PER_BATCH_MAP
        self.students: List[Student] = []

    def load_data(self, file_path: str, master_file_path: Optional[str] = None) -> None:
        raw_students = DataIngestionService.load_students(file_path)
        self.students = []
        
        # kalau ada master dari run sebelumnya, anak lama dipertahankan
        if master_file_path and os.path.exists(master_file_path):
            print(f"Membaca histori dari Master Data: {master_file_path}")
            master_students = DataIngestionService.load_students(master_file_path)

            master_dict = {s.nim: s for s in master_students if s.nim}
            for s in master_students:
                s.is_new = False
                self.students.append(s)
                
            # nim yang belum ada di master berarti maba baru
            new_count = 0
            for s in raw_students:
                if s.nim not in master_dict:
                    s.is_new = True
                    self.students.append(s)
                    master_dict[s.nim] = s
                    new_count += 1
                    
            print(f" -> Berhasil me-restore {len(master_students)} anak lama, dan menjaring {new_count} anak BARU dari file input.")
            
        else:
            # run pertama, nim dobel dibuang
            seen_nims = set()
            for s in raw_students:
                if s.nim not in seen_nims:
                    s.is_new = True
                    self.students.append(s)
                    seen_nims.add(s.nim)
                    
            print(f"Ingestion Success: Memuat {len(self.students)} mahasiswa untuk pertama kali.")

    def partition_batches_and_groups(self) -> Dict[int, List[Group]]:
        if not self.students:
            return {}

        # yang sudah punya batch & kelompok dikunci, sisanya dibagi ulang
        locked_students = [s for s in self.students if s.allocated_batch > 0 and s.allocated_group > 0]
        new_students = [s for s in self.students if s.allocated_batch == 0 or s.allocated_group == 0]

        batch_ids = list(self.groups_map.keys())
        est_new_per_batch = len(new_students) // len(batch_ids) if batch_ids else 0

        results: Dict[int, List[Group]] = {}
        
        # jumlah kelompok awal dihitung dari estimasi, dibuka seminimal mungkin
        for b_id, max_groups in self.groups_map.items():
            locked_b = [s for s in locked_students if s.allocated_batch == b_id]
            total_est = len(locked_b) + est_new_per_batch
            
            needed_groups = max(1, (total_est + MAX_STUDENTS_PER_GROUP - 1) // MAX_STUDENTS_PER_GROUP)
            
            highest_locked_group = max([s.allocated_group for s in locked_b] + [0])
            active_groups = max(needed_groups, highest_locked_group)
            active_groups = min(active_groups, max_groups) 

            results[b_id] = [Group(group_id=i + 1, batch_id=b_id) for i in range(active_groups)]

        # masukkan anak lama ke kelompoknya semula
        for s in locked_students:
            target_group = next((g for g in results.get(s.allocated_batch, []) if g.group_id == s.allocated_group), None)
            if target_group:
                target_group.add_student(s)
            else:
                new_students.append(s)

        if new_students:
            print(f"Mendistribusikan {len(new_students)} mahasiswa dengan target max. {MAX_STUDENTS_PER_GROUP} org/kelompok...")
            new_batch_pools = self._distribute_students_to_batches(new_students)
            for b_id, pool in new_batch_pools.items():
                print(f"Batch {b_id}: {len(pool)} mahasiswa baru, kapasitas total {self.groups_map[b_id] * MAX_STUDENTS_PER_GROUP}")
            
            for b_id, pool in new_batch_pools.items():
                if not pool: continue
                chunks = self._create_chunks(pool, base_size=2)
                for chunk in chunks:
                    best_group = min(results[b_id], key=lambda g: g.calculate_penalty(chunk))
                    # semua kelompok sudah penuh, buka kelompok baru selama belum mentok max_groups
                    if best_group.size + len(chunk) > MAX_STUDENTS_PER_GROUP:
                        if len(results[b_id]) < self.groups_map[b_id]:
                            best_group = Group(group_id=len(results[b_id]) + 1, batch_id=b_id)
                            results[b_id].append(best_group)
                        else:
                            # nggak ada yang muat chunk utuh, pecah dan isi slot sisa per orang
                            for student in chunk:
                                open_groups = [g for g in results[b_id] if g.size < MAX_STUDENTS_PER_GROUP]
                                if not open_groups:
                                    print(f"WARNING: batch {b_id} penuh total, {student.nim} terpaksa melebihi {MAX_STUDENTS_PER_GROUP}")
                                target = min(open_groups or results[b_id], key=lambda g: g.calculate_penalty([student]))
                                student.allocated_batch = b_id
                                student.allocated_group = target.group_id
                                target.add_student(student)
                            continue
                    for student in chunk:
                        student.allocated_batch = b_id
                        student.allocated_group = best_group.group_id
                        best_group.add_student(student)

        return results
    
    def _distribute_students_to_batches(self, students: List[Student]) -> Dict[int, List[Student]]:
        # round-robin per (prodi, gender) biar tiap batch sebarannya merata
        batch_pools: Dict[int, List[Student]] = {b: [] for b in self.groups_map}
        major_strata = collections.defaultdict(list)

        for s in students:
            major_strata[(s.major, s.gender)].append(s)

        batch_ids = list(self.groups_map.keys())
        current_b_idx = 0 

        for _, std_list in major_strata.items():
            random.shuffle(std_list)
            for student in std_list:
                target_batch = batch_ids[current_b_idx]
                batch_pools[target_batch].append(student)
                current_b_idx = (current_b_idx + 1) % len(batch_ids)

        return batch_pools

    def _create_chunks(self, pool: List[Student], base_size: int = 2) -> List[List[Student]]:
        # pecah per (prodi, gender) jadi potongan kecil, sisa pembagian ditambahkan ke chunk awal
        strata = collections.defaultdict(list)
        for s in pool:
            strata[(s.major, s.gender)].append(s)

        chunks = []
        for _, students in strata.items():
            random.shuffle(students)
            n = len(students)

            if n < base_size:
                chunks.append(students)
            else:
                num_chunks = n // base_size
                remainder = n % base_size
                start = 0
                for i in range(num_chunks):
                    size = base_size + (1 if i < remainder else 0)
                    chunks.append(students[start:start + size])
                    start += size

        # chunk terbesar ditempatkan duluan
        chunks.sort(key=len, reverse=True)
        return chunks


# Export Excel
class ExcelReportExporter:
    @staticmethod
    def export(batch_groups: Dict[int, List[Group]], all_students: List[Student], base_output_path: str) -> None:
        if not HAS_EXCEL_LIBS:
            raise ImportError("Membutuhkan library openpyxl untuk mengekspor Excel.")

        style = ExcelReportExporter._get_styles()
        fn, _ = os.path.splitext(base_output_path)

        # satu file per batch, ditambah satu file master untuk run berikutnya
        for b_id, groups in batch_groups.items():
            wb = openpyxl.Workbook()
            ExcelReportExporter._build_dashboard_tab(wb, b_id, groups, style)
            ExcelReportExporter._build_individual_group_tabs(wb, b_id, groups, style)
            output_name = f"{fn}_Batch_{b_id}.xlsx"
            wb.save(output_name)
            print(f" -> Berhasil mengexport laporan: '{output_name}'")

        ExcelReportExporter._export_master_data(all_students, f"{fn}_Master_Data.xlsx")

    @staticmethod
    def _export_master_data(students: List[Student], output_path: str) -> None:
        data = []
        for s in students:
            data.append({
                "NIM": s.nim,
                "Nama": s.name,
                "Email": s.email,
                "Gender": s.gender,
                "Fakultas": s.faculty,
                "Prodi": s.major,
                "No HP": s.phone,
                "Batch": s.allocated_batch,
                "Kelompok": s.allocated_group
            })
        
        df = pd.DataFrame(data)
        df.to_excel(output_path, index=False)
        print(f" -> Berhasil mengexport Master Data: '{output_path}'")

    @staticmethod
    def _get_styles() -> Dict[str, Any]:
        return {
            "title": Font(name="Segoe UI", size=15, bold=True, color="1F497D"),
            "subtitle": Font(name="Segoe UI", size=10, italic=True, color="595959"),
            "section": Font(name="Segoe UI", size=12, bold=True, color="1F497D"),
            "header": Font(name="Segoe UI", size=11, bold=True, color="FFFFFF"),
            "data": Font(name="Segoe UI", size=11),
            "font_new": Font(name="Segoe UI", size=11, bold=True, color="FF0000"),
            "bold": Font(name="Segoe UI", size=11, bold=True),
            "link": Font(name="Segoe UI", size=11, color="0000FF", underline="single"),
            "fill_header": PatternFill(start_color="2E4057", end_color="2E4057", fill_type="solid"),
            "fill_zebra": PatternFill(start_color="F7F9FA", end_color="F7F9FA", fill_type="solid"),
            "fill_total": PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"),
            "border": Border(
                left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
            ),
            "align_center": Alignment(horizontal="center", vertical="center"),
            "align_left": Alignment(horizontal="left", vertical="center")
        }

    @staticmethod
    def _build_dashboard_tab(wb, b_id, groups, style):
        ws = wb.active
        ws.title = f"Batch {b_id} Dashboard"
        total_students = sum(g.size for g in groups)
        
        ws["A2"] = f"LAPORAN PENGELOMPOKKAN MAHASISWA - BATCH {b_id}"
        ws["A2"].font = style["title"]
        ws["A3"] = f"Total Anggota: {total_students} Mahasiswa | Jumlah Kelompok Terkunci: {len(groups)}"
        ws["A3"].font = style["subtitle"]

        # rekap fakultas se-batch
        ws["A5"] = "REKAPAN TOTAL MAHASISWA PER FAKULTAS PADA BATCH INI"
        ws["A5"].font = style["section"]

        for c_idx, h_text in enumerate(FACULTIES_LIST + ["TOTAL MABA BATCH"], 1):
            cell = ws.cell(row=6, column=c_idx, value=h_text)
            cell.font = style["header"]; cell.fill = style["fill_header"]; cell.alignment = style["align_center"]

        fac_totals = collections.defaultdict(int)
        for g in groups:
            for fac, cnt in g.faculty_counts.items():
                fac_totals[fac] += cnt

        for c_idx, fac in enumerate(FACULTIES_LIST, 1):
            cell = ws.cell(row=7, column=c_idx, value=fac_totals[fac])
            cell.font = style["data"]; cell.border = style["border"]; cell.alignment = style["align_center"]

        tot_cell = ws.cell(row=7, column=len(FACULTIES_LIST) + 1, value=total_students)
        tot_cell.font = style["bold"]; tot_cell.fill = style["fill_total"]; tot_cell.border = style["border"]; tot_cell.alignment = style["align_center"]

        # tabel per kelompok
        ws["A10"] = "Daftar Kelompok & Sebaran Komposisi Anggota Rinci"
        ws["A10"].font = style["section"]

        headers = ["Nama Kelompok", "Total Anggota", "Laki-Laki", "Perempuan"] + FACULTIES_LIST + ["Link Akses"]
        for c_idx, h_text in enumerate(headers, 1):
            cell = ws.cell(row=11, column=c_idx, value=h_text)
            cell.font = style["header"]; cell.fill = style["fill_header"]
            cell.alignment = style["align_center"] if c_idx > 1 else style["align_left"]

        curr_r = 11
        for g in groups:
            curr_r += 1
            ws.cell(row=curr_r, column=1, value=f"Kelompok {g.group_id}")
            ws.cell(row=curr_r, column=2, value=g.size).alignment = style["align_center"]
            ws.cell(row=curr_r, column=3, value=g.gender_counts.get("LAKI-LAKI", 0)).alignment = style["align_center"]
            ws.cell(row=curr_r, column=4, value=g.gender_counts.get("PEREMPUAN", 0)).alignment = style["align_center"]

            for f_idx, fac in enumerate(FACULTIES_LIST, 5):
                ws.cell(row=curr_r, column=f_idx, value=g.faculty_counts.get(fac, 0)).alignment = style["align_center"]

            # kolom link ada setelah kolom fakultas terakhir
            link_col = 5 + len(FACULTIES_LIST)
            c_link = ws.cell(row=curr_r, column=link_col, value="Buka Sheet Regu →")
            c_link.font = style["link"]; c_link.hyperlink = f"#'Kelompok {g.group_id}'!A1"; c_link.alignment = style["align_center"]

            for col_i in range(1, link_col + 1):
                cell = ws.cell(row=curr_r, column=col_i)
                cell.border = style["border"]
                if col_i != link_col: cell.font = style["data"]
                if g.group_id % 2 == 0: cell.fill = style["fill_zebra"]

        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            ws.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + 4, 13)

    @staticmethod
    def _build_individual_group_tabs(wb, b_id, groups, style):
        for g in groups:
            g_name = f"Kelompok {g.group_id}"
            ws = wb.create_sheet(title=g_name)

            ws["A2"] = f"DAFTAR ANGGOTA - {g_name.upper()} (BATCH {b_id})"
            ws["A2"].font = style["title"]

            ws["A3"] = "← Kembali ke Dashboard Utama Batch"
            ws["A3"].font = style["link"]
            ws["A3"].hyperlink = f"#'Batch {b_id} Dashboard'!A1"

            t_headers = ["No", "NIM", "Nama Mahasiswa", "Email", "No HP", "Prodi", "Fakultas", "Gender", "Status"]
            for c_idx, h_text in enumerate(t_headers, 1):
                cell = ws.cell(row=6, column=c_idx, value=h_text)
                cell.font = style["header"]; cell.fill = style["fill_header"]
                cell.alignment = style["align_center"] if h_text in ["No", "NIM", "Gender", "No HP", "Status"] else style["align_left"]

            for r_idx, s in enumerate(g.students, 1):
                row_num = 6 + r_idx
                ws.cell(row=row_num, column=1, value=r_idx).alignment = style["align_center"]
                
                # NIM & No HP disimpan sebagai teks biar angka 0 di depan nggak hilang
                c_nim = ws.cell(row=row_num, column=2, value=str(s.nim))
                c_nim.number_format = '@'
                c_nim.alignment = style["align_center"]

                ws.cell(row=row_num, column=3, value=s.name).alignment = style["align_left"]
                ws.cell(row=row_num, column=4, value=s.email).alignment = style["align_left"]

                c_hp = ws.cell(row=row_num, column=5, value=str(s.phone))
                c_hp.number_format = '@'
                c_hp.alignment = style["align_center"]

                ws.cell(row=row_num, column=6, value=s.major).alignment = style["align_left"]
                ws.cell(row=row_num, column=7, value=s.faculty).alignment = style["align_left"]
                ws.cell(row=row_num, column=8, value=s.gender).alignment = style["align_center"]

                # anak baru ditandai NEW merah
                c_status = ws.cell(row=row_num, column=9, value="NEW" if s.is_new else "")
                c_status.alignment = style["align_center"]

                for col_i in range(1, 10):
                    cell = ws.cell(row=row_num, column=col_i)
                    cell.border = style["border"]
                
                    if col_i == 9 and s.is_new:
                        cell.font = style["font_new"]
                    else:
                        cell.font = style["data"]
                        
                    if r_idx % 2 == 0: 
                        cell.fill = style["fill_zebra"]

            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col if cell.row >= 6)
                ws.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + 4, 12)


# Main
if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description="Student Grouping System (Stateful Auto-Match)")
    parser.add_argument("--input", "-i", type=str, required=True, help="Path file data input raw (Excel/CSV).")
    parser.add_argument("--master", "-m", type=str, default=None, help="Path file Master Data dari run sebelumnya untuk mempertahankan kelompok.")
    parser.add_argument("--output", "-o", type=str, default="hasil_pengelompokkan", help="Nama base file output tanpa ekstensi.")
    args = parser.parse_args()

    system = StudentGroupingSystem(groups_map=GROUPS_PER_BATCH_MAP)
    system.load_data(args.input, args.master)

    batch_results = system.partition_batches_and_groups()
    
    ExcelReportExporter.export(batch_results, system.students, args.output)

    all_sizes = [g.size for b in batch_results.values() for g in b]
    min_size = min(all_sizes) if all_sizes else 0
    max_size = max(all_sizes) if all_sizes else 0

    print("\n" + "=" * 60)
    print("                    PEMERIKSAAN SELESAI                    ")
    print("=" * 60)
    print(f"-> Total Kelompok Aktif       : {len(all_sizes)} Kelompok")
    print(f"-> Ukuran Kelompok Terkecil   : {min_size} Mahasiswa")
    print(f"-> Ukuran Kelompok Terbesar   : {max_size} Mahasiswa")
    print("-" * 60)