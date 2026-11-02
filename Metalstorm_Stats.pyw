import sqlite3
import threading
from datetime import datetime, timedelta
import os
import requests
import pandas as pd
import customtkinter as ctk
from tkinter import ttk, messagebox
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Ρύθμιση εμφάνισης CustomTkinter
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

# ---------------------------------------------------------
# 1. Διαχείριση SQLite Βάσης Δεδομένων
# ---------------------------------------------------------
DB_FILE = "metalstorm_stats.db"

def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS profile_snapshots (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp DATETIME,
            player_name TEXT,
            friend_code TEXT,
            trophies INTEGER,
            peak_trophies INTEGER,
            player_xp INTEGER,
            hangar_points INTEGER,
            total_matches INTEGER,
            total_wins INTEGER,
            total_deaths INTEGER,
            total_mvps INTEGER,
            win_rate REAL
        )
    """)
    conn.commit()
    conn.close()

def save_profile_snapshot(data):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO profile_snapshots (
            timestamp, player_name, friend_code, trophies, peak_trophies,
            player_xp, hangar_points, total_matches, total_wins,
            total_deaths, total_mvps, win_rate
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        datetime.now(),
        data['player_name'], data['friend_code'], data['trophies'],
        data['peak_trophies'], data['player_xp'], data['hangar_points'],
        data['total_matches'], data['total_wins'], data['total_deaths'],
        data['total_mvps'], data['win_rate']
    ))
    conn.commit()
    conn.close()

def get_24h_delta(current_data):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    
    target_time = datetime.now() - timedelta(hours=24)
    cursor.execute("""
        SELECT trophies, total_matches, total_wins, total_mvps, win_rate
        FROM profile_snapshots
        WHERE timestamp <= ?
        ORDER BY timestamp DESC LIMIT 1
    """, (target_time,))
    
    past_row = cursor.fetchone()
    
    if not past_row:
        cursor.execute("""
            SELECT trophies, total_matches, total_wins, total_mvps, win_rate
            FROM profile_snapshots
            ORDER BY timestamp DESC LIMIT 1 OFFSET 1
        """)
        past_row = cursor.fetchone()
        
    conn.close()
    
    if past_row:
        return {
            'trophies': current_data['trophies'] - past_row[0],
            'matches': current_data['total_matches'] - past_row[1],
            'wins': current_data['total_wins'] - past_row[2],
            'mvps': current_data['total_mvps'] - past_row[3],
            'win_rate': round(current_data['win_rate'] - past_row[4], 1)
        }
    return None

# ---------------------------------------------------------
# 2. Αθροιστική Εξαγωγή Excel
# ---------------------------------------------------------
def export_cumulative_excel(df_profile, df_planes, df_modes, file_name="Metalstorm_Stats_History.xlsx"):
    sheet_suffix = datetime.now().strftime("%Y-%m-%d_%H-%M")
    
    if os.path.exists(file_name):
        wb = openpyxl.load_workbook(file_name)
    else:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    header_font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=12)
    summary_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    summary_font = Font(name="Calibri", size=12, bold=True)

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'), right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'), bottom=Side(style='thin', color='CBD5E1')
    )
    totals_border = Border(
        left=Side(style='thin', color='CBD5E1'), right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='1E3A8A'), bottom=Side(style='double', color='1E3A8A')
    )

    sheets_to_add = [
        (f"Profile_{sheet_suffix}", df_profile),
        (f"Stats_{sheet_suffix}", df_planes),
        (f"Modes_{sheet_suffix}", df_modes)
    ]

    for sheet_title, df in sheets_to_add:
        ws = wb.create_sheet(title=sheet_title[:31])
        headers = list(df.columns)
        ws.append(headers)
        
        # Headers Styling
        for col_num in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
        # Data Rows
        for row in df.itertuples(index=False):
            ws.append(list(row))
            
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                cell.font = data_font
                cell.border = thin_border
                if isinstance(cell.value, float):
                    cell.number_format = '0.0'
                elif isinstance(cell.value, int):
                    cell.number_format = '#,##0'

        # Γραμμή Totals για την καρτέλα Stats
        if "Stats_" in sheet_title:
            last_row = ws.max_row
            tot_row_idx = last_row + 1
            
            totals_formulas = [
                'Totals',
                f'=SUM(B2:B{last_row})',
                f'=SUM(C2:C{last_row})',
                f'=AVERAGE(D2:D{last_row})',
                f'=SUM(E2:E{last_row})',
                f'=AVERAGE(F2:F{last_row})',
                f'=SUM(G2:G{last_row})',
                f'=AVERAGE(H2:H{last_row})'
            ]
            ws.append(totals_formulas)
            
            for col_num in range(1, len(totals_formulas) + 1):
                cell = ws.cell(row=tot_row_idx, column=col_num)
                cell.font = summary_font
                cell.fill = summary_fill
                cell.border = totals_border
                if col_num in [4, 6, 8]:
                    cell.number_format = '0.0'
                else:
                    cell.number_format = '#,##0'

        # Auto adjust column width
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

    wb.save(file_name)
    return file_name

# ---------------------------------------------------------
# 3. GUI Εφαρμογή
# ---------------------------------------------------------
class MetalstormApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        init_db()
        
        self.title("Metalstorm Stats Scout")
        self.geometry("1250x750")
        self.minsize(1000, 650)
        
        self.raw_data = None
        self.df_planes = pd.DataFrame()
        self.df_modes = pd.DataFrame()
        self.profile_data = {}
        
        self.setup_ui()

    def setup_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # --- SIDEBAR ---
        self.sidebar = ctk.CTkFrame(self, width=240, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_rowconfigure(6, weight=1)

        st_label = ctk.CTkLabel(self.sidebar, text="✈️ Metalstorm Scout", font=ctk.CTkFont(size=20, weight="bold"))
        st_label.grid(row=0, column=0, padx=20, pady=(20, 15))

        token_lbl = ctk.CTkLabel(self.sidebar, text="Startoken:", font=ctk.CTkFont(size=14, weight="bold"))
        token_lbl.grid(row=1, column=0, padx=20, pady=(10, 0), sticky="w")

        self.token_entry = ctk.CTkEntry(self.sidebar, show="*", placeholder_text="Επικολλήστε το token...", font=ctk.CTkFont(size=13))
        self.token_entry.grid(row=2, column=0, padx=20, pady=(5, 15), sticky="ew")

        self.enable_context_menu(self.token_entry)

        self.fetch_btn = ctk.CTkButton(self.sidebar, text="🔄 Φόρτωση Στατιστικών", font=ctk.CTkFont(size=14, weight="bold"), height=38, command=self.start_fetch_thread)
        self.fetch_btn.grid(row=3, column=0, padx=20, pady=10, sticky="ew")

        self.export_btn = ctk.CTkButton(self.sidebar, text="📥 Εξαγωγή σε Excel", font=ctk.CTkFont(size=14, weight="bold"), height=38, command=self.export_excel, fg_color="green", hover_color="darkgreen", state="disabled")
        self.export_btn.grid(row=4, column=0, padx=20, pady=10, sticky="ew")

        self.status_lbl = ctk.CTkLabel(self.sidebar, text="Έτοιμο", font=ctk.CTkFont(size=13), text_color="gray")
        self.status_lbl.grid(row=5, column=0, padx=20, pady=10, sticky="w")

        # --- MAIN TABVIEW ---
        self.tabview = ctk.CTkTabview(self)
        self.tabview.grid(row=0, column=1, padx=20, pady=10, sticky="nsew")
        
        self.tabview._segmented_button.configure(font=ctk.CTkFont(size=14, weight="bold"))

        self.tab_overview = self.tabview.add("👤 Overview & 24h Progress")
        self.tab_planes = self.tabview.add("✈ Aircraft Stats & Comparison")
        self.tab_modes = self.tabview.add("🎮 Game Modes")

        self.setup_overview_tab()
        self.setup_planes_tab()
        self.setup_modes_tab()

    def enable_context_menu(self, entry_widget):
        import tkinter as tk
        context_menu = tk.Menu(entry_widget, tearoff=0)
        context_menu.add_command(label="Αντιγραφή", command=lambda: entry_widget.event_generate("<<Copy>>"))
        context_menu.add_command(label="Επικόλληση", command=lambda: entry_widget.event_generate("<<Paste>>"))
        context_menu.add_command(label="Επιλογή Όλων", command=lambda: entry_widget.event_generate("<<SelectAll>>"))

        entry_widget.bind("<Button-3>", lambda e: context_menu.tk_popup(e.x_root, e.y_root))
        entry_widget.bind("<Control-v>", lambda e: entry_widget.event_generate("<<Paste>>"))
        entry_widget.bind("<Control-V>", lambda e: entry_widget.event_generate("<<Paste>>"))

    def setup_overview_tab(self):
        self.tab_overview.grid_columnconfigure((0, 1, 2), weight=1)
        
        self.header_lbl = ctk.CTkLabel(self.tab_overview, text="Παρακαλώ εισάγετε Token και πατήστε Φόρτωση", font=ctk.CTkFont(size=16, weight="bold"))
        self.header_lbl.grid(row=0, column=0, columnspan=3, padx=10, pady=15, sticky="w")

        self.cards_frame = ctk.CTkFrame(self.tab_overview, fg_color="transparent")
        self.cards_frame.grid(row=1, column=0, columnspan=3, sticky="nsew", padx=5, pady=5)
        self.cards_frame.grid_columnconfigure((0, 1, 2), weight=1)

    def setup_planes_tab(self):
        self.tab_planes.grid_rowconfigure(2, weight=1)
        self.tab_planes.grid_columnconfigure((0, 1), weight=1)

        comp_frame = ctk.CTkFrame(self.tab_planes)
        comp_frame.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky="ew")
        comp_frame.grid_columnconfigure((0, 1, 2, 3, 4), weight=1)

        ctk.CTkLabel(comp_frame, text="⚡ Σύγκριση:", font=ctk.CTkFont(size=15, weight="bold")).grid(row=0, column=0, padx=10, pady=10, sticky="w")

        self.combo_plane_a = ctk.CTkComboBox(comp_frame, values=["-"], font=ctk.CTkFont(size=14), command=self.update_plane_comparison)
        self.combo_plane_a.grid(row=0, column=1, padx=5, pady=10, sticky="ew")

        ctk.CTkLabel(comp_frame, text="VS", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=2, padx=5, pady=10)

        self.combo_plane_b = ctk.CTkComboBox(comp_frame, values=["-"], font=ctk.CTkFont(size=14), command=self.update_plane_comparison)
        self.combo_plane_b.grid(row=0, column=3, padx=5, pady=10, sticky="ew")

        self.clear_comp_btn = ctk.CTkButton(comp_frame, text="🧹 Καθαρισμός", font=ctk.CTkFont(size=13, weight="bold"), width=100, fg_color="#475569", hover_color="#334155", command=self.clear_comparison)
        self.clear_comp_btn.grid(row=0, column=4, padx=10, pady=10)

        self.comp_results_lbl = ctk.CTkLabel(comp_frame, text="Επιλέξτε 2 αεροσκάφη για σύγκριση", font=ctk.CTkFont(size=14))
        self.comp_results_lbl.grid(row=1, column=0, columnspan=5, padx=10, pady=(0, 10))

        ctk.CTkLabel(self.tab_planes, text="📋 Όλα τα Αεροσκάφη", font=ctk.CTkFont(size=16, weight="bold")).grid(row=1, column=0, padx=10, pady=(10, 0), sticky="w")
        
        # Αφαίρεση στήλης Trophies
        self.tree_planes = self.create_treeview(self.tab_planes, ["Aircraft", "Matches", "Wins", "Win Rate (%)", "Deaths", "Deaths/Match", "MVPs", "Avg Score"])
        self.tree_planes.grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

    def setup_modes_tab(self):
        self.tab_modes.grid_rowconfigure(0, weight=1)
        self.tab_modes.grid_columnconfigure(0, weight=1)
        self.tree_modes = self.create_treeview(self.tab_modes, ["Game Mode", "Matches", "Wins", "Win Rate (%)", "MVPs", "Deaths", "Avg Score"])
        self.tree_modes.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

    def create_treeview(self, parent, columns):
        style = ttk.Style()
        style.theme_use("clam")
        
        style.configure("Treeview", 
                        background="#1e293b", 
                        foreground="#f8fafc", 
                        fieldbackground="#1e293b", 
                        rowheight=32,
                        font=('Segoe UI', 12))
        
        style.configure("Treeview.Heading", 
                        background="#0f172a", 
                        foreground="#38bdf8", 
                        font=('Segoe UI', 13, 'bold'))
        
        style.map("Treeview", background=[('selected', '#3b82f6')])

        tree = ttk.Treeview(parent, columns=columns, show="headings")
        for col in columns:
            tree.heading(col, text=col, command=lambda c=col: self.sort_treeview(tree, c, False))
            tree.column(col, anchor="center", width=125)
        return tree

    def sort_treeview(self, tree, col, reverse):
        l = [(tree.set(k, col), k) for k in tree.get_children('')]
        try:
            l.sort(key=lambda x: float(str(x[0]).replace('%', '').replace(',', '')), reverse=reverse)
        except ValueError:
            l.sort(reverse=reverse)
        for index, (val, k) in enumerate(l):
            tree.move(k, '', index)
        tree.heading(col, command=lambda: self.sort_treeview(tree, col, not reverse))

    def clear_comparison(self):
        self.combo_plane_a.set("-")
        self.combo_plane_b.set("-")
        self.comp_results_lbl.configure(text="Επιλέξτε 2 αεροσκάφη για σύγκριση", font=ctk.CTkFont(size=14))

    def start_fetch_thread(self):
        token = self.token_entry.get().strip()
        if not token:
            messagebox.showwarning("Προσοχή", "Παρακαλώ εισάγετε το Startoken σας.")
            return

        self.fetch_btn.configure(state="disabled")
        self.status_lbl.configure(text="Φόρτωση...", text_color="#38bdf8")
        threading.Thread(target=self.fetch_data_async, args=(token,), daemon=True).start()

    def fetch_data_async(self, token):
        url = "https://playmetalstorm.com/stats/users/my"
        clean_token = token.replace("startoken ", "").strip()
        headers = {
            "Authorization": f"startoken {clean_token}",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
            "Accept": "application/json"
        }
        
        try:
            res = requests.get(url, headers=headers, timeout=10)
            if res.status_code == 200:
                self.raw_data = res.json()
                self.after(0, self.process_and_display_data)
            elif res.status_code == 401:
                self.after(0, self.show_error, "Μη έγκυρο ή ληγμένο Token.")
            else:
                self.after(0, self.show_error, f"Σφάλμα διακομιστή ({res.status_code}).")
        except Exception as e:
            self.after(0, self.show_error, f"Σφάλμα σύνδεσης: {str(e)}")

    def show_error(self, msg):
        self.status_lbl.configure(text="Σφάλμα", text_color="red")
        self.fetch_btn.configure(state="normal")
        messagebox.showerror("Σφάλμα", msg)

    def process_and_display_data(self):
        raw = self.raw_data
        try:
            accounts = raw.get("accounts", {})
            first_acc = list(accounts.keys())[0]
            player_data = accounts[first_acc].get("playerData", {})
        except Exception:
            player_data = raw

        p_name = player_data.get("playerName", "Unknown")
        f_code = player_data.get("friendCode", "-")
        trophies = player_data.get("cachedTotalVehicleRankingPoints", 0)
        peak_trophies = player_data.get("highestTotalVehicleRankingPoints", 0)
        player_xp = player_data.get("playerXp", 0)
        hangar_points = player_data.get("cachedHangerPointsSum", 0)

        career = player_data.get("careerCounters", {})
        overall = career.get("overall", {})
        
        t_matches = overall.get("pvpMatchesPlayed", 0)
        t_wins = overall.get("pvpMatchesWon", 0)
        t_mvps = overall.get("pvpMvpCount", 0)
        t_deaths = overall.get("pvpDeaths", 0)
        win_rate = round((t_wins / t_matches) * 100, 1) if t_matches > 0 else 0.0

        self.profile_data = {
            'player_name': p_name, 'friend_code': f_code, 'trophies': trophies,
            'peak_trophies': peak_trophies, 'player_xp': player_xp,
            'hangar_points': hangar_points, 'total_matches': t_matches,
            'total_wins': t_wins, 'total_deaths': t_deaths, 'total_mvps': t_mvps,
            'win_rate': win_rate
        }

        delta = get_24h_delta(self.profile_data)
        save_profile_snapshot(self.profile_data)

        now_str = datetime.now().strftime("%d/%m/%Y %H:%M")
        self.header_lbl.configure(text=f"👤 Παίκτης: {p_name}  |  Tag: {f_code}  |  📅 Ενημέρωση: {now_str}")
        self.update_metric_cards(delta)

        # Parse Planes (Χωρίς Trophies)
        planes_list = []
        for p_id, stats in career.get("countersByVehicleConfig", {}).items():
            c_name = p_id.replace("aircraft:", "").replace("-v2", "").replace("-v3", "").replace("-v4", "").replace("-", " ").title()
            
            m = stats.get("pvpMatchesPlayed", 0)
            w = stats.get("pvpMatchesWon", 0)
            d = stats.get("pvpDeaths", 0)
            s = stats.get("pvpScore", 0)
            mvp = stats.get("pvpMvpCount", 0)
            
            wr = round((w / m) * 100, 1) if m > 0 else 0.0
            avg_s = round(s / m, 1) if m > 0 else 0.0
            d_m = round(d / m, 2) if m > 0 else 0.0

            planes_list.append({
                "Aircraft": c_name, "Matches": m, "Wins": w, 
                "Win Rate (%)": wr, "Deaths": d, "Deaths/Match": d_m, "MVPs": mvp, "Avg Score": avg_s
            })

        self.df_planes = pd.DataFrame(planes_list).sort_values(by="Matches", ascending=False) if planes_list else pd.DataFrame()
        self.populate_treeview(self.tree_planes, self.df_planes)

        plane_names = sorted(self.df_planes["Aircraft"].tolist()) if not self.df_planes.empty else ["-"]
        self.combo_plane_a.configure(values=plane_names)
        self.combo_plane_b.configure(values=plane_names)

        # Parse Modes
        modes_list = []
        mode_map = {"ControlPoints": "Air Superiority", "Deathmatch": "Deathmatch", "PriorityTarget": "Priority Target", "CaptureTheFlag": "Capture The Flag", "Elimination": "Elimination"}
        for m_key, m_stats in career.get("countersByGameModeCategory", {}).items():
            clean_m = m_key.replace("GameModeCategory", "").replace("GameMode", "")
            disp_m = mode_map.get(clean_m, clean_m)
            mm = m_stats.get("pvpMatchesPlayed", 0)
            mw = m_stats.get("pvpMatchesWon", 0)
            mvp = m_stats.get("pvpMvpCount", 0)
            md = m_stats.get("pvpDeaths", 0)
            ms = m_stats.get("pvpScore", 0)
            mwr = round((mw / mm) * 100, 1) if mm > 0 else 0.0
            m_avg_s = round(ms / mm, 1) if mm > 0 else 0.0

            modes_list.append({
                "Game Mode": disp_m, "Matches": mm, "Wins": mw, "Win Rate (%)": mwr,
                "MVPs": mvp, "Deaths": md, "Avg Score": m_avg_s
            })

        self.df_modes = pd.DataFrame(modes_list).sort_values(by="Matches", ascending=False) if modes_list else pd.DataFrame()
        self.populate_treeview(self.tree_modes, self.df_modes)

        self.status_lbl.configure(text="Επιτυχία!", text_color="green")
        self.fetch_btn.configure(state="normal")
        self.export_btn.configure(state="normal")

    def update_metric_cards(self, delta):
        for widget in self.cards_frame.winfo_children():
            widget.destroy()

        def format_delta(val, is_percent=False):
            if val is None or val == 0:
                return ""
            symbol = "+" if val > 0 else ""
            unit = "%" if is_percent else ""
            return f" ({symbol}{val}{unit} 24h)"

        metrics = [
            ("Trophies 🏆", f"{self.profile_data['trophies']:,}", format_delta(delta['trophies'] if delta else None)),
            ("Max Trophies 🔝", f"{self.profile_data['peak_trophies']:,}", ""),
            ("Win Rate 🎯", f"{self.profile_data['win_rate']}%", format_delta(delta['win_rate'] if delta else None, True)),
            ("Matches Played ⚔️", f"{self.profile_data['total_matches']:,}", format_delta(delta['matches'] if delta else None)),
            ("MVPs 🎖️", f"{self.profile_data['total_mvps']:,}", format_delta(delta['mvps'] if delta else None)),
            ("Player XP ⭐", f"{self.profile_data['player_xp']:,}", "")
        ]

        for idx, (title, val, d_text) in enumerate(metrics):
            r, c = divmod(idx, 3)
            card = ctk.CTkFrame(self.cards_frame)
            card.grid(row=r, column=c, padx=10, pady=10, sticky="nsew")
            
            ctk.CTkLabel(card, text=title, font=ctk.CTkFont(size=14, weight="bold")).pack(pady=(12, 0))
            ctk.CTkLabel(card, text=val, font=ctk.CTkFont(size=22, weight="bold"), text_color="#38bdf8").pack(pady=4)
            if d_text:
                color = "#4ade80" if "+" in d_text else "#f87171"
                ctk.CTkLabel(card, text=d_text, font=ctk.CTkFont(size=12, weight="bold"), text_color=color).pack(pady=(0, 12))

    def populate_treeview(self, tree, df):
        for item in tree.get_children():
            tree.delete(item)
        for _, row in df.iterrows():
            tree.insert("", "end", values=list(row))

    def update_plane_comparison(self, _=None):
        p_a = self.combo_plane_a.get()
        p_b = self.combo_plane_b.get()

        if p_a == "-" or p_b == "-" or p_a == p_b or self.df_planes.empty:
            self.comp_results_lbl.configure(text="Επιλέξτε 2 διαφορετικά αεροσκάφη για σύγκριση.")
            return

        row_a = self.df_planes[self.df_planes["Aircraft"] == p_a].iloc[0]
        row_b = self.df_planes[self.df_planes["Aircraft"] == p_b].iloc[0]

        res_text = (
            f"📊 [{p_a}]  VS  [{p_b}]\n\n"
            f"Matches: {row_a['Matches']} vs {row_b['Matches']}   |   "
            f"Win Rate: {row_a['Win Rate (%)']}% vs {row_b['Win Rate (%)']}%\n"
            f"Deaths/Match: {row_a['Deaths/Match']} vs {row_b['Deaths/Match']}   |   "
            f"MVPs: {row_a['MVPs']} vs {row_b['MVPs']}   |   "
            f"Avg Score: {row_a['Avg Score']} vs {row_b['Avg Score']}"
        )
        self.comp_results_lbl.configure(text=res_text, font=ctk.CTkFont(size=14, weight="bold"))

    def export_excel(self):
        if self.df_planes.empty:
            return
        
        try:
            df_profile_exp = pd.DataFrame([{
                "Player": self.profile_data['player_name'],
                "Friend Code": self.profile_data['friend_code'],
                "Trophies": self.profile_data['trophies'],
                "Max Trophies": self.profile_data['peak_trophies'],
                "Player XP": self.profile_data['player_xp'],
                "Hangar Points": self.profile_data['hangar_points'],
                "Total Win Rate (%)": self.profile_data['win_rate'],
                "Total Matches": self.profile_data['total_matches'],
                "Total Wins": self.profile_data['total_wins'],
                "Deaths": self.profile_data['total_deaths'],
                "Deaths/Match": round(self.profile_data['total_deaths'] / self.profile_data['total_matches'], 2) if self.profile_data['total_matches'] > 0 else 0.0
            }])
            
            out_file = export_cumulative_excel(df_profile_exp, self.df_planes, self.df_modes)
            messagebox.showinfo("Επιτυχία", f"Τα δεδομένα προστέθηκαν με επιτυχία στο αρχείο:\n{out_file}")
        except Exception as e:
            messagebox.showerror("Σφάλμα Εξαγωγής", str(e))

if __name__ == "__main__":
    app = MetalstormApp()
    app.mainloop()