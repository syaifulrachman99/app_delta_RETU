import streamlit as st
import pandas as pd
import datetime
import os
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# Konfigurasi halaman web Streamlit
st.set_page_config(
    page_title="Aplikasi Delta Parameter Jaringan", 
    page_icon="📊", 
    layout="wide"
)

st.title("📊 Aplikasi Penghitung Delta Parameter RETU & CONCHEK")
st.markdown("Unggah file CSV Anda di bawah ini, lalu klik tombol **Proses Laporan** untuk menghasilkan file Excel secara otomatis.")

# Sidebar untuk upload file
st.sidebar.header("📁 Unggah Berkas Data")
uploaded_retu = st.sidebar.file_uploader("Upload file `retu.csv`", type=["csv"])
uploaded_conchek = st.sidebar.file_uploader("Upload file `conchek.csv`", type=["csv"])
uploaded_sitelist = st.sidebar.file_uploader("Upload file `site list.xlsx`", type=["xlsx", "xls"])

# Tombol proses utama
if st.sidebar.button("🚀 Proses Laporan Sekarang", type="primary"):
    if uploaded_retu is not None and uploaded_conchek is not None:
        with st.spinner("Sedang memproses data dan menghitung delta... Mohon tunggu."):
            
            # --- FUNGSI PENGOLAHAN DATA ---
            DATE_FORMAT = "%m/%d/%Y"
            CONCHEK_PARAMS = ["pMax", "dlRsBoost"]

            def baca_csv_uploaded(uploaded_file):
                lines = []
                for line in uploaded_file.getvalue().decode("utf-8-sig", errors="ignore").splitlines():
                    lines.append(line.strip())
                if not lines:
                    raise ValueError("File kosong.")
                sep = ';' if ';' in lines[0] else ','
                parsed_data = [[val.strip().strip('"').strip("'") for val in line.split(sep)] for line in lines if line]
                
                df = pd.DataFrame(parsed_data[1:], columns=parsed_data[0])
                df.columns = df.columns.str.strip().str.strip('"').str.strip("'")
                
                kolom_lower = {c.lower(): c for c in df.columns}
                if "xdate" not in df.columns:
                    if "xdate" in kolom_lower:
                        df = df.rename(columns={kolom_lower["xdate"]: "xdate"})
                    elif "date" in kolom_lower:
                        df = df.rename(columns={kolom_lower["date"]: "xdate"})
                    else:
                        raise KeyError(f"Kolom 'xdate' atau 'date' tidak ditemukan. Kolom yang ada: {list(df.columns)}")
                
                df["xdate"] = pd.to_datetime(df["xdate"], format="mixed", errors="coerce")
                return df

            def baca_sitelist_uploaded(uploaded_file):
                if uploaded_file is None:
                    return []
                df = pd.read_excel(uploaded_file)
                kolom_site = [c for c in df.columns if 'site' in str(c).lower()]
                kolom = kolom_site[0] if kolom_site else df.columns[0]
                return df[kolom].dropna().astype(str).str.strip().unique().tolist()

            # Baca data dari ulasan web
            conchek = baca_csv_uploaded(uploaded_conchek)
            retu = baca_csv_uploaded(uploaded_retu)
            sitelist = baca_sitelist_uploaded(uploaded_sitelist)

            # Proses RETU
            peta = conchek.drop_duplicates("MRBTS").set_index("MRBTS")["Site_ID"] if "MRBTS" in conchek.columns and "Site_ID" in conchek.columns else None
            retu["site_id"] = retu["MRBTS"].map(peta) if peta is not None else None
            
            kolom_dikenal = {"oss", "xdate", "distname", "distName", "mrbts", "MRBTS", "ald", "ald", "retu", "RETU", "basestationid", "sectorid", "sectorID", "eqm", "apeqm", "version", "antmodel", "antserial", "angle", "site_id"}
            kolom_aneh = [c for c in retu.columns if c not in kolom_dikenal]
            if kolom_aneh and "site_id" in retu.columns:
                retu["site_id"] = retu["site_id"].fillna(retu[kolom_aneh[0]])

            retu["angle"] = pd.to_numeric(retu["angle"], errors="coerce")
            retu = retu.sort_values(["distName", "xdate"]) if "distName" in retu.columns else retu
            retu["delta"] = retu.groupby("distName")["angle"].diff() if "distName" in retu.columns else 0

            if "sectorID" not in retu.columns:
                retu["sectorID"] = "-"

            out_retu = retu[["xdate", "distName", "sectorID", "RETU", "site_id", "angle", "delta"]].copy()
            out_retu.columns = ["date", "distname", "sector_id", "retu", "site_id", "angle", "delta"]

            # Proses CONCHEK
            conchek = conchek.sort_values(["Site_ID", "LNCEL", "xdate"])
            hasil_conchek = []
            for p in CONCHEK_PARAMS:
                if p in conchek.columns:
                    conchek[p] = pd.to_numeric(conchek[p], errors="coerce")
                    d = conchek[["xdate", "distName", "LNCEL", "Site_ID", p]].copy()
                    d["delta"] = d.groupby(["Site_ID", "LNCEL"])[p].diff()
                    d = d.rename(columns={"xdate": "date", "distName": "distname", "LNCEL": "lncel", "Site_ID": "site_id", p: "nilai"})
                    d.insert(4, "parameter", p)
                    hasil_conchek.append(d)
            
            out_conchek = pd.concat(hasil_conchek, ignore_index=True) if hasil_conchek else pd.DataFrame()

            # Filter Sitelist
            if sitelist:
                out_retu = out_retu[out_retu["site_id"].isin(sitelist)]
                if not out_conchek.empty:
                    out_conchek = out_conchek[out_conchek["site_id"].isin(sitelist)]
                site_di_retu = set(out_retu["site_id"].dropna().unique())
                site_tanpa_retu = [s for s in sitelist if s not in site_di_retu]
            else:
                site_tanpa_retu = []

            # Buat Workbook Excel
            wb = Workbook()
            wb.remove(wb.active)

            def label_h(tanggal, tanggal_h):
                selisih = (tanggal_h - tanggal).days
                return "H" if selisih == 0 else f"H-{selisih}"

            def buat_mendatar(df, kolom_kunci, kolom_nilai, tanggal_h=None):
                if df.empty:
                    return pd.DataFrame(), kolom_kunci, kolom_nilai, [], []
                if tanggal_h is None:
                    tanggal_h = df["date"].max()
                df = df[df["date"] <= tanggal_h].copy()
                df["hari"] = df["date"].apply(lambda t: label_h(t, tanggal_h))
                tanggal_urut = sorted(df["date"].unique())
                label_urut = [label_h(t, tanggal_h) for t in tanggal_urut]
                lebar = df.pivot_table(index=kolom_kunci, columns="hari", values=kolom_nilai, aggfunc="first")
                urutan = [(v, h) for h in label_urut for v in kolom_nilai]
                lebar = lebar.reindex(columns=pd.MultiIndex.from_tuples(urutan)).reset_index().sort_values(kolom_kunci)
                return lebar, kolom_kunci, kolom_nilai, label_urut, tanggal_urut

            def tulis_sheet(ws, lebar, kolom_kunci, kolom_nilai, label_urut, tanggal_urut):
                if lebar.empty:
                    ws.append(["Data kosong"])
                    return
                biru = PatternFill("solid", fgColor="1F4E78")
                abu = PatternFill("solid", fgColor="D9E1F2")
                fill_biru_muda = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
                tebal_putih = Font(bold=True, color="FFFFFF")
                tengah = Alignment(horizontal="center", vertical="center")
                sisi = Side(style="thin", color="999999")
                garis = Border(left=sisi, right=sisi, top=sisi, bottom=sisi)

                n_kunci, n_nilai = len(kolom_kunci), len(kolom_nilai)
                for i, nama in enumerate(kolom_kunci, start=1):
                    ws.cell(row=1, column=i, value=nama)
                    ws.merge_cells(start_row=1, start_column=i, end_row=3, end_column=i)
                    c = ws.cell(row=1, column=i)
                    c.fill, c.font, c.alignment = biru, tebal_putih, tengah

                for j, (h, t) in enumerate(zip(label_urut, tanggal_urut)):
                    c0 = n_kunci + 1 + j * n_nilai
                    ws.cell(row=1, column=c0, value=h)
                    ws.cell(row=2, column=c0, value=t.strftime(DATE_FORMAT))
                    ws.merge_cells(start_row=1, start_column=c0, end_row=1, end_column=c0 + n_nilai - 1)
                    ws.merge_cells(start_row=2, start_column=c0, end_row=2, end_column=c0 + n_nilai - 1)
                    warna = biru if h == "H" else abu
                    for k, v in enumerate(kolom_nilai):
                        ws.cell(row=3, column=c0 + k, value=v)
                    for r in (1, 2, 3):
                        for k in range(n_nilai):
                            c = ws.cell(row=r, column=c0 + k)
                            c.fill, c.font, c.alignment = warna, (tebal_putih if h == "H" else Font(bold=True)), tengah

                for r in (1, 2, 3):
                    for c in range(1, n_kunci + n_nilai * len(label_urut) + 1):
                        ws.cell(row=r, column=c).border = garis

                row_idx = 4
                for baris in lebar.itertuples(index=False):
                    val_list = [None if pd.isna(x) else (x.item() if hasattr(x, "item") else x) for x in baris]
                    ws.append(val_list)
                    for col_idx_1based, val in enumerate(val_list, start=1):
                        if col_idx_1based > n_kunci:
                            rel_idx = (col_idx_1based - n_kunci - 1) % n_nilai
                            if kolom_nilai[rel_idx] == "delta":
                                if val is not None and str(val).strip() not in ["", "0", "0.0"]:
                                    ws.cell(row=row_idx, column=col_idx_1based).fill = fill_biru_muda
                    row_idx += 1

                ws.freeze_panes = ws.cell(row=4, column=n_kunci + 1)
                ws.auto_filter.ref = f"A3:{get_column_letter(ws.max_column)}{ws.max_row}"

            # Tulis Sheet RETU
            hasil_retu = buat_mendatar(out_retu, ["site_id", "distname", "sector_id", "retu"], ["angle", "delta"])
            ws_retu = wb.create_sheet("retu")
            tulis_sheet(ws_retu, *hasil_retu)

            # Tulis Sheet Conchek
            if not out_conchek.empty:
                for p in CONCHEK_PARAMS:
                    df_param = out_conchek[out_conchek["parameter"] == p].copy()
                    if not df_param.empty:
                        hasil_c = buat_mendatar(df_param, ["site_id", "distname", "lncel"], ["nilai", "delta"])
                        ws_c = wb.create_sheet(f"conchek_{p}")
                        tulis_sheet(ws_c, *hasil_c)

            # Sheet Tidak Ada RETU
            if site_tanpa_retu:
                ws_missing = wb.create_sheet("Tidak Ada RETU")
                ws_missing.append(["Site_ID yang Tidak Ada di RETU"])
                ws_missing.cell(row=1, column=1).fill = PatternFill("solid", fgColor="C00000")
                ws_missing.cell(row=1, column=1).font = Font(bold=True, color="FFFFFF")
                for site in sorted(site_tanpa_retu):
                    ws_missing.append([site])

            # Simpan ke Buffer memori untuk tombol download
            output_path = "hasil_delta.xlsx"
            wb.save(output_path)
            
            st.success("✅ Proses Berhasil! Laporan Excel Anda sudah siap.")
            
            # Tombol Download di Web
            with open(output_path, "rb") as f:
                st.download_button(
                    label="📥 Unduh File Excel (hasil_delta.xlsx)",
                    data=f,
                    file_name="hasil_delta.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    else:
        st.warning("⚠️ Mohon unggah file `retu.csv` dan `conchek.csv` terlebih dahulu melalui panel di sebelah kiri.")