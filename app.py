import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import plotly.express as px
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="Ödeal Ciro Simülasyonu", layout="wide")

st.title("📊 Ödeal Ciro & Hakkediş Simülasyonu")
st.markdown("Günlük kümülatif raporları tekilleştirin, net günlük satışları otomatik görün ve hakkedişlerinizi hesaplayın.")

MEMORY_DIR = "veri_hafizasi"
if not os.path.exists(MEMORY_DIR):
    os.makedirs(MEMORY_DIR)
MEMORY_FILE = os.path.join(MEMORY_DIR, "son_kumulatif.csv")

with st.sidebar:
    st.header("📂 Günlük Rapor Yükleme")
    file1 = st.file_uploader("Dinamik Pos Ciro Raporu", type=["xlsx"])
    file2 = st.file_uploader("Dinamik Pos (Genel) Raporu", type=["xlsx"])
    file3 = st.file_uploader("TTBP Raporu", type=["xlsx"])
    
    st.markdown("---")
    gun_sayisi = st.number_input("İçinde Bulunulan Ayın Gün Sayısı", min_value=1, max_value=31, value=15)
    toplam_gun = st.number_input("Bu Ay Toplam Kaç Gün?", min_value=28, max_value=31, value=30)
    
    calistir = st.button("🚀 Simülasyonu Çalıştır", use_container_width=True)

def get_hakkedis_rate(kanal):
    if kanal == 'Dinamik Pos': return 0.0025
    if kanal == 'Dinamik Pos Genel': return 0.0015
    if kanal == 'Dinamik Pos TTBP': return 0.0015
    return 0

def generate_excel(df, kanal_ozet, current_month):
    wb = Workbook()
    ws = wb.active
    ws.title = "Müşteri Analizi"
    
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)
        
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="2C3E50")
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    excel_file = io.BytesIO()
    wb.save(excel_file)
    excel_file.seek(0)
    return excel_file

if calistir and file1 and file2 and file3:
    with st.spinner("Veriler işleniyor..."):
        try:
            df1 = pd.read_excel(file1)
            df2 = pd.read_excel(file2)
            df3 = pd.read_excel(file3)
            
            df1['Kanal'] = 'Dinamik Pos'
            df2['Kanal'] = 'Dinamik Pos Genel'
            df3['Kanal'] = 'Dinamik Pos TTBP'
            
            df = pd.concat([df1, df2, df3], ignore_index=True)
            
            all_month_cols = sorted([col for col in df.columns if str(col).startswith('202')])
            month_cols = [col for col in all_month_cols if df[col].sum() > 0]
            
            current_month = month_cols[-1]
            last_3_months = month_cols[-4:-1]
            last_4_months = month_cols[-4:]
            
            for col in all_month_cols:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            
            grouped = df.groupby(['UyeIsyeriID', 'Unvan', 'Kanal'])[all_month_cols].sum().reset_index()
            grouped['Son_4_Ay_Toplam'] = grouped[last_4_months].sum(axis=1)
            active = grouped[grouped['Son_4_Ay_Toplam'] > 0].copy()
            
            if os.path.exists(MEMORY_FILE):
                gecmis_df = pd.read_csv(MEMORY_FILE)
                active = active.merge(gecmis_df[['UyeIsyeriID', 'Kanal', 'Son_Kumulatif']], on=['UyeIsyeriID', 'Kanal'], how='left')
                active['Son_Kumulatif'] = active['Son_Kumulatif'].fillna(0)
                active['Günlük_Net_Satis'] = active[current_month] - active['Son_Kumulatif']
                active['Günlük_Net_Satis'] = np.where(active['Günlük_Net_Satis'] < 0, 0, active['Günlük_Net_Satis'])
            else:
                active['Günlük_Net_Satis'] = 0
                st.info("💡 Sistem ilk kez çalıştırıldı. 'Günlük Net Satış' sonraki yüklemelerde net olarak hesaplanacaktır.")

            active['Aylık_Ort_Ciro_Gecmis'] = active[last_3_months].mean(axis=1)
            active['Günlük_Ort_Ciro_Gecmis'] = active['Aylık_Ort_Ciro_Gecmis'] / toplam_gun
            active['Mevcut_Günlük_Ort'] = active[current_month] / gun_sayisi
            active['Ay_Sonu_Projeksiyonu'] = active['Mevcut_Günlük_Ort'] * toplam_gun
            
            active['Hakkedis_Orani'] = active['Kanal'].apply(get_hakkedis_rate)
            active['Kümülatif_Hakkedis_TL'] = active[current_month] * active['Hakkedis_Orani']
            active['Projeksiyon_Hakkedis_TL'] = active['Ay_Sonu_Projeksiyonu'] * active['Hakkedis_Orani']
            
            active['Alarm'] = np.where(active['Mevcut_Günlük_Ort'] < active['Günlük_Ort_Ciro_Gecmis'], '📉 Ortalama Altı', '✅ İyi')
            active = active.sort_values(by=current_month, ascending=False)
            
            presentation_cols = ['Unvan', 'Kanal'] + last_3_months + [current_month, 'Günlük_Net_Satis', 'Aylık_Ort_Ciro_Gecmis', 'Mevcut_Günlük_Ort', 'Ay_Sonu_Projeksiyonu', 'Alarm', 'Kümülatif_Hakkedis_TL']
            df_presentation = active[presentation_cols].copy()
            
            hafiza_kayit = active[['UyeIsyeriID', 'Kanal', current_month]].copy()
            hafiza_kayit.rename(columns={current_month: 'Son_Kumulatif'}, inplace=True)
            hafiza_kayit.to_csv(MEMORY_FILE, index=False)
            
            st.success(f"Veriler başarıyla işlendi! Baz alınan güncel ay: {current_month}")
            
            # 5 Sütunlu Metrik Alanı (Ciro Projeksiyonu Eklendi)
            col1, col2, col3, col4, col5 = st.columns(5)
            col1.metric("Toplam Kümülatif Ciro", f"₺{active[current_month].sum():,.2f}")
            col2.metric("Günlük Net Satış (Fark)", f"₺{active['Günlük_Net_Satis'].sum():,.2f}")
            col3.metric("Ay Sonu Ciro Projeksiyonu", f"₺{active['Ay_Sonu_Projeksiyonu'].sum():,.2f}")
            col4.metric("Mevcut Kümülatif Hakkediş", f"₺{active['Kümülatif_Hakkedis_TL'].sum():,.2f}")
            col5.metric("Ay Sonu Hakkediş Projeksiyonu", f"₺{active['Projeksiyon_Hakkedis_TL'].sum():,.2f}")
            
            st.markdown("---")
            
            col_tbl, col_chart = st.columns([1, 2])
            with col_tbl:
                st.subheader("Kanal Bazlı Görünüm")
                kanal_ozet = active.groupby('Kanal').agg(
                    Net_Satis_Bugun=('Günlük_Net_Satis', 'sum'),
                    Kümülatif_Ciro=(current_month, 'sum'),
                    Kümülatif_Hakkedis=('Kümülatif_Hakkedis_TL', 'sum')
                ).reset_index()
                st.dataframe(kanal_ozet, use_container_width=True)
            
            with col_chart:
                st.subheader("İlk 20 Müşteri Seyri")
                top20 = active.head(20).copy()
                top20['Unvan'] = top20['Unvan'].apply(lambda x: str(x)[:20] + "..." if len(str(x)) > 20 else x)
                melted = top20.melt(id_vars=['Unvan'], value_vars=last_4_months, var_name='Ay', value_name='Ciro')
                fig = px.line(melted, x='Ay', y='Ciro', color='Unvan', markers=True)
                st.plotly_chart(fig, use_container_width=True)
            
            st.markdown("---")
            st.subheader("Müşteri Detay Analizi")
            st.dataframe(df_presentation, height=400, use_container_width=True)
            
            excel_data = generate_excel(df_presentation, kanal_ozet, current_month)
            st.download_button(
                label="📥 Formüllü Excel Raporunu İndir",
                data=excel_data,
                file_name=f"Odeal_Hakkedis_{current_month}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

        except Exception as e:
            st.error(f"Bir hata oluştu: {str(e)}")
elif calistir:
    st.warning("Lütfen sol panelden 3 kanal dosyasını da yükleyin.")
