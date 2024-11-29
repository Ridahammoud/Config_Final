import pandas as pd
import streamlit as st
from datetime import datetime, timedelta
import plotly.express as px
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

@st.cache_data
def charger_donnees(fichier):
    return pd.read_excel(fichier)

def convert_df_to_xlsx(df):
    # Convert DataFrame to Excel using openpyxl
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
    return output.getvalue()

def convert_df_to_pdf(df):
    # Convert DataFrame to PDF using ReportLab
    output = BytesIO()
    c = canvas.Canvas(output, pagesize=letter)
    width, height = letter
    
    # Set up table properties
    x_offset = 30
    y_offset = height - 50
    row_height = 20
    col_widths = [100, 100, 100, 100]  # adjust this for your columns
    
    # Header
    c.setFont("Helvetica-Bold", 10)
    columns = df.columns.tolist()
    for i, col in enumerate(columns):
        c.drawString(x_offset + col_widths[i] * i, y_offset, col)
    
    # Rows
    c.setFont("Helvetica", 8)
    y_offset -= row_height
    for row in df.values.tolist():
        for i, val in enumerate(row):
            c.drawString(x_offset + col_widths[i] * i, y_offset, str(val))
        y_offset -= row_height
    
    c.save()
    return output.getvalue()

st.set_page_config(page_title="Analyse des Interventions", page_icon="📊", layout="wide")
st.title("📊 Analyse des interventions des opérateurs")

fichier_principal = st.file_uploader("Choisissez le fichier principal (donnee_Aesma.xlsx)", type="xlsx")

if fichier_principal is not None:
    df_principal = charger_donnees(fichier_principal)
    
    col1, col2 = st.columns([1, 2])
    
    with col1:
        col_prenom_nom = 'Prénom et nom'  # Utilisation directe de la colonne 'Prénom et nom'
        col_date = st.selectbox("Choisissez la colonne de date", df_principal.columns)
        
        # Utilisation de la colonne 'Prénom et nom' pour choisir les opérateurs
        operateurs = df_principal[col_prenom_nom].unique()
        operateurs_selectionnes = st.multiselect("Choisissez un ou plusieurs opérateurs", operateurs)
        
        # Sélection des dates de début et de fin
        date_min = pd.to_datetime(df_principal[col_date], errors='coerce').min().date()
        date_max = pd.to_datetime(df_principal[col_date], errors='coerce').max().date()
        debut_periode = st.date_input("Début de la période", min_value=date_min, max_value=date_max, value=date_min)
        fin_periode = st.date_input("Fin de la période", min_value=debut_periode, max_value=date_max, value=date_max)
        
        # Sélection de la période pour l'affichage
        periodes = ["Jour", "Semaine", "Mois", "Trimestre", "Année"]
        periode_selectionnee = st.selectbox("Choisissez la période d'affichage", periodes)

    if st.button("Analyser"):
        # Conversion des dates
        df_principal[col_date] = pd.to_datetime(df_principal[col_date], errors='coerce')
        df_principal['Jour'] = df_principal[col_date].dt.date
        df_principal['Semaine'] = df_principal[col_date].dt.to_period('W').astype(str)
        df_principal['Mois'] = df_principal[col_date].dt.to_period('M').astype(str)
        df_principal['Trimestre'] = df_principal[col_date].dt.to_period('Q').astype(str)
        df_principal['Année'] = df_principal[col_date].dt.year

        # Filtrage des données en fonction des dates de début et de fin
        df_graph = df_principal[(df_principal[col_date].dt.date >= debut_periode) & 
                                (df_principal[col_date].dt.date <= fin_periode)]

        # Regroupement des données selon la période choisie
        groupby_cols = [col_prenom_nom]
        if periode_selectionnee != "Jour":
            groupby_cols.append(periode_selectionnee)
        
        repetitions_graph = df_graph[df_graph[col_prenom_nom].isin(operateurs_selectionnes)].groupby(groupby_cols).size().reset_index(name='Repetitions')

        with col2:
            # Affichage du graphique avec les valeurs des répétitions
            fig = px.bar(repetitions_graph, x=periode_selectionnee if periode_selectionnee != "Jour" else col_prenom_nom,
                         y='Repetitions', title=f"Nombre de rapports d'intervention (du {debut_periode} au {fin_periode})")
            fig.update_traces(text=repetitions_graph['Repetitions'], textposition='outside')
            st.plotly_chart(fig)
        
        st.subheader("Tableau du nombre des rapports d'interventions par opérateur")
        st.dataframe(repetitions_graph, use_container_width=True)

        # Tirage au sort
        st.subheader("Tirage au sort de deux lignes par opérateur")
        df_filtre = df_principal[(df_principal[col_date].dt.date >= debut_periode) & 
                                 (df_principal[col_date].dt.date <= fin_periode)]
        for operateur in operateurs_selectionnes:
            st.write(f"Tirage pour {operateur}:")
            df_operateur = df_filtre[df_filtre[col_prenom_nom] == operateur]
            lignes_tirees = df_operateur.sample(n=min(2, len(df_operateur)))
            if not lignes_tirees.empty:
                st.dataframe(lignes_tirees, use_container_width=True)
            else:
                st.write("Pas de données disponibles pour cet opérateur dans la période sélectionnée.")
            st.write("---")
        
        # Options pour télécharger les tableaux
        st.subheader("Téléchargement des données")
        # Télécharger le tableau en format Excel
        xlsx_data = convert_df_to_xlsx(repetitions_graph)
        st.download_button(label="Télécharger en format Excel", data=xlsx_data, file_name="repetitions.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # Télécharger le tableau en format PDF
        pdf_data = convert_df_to_pdf(repetitions_graph)
        st.download_button(label="Télécharger en format PDF", data=pdf_data, file_name="repetitions.pdf", mime="application/pdf")

    if st.checkbox("Afficher toutes les données"):
        st.dataframe(df_principal)
