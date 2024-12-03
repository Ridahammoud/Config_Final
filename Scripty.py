import pandas as pd
import streamlit as st
from datetime import datetime, timedelta
import plotly.express as px
from io import BytesIO
import xlsxwriter
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

@st.cache_data
def charger_donnees(fichier):
    try:
        return pd.read_excel(fichier)
    except Exception as e:
        st.error(f"Erreur lors du chargement du fichier : {str(e)}")
        return None

def convert_df_to_xlsx(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

def generate_pdf(df, filename="tableau.pdf"):
    c = canvas.Canvas(filename, pagesize=letter)
    width, height = letter
    c.drawString(30, height - 40, "Tableau des répétitions des opérateurs")
    y_position = height - 60
    for i, row in df.iterrows():
        text = f"{row['Prénom et nom']} : {row['Repetitions']}"
        c.drawString(30, y_position, text)
        y_position -= 20
    c.save()

st.set_page_config(page_title="Analyse des Interventions", page_icon="📊", layout="wide")
st.title("📊 Analyse des interventions des opérateurs")

fichier_principal = st.file_uploader("Choisissez le fichier principal (donnee_Aesma.xlsx)", type="xlsx")

if fichier_principal is not None:
    df_principal = charger_donnees(fichier_principal)
    
    col1, col2 = st.columns([1, 2])
    
    with col1:
        col_prenom_nom = df_principal.columns[4]
        col_date = st.selectbox("Choisissez la colonne de date", df_principal.columns)
        
        operateurs = df_principal[col_prenom_nom].unique()
        operateurs.append("Total")  # Ajout de l'option "Total"
        operateurs_selectionnes = st.multiselect("Choisissez un ou plusieurs opérateurs", operateurs)
        
        # Si "Total" est sélectionné, on inclut tous les opérateurs
        if "Total" in operateurs_selectionnes:
            operateurs_selectionnes = df_principal[col_prenom_nom].unique().tolist()  # Utilise tous les opérateurs disponibles
        
        periodes = ["Jour", "Semaine", "Mois", "Trimestre", "Année"]
        periode_selectionnee = st.selectbox("Choisissez une période", periodes)
        
        df_principal[col_date] = pd.to_datetime(df_principal[col_date], errors='coerce')
        
        date_min = df_principal[col_date].min()
        date_max = df_principal[col_date].max()

        if pd.isna(date_min) or pd.isna(date_max):
            st.warning("Certaines dates dans le fichier sont invalides. Elles ont été ignorées.")
            date_min = date_max = None
        
        debut_periode = st.date_input("Début de la période", min_value=date_min, max_value=date_max, value=date_min)
        fin_periode = st.date_input("Fin de la période", min_value=debut_periode, max_value=date_max, value=date_max)
    
    if st.button("Analyser"):
        df_principal = df_principal.dropna(subset=[col_date])

        df_principal['Jour'] = df_principal[col_date].dt.date
        df_principal['Semaine'] = df_principal[col_date].dt.to_period('W').astype(str)
        df_principal['Mois'] = df_principal[col_date].dt.to_period('M').astype(str)
        df_principal['Trimestre'] = df_principal[col_date].dt.to_period('Q').astype(str)
        df_principal['Année'] = df_principal[col_date].dt.year

        df_graph = df_principal[(df_principal[col_date].dt.date >= debut_periode) & (df_principal[col_date].dt.date <= fin_periode)]

        groupby_cols = [col_prenom_nom]
        if periode_selectionnee != "Total":
            groupby_cols.append(periode_selectionnee)
        
        repetitions_graph = df_graph[df_graph[col_prenom_nom].isin(operateurs_selectionnes)].groupby(groupby_cols).size().reset_index(name='Repetitions')
        repetitions_tableau = df_principal[df_principal[col_prenom_nom].isin(operateurs_selectionnes)].groupby(groupby_cols).size().reset_index(name='Repetitions')
        
        with col2:
            fig = px.bar(repetitions_graph, 
                         x=periode_selectionnee if periode_selectionnee != "Jour" else col_prenom_nom,
                         y='Repetitions',
                         barmode='group',
                         color=col_prenom_nom,
                         title=f"Nombre de rapports d'intervention (du {debut_periode} au {fin_periode})")
            fig.update_traces(texttemplate='%{y}', textposition='outside')
            st.plotly_chart(fig)

        st.subheader("Moyennes mensuelles")
        df_mensuel = df_principal[df_principal[col_prenom_nom].isin(operateurs_selectionnes)].groupby([col_prenom_nom, df_principal[col_date].dt.to_period('M')]).size().reset_index(name='Repetitions')
        moyennes_mensuelles = df_mensuel.groupby(col_prenom_nom)['Repetitions'].mean().reset_index()
        moyennes_mensuelles = moyennes_mensuelles.sort_values('Repetitions', ascending=False)
        
        moyenne_totale = moyennes_mensuelles['Repetitions'].mean()
        st.write(f"Moyenne mensuelle totale : {moyenne_totale:.2f}")

        col4, col5, col6 = st.columns(3)
        with col4:
            st.write("Moyennes mensuelles par opérateur :")
            st.dataframe(moyennes_mensuelles)
        with col5:
            st.write("Top 5 des moyennes mensuelles maximales :")
            st.dataframe(moyennes_mensuelles.head())
        with col6:
            st.write("5 moyennes mensuelles minimales :")
            st.dataframe(moyennes_mensuelles.tail())

        st.subheader(f"Tableau du nombre des rapports d'intervention par {periode_selectionnee.lower()} (toutes les dates)")
        colonnes_affichage = [col_prenom_nom, periode_selectionnee, 'Repetitions'] if periode_selectionnee != "Total" else [col_prenom_nom, 'Repetitions']
        tableau_affichage = repetitions_tableau[colonnes_affichage]
        st.dataframe(tableau_affichage, use_container_width=True)

        st.subheader("Tirage au sort de deux lignes par opérateur")
        df_filtre = df_principal[(df_principal[col_date].dt.date >= debut_periode) & (df_principal[col_date].dt.date <= fin_periode)]
        for operateur in operateurs_selectionnes:
            st.write(f"Tirage pour {operateur}:")
            df_operateur = df_filtre[df_filtre[col_prenom_nom] == operateur]
            lignes_tirees = df_operateur.sample(n=min(2, len(df_operateur)))
            if not lignes_tirees.empty:
                lignes_tirees['Photo'] = lignes_tirees['Photo'].apply(lambda x: f'<img src="{x}" width="100"/>')
                lignes_tirees['Photo 2'] = lignes_tirees['Photo 2'].apply(lambda x: f'<img src="{x}" width="100"/>')
                st.markdown(lignes_tirees.to_html(escape=False), unsafe_allow_html=True)
            else:
                st.write("Pas de données disponibles pour cet opérateur dans la période sélectionnée.")
            st.write("---")

        st.subheader("Télécharger le tableau des rapports d'interventions")
        xlsx_data = convert_df_to_xlsx(repetitions_tableau)
        st.download_button(label="Télécharger en XLSX", data=xlsx_data, file_name="NombredesRapports.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.subheader("Télécharger le tableau des rapports d'interventions en PDF")
        pdf_filename = "repetitions.pdf"
        generate_pdf(repetitions_tableau, pdf_filename)
        with open(pdf_filename, "rb") as f:
            st.download_button(label="Télécharger en PDF", data=f, file_name=pdf_filename, mime="application/pdf")

    if st.checkbox("Afficher toutes les données"):
        st.dataframe(df_principal)
