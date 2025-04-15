import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
import base64

# Set page configuration to wide mode
st.set_page_config(layout="wide")

# Setze die Seitenüberschrift
st.title("Likert-Skala Umfrage-Visualisierer")

# Seitenspalten für Kontrollelemente und Beispieldateien
col1, col2 = st.columns([3, 1])

with col2:
    st.subheader("Beispiel-Excel herunterladen")
    
    # Beispieldaten erstellen
    example_data = {
        'Frage': [
            'Q1: Das Produkt ist einfach zu benutzen',
            'Q2: Die Benutzeroberfläche ist intuitiv',
            'Q3: Der Kundensupport ist hilfreich',
            'Q4: Das Produkt bietet einen guten Wert',
            'Q5: Ich würde dieses Produkt empfehlen',
            'Q6: Updates verbessern das Produkt',
            'Q7: Das Produkt erfüllt meine Bedürfnisse'
        ],
        '1 - Stimme überhaupt nicht zu': [5, 8, 3, 10, 7, 9, 6],
        '2 - Stimme nicht zu': [8, 12, 7, 18, 13, 15, 9],
        '3 - Neutral': [15, 22, 12, 30, 20, 28, 15],
        '4 - Stimme zu': [42, 35, 38, 25, 31, 26, 35],
        '5 - Stimme voll zu': [30, 23, 40, 17, 29, 22, 35],
        'Vorheriger Durchschnitt': [3.42, 3.65, 3.70, 3.27, 3.20, 3.75, 3.50]
    }
    
    example_df = pd.DataFrame(example_data).set_index('Frage')
    
    # Excel buffer erstellen
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        example_df.to_excel(writer, sheet_name='Umfragedaten')
    
    # Download button for example file
    st.download_button(
        label="Beispiel-Excel-Datei herunterladen",
        data=buffer.getvalue(),
        file_name="likert_skala_beispiel.xlsx",
        mime="application/vnd.ms-excel"
    )
    
    st.markdown("""
    ### Dateiformat
    Die Excel-Datei sollte folgendes Format haben:
    - Fragen als Zeilenindex
    - Spalten für jede Likert-Antwort (1-5)
    - Optionale Spalte "Vorheriger Durchschnitt" für Jahresvergleich
    """)

with col1:
    # Dateihochlade-Widget
    uploaded_file = st.file_uploader("Excel-Datei hochladen", type=["xlsx"])
    
    # Checkbox für Vergleich mit Vorjahr
    show_previous = st.checkbox("Vergleich mit Vorjahreswerten anzeigen", value=True)

if uploaded_file:
    df = pd.read_excel(uploaded_file, index_col=0)
    
    # Überprüfe, ob die erforderlichen Spalten vorhanden sind
    likert_columns = [
        '1 - Stimme überhaupt nicht zu', 
        '2 - Stimme nicht zu', 
        '3 - Neutral', 
        '4 - Stimme zu', 
        '5 - Stimme voll zu'
    ]
    
    # Prüfen, ob die Datei deutsche oder englische Spalten hat
    english_columns = [
        '1 - Strongly Disagree', 
        '2 - Disagree', 
        '3 - Neutral', 
        '4 - Agree', 
        '5 - Strongly Agree'
    ]
    
    if all(col in df.columns for col in likert_columns):
        # Nutze deutsche Spalten
        columns_to_use = likert_columns
    elif all(col in df.columns for col in english_columns):
        # Nutze englische Spalten, benenne sie aber um
        df = df.rename(columns=dict(zip(english_columns, likert_columns)))
        columns_to_use = likert_columns
    else:
        st.error("Die Excel-Datei muss die 5 Likert-Skala Antwortspalten enthalten.")
        st.stop()
    
    # Berechne Durchschnittswerte
    df['Durchschnitt'] = [
        sum((i + 1) * count for i, count in enumerate(row)) / sum(row)
        for row in df.iloc[:, :5].values
    ]
    
    # Prüfe, ob Vorjahreswerte vorhanden sind
    has_previous = 'Vorheriger Durchschnitt' in df.columns
    
    # Berechnung der Veränderung, falls vorhanden
    if has_previous and show_previous:
        df['Veränderung'] = df['Durchschnitt'] - df['Vorheriger Durchschnitt']
        # Sortiere nach Veränderung (größte positive Veränderung zuerst)
        df = df.sort_values('Veränderung', ascending=False)
    else:
        # Sortiere nach Durchschnitt
        df = df.sort_values('Durchschnitt')
    
    # Text für Info-Box im Graph basieren auf der Verfügbarkeit von Vorjahreswerten
    info_text = """Fragen sortiert nach Veränderung zum Vorjahr. 
                  ▲ zeigt Verbesserung, ▼ zeigt Verschlechterung.""" if has_previous and show_previous else "Fragen sortiert nach Durchschnittswert."
    
    # Berechne Prozentsätze
    df_percentage = df.iloc[:, 0:5].div(df.iloc[:, 0:5].sum(axis=1), axis=0) * 100
    
    # Definiere Farben für die Balken
    colors = ['#d9534f', '#f0ad4e', '#f5f5f5', '#5cb85c', '#2ca02c']
    
    # Erstelle Plotly-Diagramm
    fig = go.Figure()
    
    # Berechne cumulative x positions für jeden Balken
    for idx, row in df.iterrows():
        # Für jede Kategorie
        for i, col in enumerate(columns_to_use):
            value = df_percentage.loc[idx, col]
            count = row[col]
            
            # Füge Balken hinzu
            fig.add_trace(go.Bar(
                x=[value],
                y=[idx],
                orientation='h',
                name=col,
                text=int(count) if count > 5 else '',
                textposition='inside',
                insidetextanchor='middle',
                marker=dict(color=colors[i]),
                showlegend=True if idx == df.index[0] else False,
                legendgroup=col,
                hoverinfo='text',
                hovertext=f"{col}: {count} ({value:.1f}%)"
            ))
    
    # Für jeden Indikator, füge den Durchschnittswert rechts hinzu
    for idx, row in df.iterrows():
        # Basistextformat für den Durchschnitt
        avg_text = f"Ø: {row['Durchschnitt']:.2f}"
        
        # Wenn Vorjahreswerte vorhanden sind und angezeigt werden sollen
        if has_previous and show_previous:
            change = row['Veränderung']
            change_symbol = "▲" if change > 0 else "▼" if change < 0 else "○"
            change_sign = "+" if change > 0 else ""
            change_color = "green" if change > 0 else "red" if change < 0 else "black"
            
            avg_text = f"Ø: {row['Durchschnitt']:.2f} ({change_symbol} {change_sign}{change:.2f})"
            
            # Füge Vorjahreswert hinzu
            fig.add_annotation(
                x=110,
                y=idx,
                text=f"Vorjahr: {row['Vorheriger Durchschnitt']:.2f}",
                showarrow=False,
                xanchor='left',
                font=dict(
                    size=11,
                    color='gray',
                    family='Arial'
                ),
                align='left'
            )
        
        # Füge Durchschnittswert hinzu (mit Veränderung wenn verfügbar)
        fig.add_annotation(
            x=100,
            y=idx,
            text=avg_text,
            showarrow=False,
            xanchor='left',
            font=dict(
                size=12,
                color='black' if not (has_previous and show_previous) else change_color,
                family='Arial',
                weight='bold'
            ),
            align='left'
        )
    
    # Layout anpassen
    title_text = "Verteilung der Likert-Skala Antworten" + (" mit Vergleich zum Vorjahr" if has_previous and show_previous else "")
    
    fig.update_layout(
        title=title_text,
        font=dict(
            family="Arial",
            size=12
        ),
        barmode='stack',
        xaxis=dict(
            title='Prozentsatz der Antworten',
            range=[0, 120],  # Noch mehr Platz für Vorjahreswerte
            tickvals=[0, 25, 50, 75, 100],
            ticktext=['0%', '25%', '50%', '75%', '100%']
        ),
        yaxis=dict(
            title='Fragen',
            autorange="reversed"
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.2,
            xanchor="center",
            x=0.5
        ),
        height=max(500, 50 * len(df) + 150),  # Dynamische Höhe basierend auf Anzahl der Zeilen
        margin=dict(l=20, r=190 if has_previous and show_previous else 130, t=70, b=100),
        plot_bgcolor='white',
        hoverlabel=dict(
            bgcolor="white",
            font_size=12,
            font_family="Arial"
        ),
        annotations=[
            dict(
                x=0.5,
                y=-0.25,
                xref="paper",
                yref="paper",
                text=info_text,
                showarrow=False,
                font=dict(size=11, color="gray")
            )
        ]
    )
    
    # Tabellarische Darstellung der Daten
    if has_previous and show_previous:
        st.markdown("### Detaillierte Übersicht der Umfrageergebnisse")
        
        # Erstelle eine formatierte Tabelle
        display_df = df[['Durchschnitt', 'Vorheriger Durchschnitt', 'Veränderung']]
        display_df = display_df.round(2)
        
        # Erstelle HTML für farbkodierte Zellen basierend auf Veränderung
        def color_change(val):
            color = 'green' if val > 0 else 'red' if val < 0 else 'black'
            return f'color: {color}'
        
        # Zeige formatierte Tabelle mit Farbkodierung
        st.dataframe(
            display_df.style.applymap(color_change, subset=['Veränderung']),
            use_container_width=True
        )
    
    # Zeige das Diagramm über die volle Seitenbreite
    st.plotly_chart(fig, use_container_width=True)
