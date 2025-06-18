from flask import Flask, render_template, request
import requests
import pandas as pd
from io import BytesIO
import plotly.graph_objs as go
import plotly.offline as pyo
from datetime import datetime

app = Flask(__name__)

meses_es = [
    "01_ENERO", "02_FEBRERO", "03_MARZO", "04_ABRIL", "05_MAYO", "06_JUNIO",
    "07_JULIO", "08_AGOSTO", "09_SEPTIEMBRE", "10_OCTUBRE", "11_NOVIEMBRE", "12_DICIEMBRE"
]

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        fecha_str = request.form.get("fecha")
        try:
            fecha = datetime.strptime(fecha_str, "%Y-%m-%d")
        except:
            return render_template("index.html", error="Formato de fecha inválido. Use YYYY-MM-DD.")

        anio = fecha.year
        mes = meses_es[fecha.month - 1]
        dia = f"{fecha.day:02d}"
        fecha_excel = fecha.strftime("%Y%m%d")
        nombre_archivo = f"Anexo1_Despacho_{fecha_excel}.xlsx"
        url = f"https://www.coes.org.pe/portal/browser/download?url=Operaci%C3%B3n%2FPrograma%20de%20Operaci%C3%B3n%2FPrograma%20Diario%2F{anio}%2F{mes}%2FD%C3%ADa%20{dia}%2F{nombre_archivo}"

        try:
            response = requests.get(url)
            if response.status_code != 200 or not response.content:
                return render_template("index.html", error="⚠️ No se encontró el archivo Excel para esa fecha. Elija otra fecha.")
            
            archivo_bytes = BytesIO(response.content)
            df = pd.read_excel(archivo_bytes, header=5)

        except Exception:
            return render_template("index.html", error="⚠️ No se pudo procesar el archivo Excel. Intente con otra fecha.")

        columnas = ["Día Hora", "MANTARO", "RESTITUCION", "TUMBES MAK 1", "TUMBES MAK 2", "MALPASO"]
        df_filtrado = df[columnas].iloc[0:48]

        fig = go.Figure()
        for col in columnas[1:]:
            fig.add_trace(go.Scatter(
                x=df_filtrado["Día Hora"],
                y=df_filtrado[col],
                mode='lines+markers',
                name=col
            ))

        fig.update_layout(
            title=f"Programa Diario de Operación ({fecha_str})",
            xaxis_title="Hora",
            yaxis_title="MW",
            template="plotly_white",
            legend_title="Central",
            xaxis=dict(tickangle=270),
            yaxis=dict(range=[0,700], showgrid=True)
        )

        graph_html = pyo.plot(fig, output_type='div', include_plotlyjs=True)
        return render_template("index.html", graph_html=graph_html)

    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)
