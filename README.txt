━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  PNL DIARIO — App Streamlit
  Punto Casa de Bolsa
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

────────────────────────────────────────
 CONTENIDO
────────────────────────────────────────
  app.py             La aplicación Streamlit
  requirements.txt   Dependencias para Streamlit Cloud
  README.txt         Este archivo

────────────────────────────────────────
 PROBARLA EN LOCAL (opcional)
────────────────────────────────────────
  pip install -r requirements.txt
  streamlit run app.py

  Se abre en http://localhost:8501

────────────────────────────────────────
 DESPLEGAR EN STREAMLIT CLOUD (gratis)
────────────────────────────────────────

 PASO 1 — Subir a GitHub
  Crea un repo nuevo (puede ser privado) y sube
  estos dos archivos a la raíz:
     app.py
     requirements.txt

  Desde terminal:
     git init
     git add app.py requirements.txt
     git commit -m "PNL diario app"
     git branch -M main
     git remote add origin https://github.com/TU_USUARIO/pnl-punto.git
     git push -u origin main

 PASO 2 — Conectar Streamlit Cloud
  1. Entra a  https://share.streamlit.io
  2. Inicia sesión con tu cuenta de GitHub
  3. Click en "New app"
  4. Selecciona:
       Repository : TU_USUARIO/pnl-punto
       Branch     : main
       Main file  : app.py
  5. Click en "Deploy"

  En ~2 minutos tendrás una URL tipo:
     https://pnl-punto.streamlit.app

 PASO 3 — Usarla
  Abre la URL, arrastra el COEBMV101.xlsx,
  y descarga el PNL ya formateado. Listo.

────────────────────────────────────────
 PRIVACIDAD
────────────────────────────────────────
  • Si el repo es PÚBLICO, cualquiera con la URL
    puede usar la app (pero NO ve tus archivos:
    cada quien sube el suyo y se procesa en memoria,
    nada se guarda en el servidor).

  • Para restringir el acceso solo a ti:
    En Streamlit Cloud → Settings → Sharing →
    "Only specific people can view this app"
    y agrega tu correo. (Requiere plan o repo privado
    según la política vigente de Streamlit.)

  • Los archivos que subes NO se almacenan: se
    procesan en RAM y se descartan al cerrar la sesión.

────────────────────────────────────────
 LÓGICA DE CÁLCULO
────────────────────────────────────────
  • Long  (C → V): compra primero, vende al cierre
  • Short (V → C): vende primero, compra para cerrar
  • PNL = Importe Venta − Importe Compra (universal)
  • Posiciones incompletas (solo C o solo V) se
    marcan y se excluyen del PNL Total
  • Detecta el sentido por el orden de aparición
    de cada operación en el archivo

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
