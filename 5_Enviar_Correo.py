"""
Script 5: Envío automático del boletín PDF por correo Gmail
- Lee Boletin_latest.pdf desde 2_Results/
- Envía a lista de destinatarios configurada en GitHub Secrets
- Usa Gmail con contraseña de aplicación (App Password)
"""

import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime

# ---------------------------------------------------------
# Configuración desde GitHub Secrets
# ---------------------------------------------------------
GMAIL_USER    = os.environ.get("GMAIL_USER", "")
GMAIL_PASS    = os.environ.get("GMAIL_PASSWORD", "")
DESTINATARIOS = os.environ.get("DESTINATARIOS", "")

base_dir     = os.path.dirname(os.path.abspath(__file__))
results_dir  = os.path.join(base_dir, "2_Results")
pdf_path     = os.path.join(results_dir, "Boletin_latest.pdf")

# ---------------------------------------------------------
# Verificaciones
# ---------------------------------------------------------
if not GMAIL_USER or not GMAIL_PASS:
    print("⚠️  GMAIL_USER o GMAIL_PASSWORD no configurados. Saltando envío.")
    exit(0)

if not DESTINATARIOS:
    print("⚠️  DESTINATARIOS no configurado. Saltando envío.")
    exit(0)

if not os.path.exists(pdf_path):
    print(f"⚠️  No se encontró el PDF en: {pdf_path}")
    exit(1)

# ---------------------------------------------------------
# Preparar correo
# ---------------------------------------------------------
fecha_hoy     = datetime.now().strftime("%d/%m/%Y")
correlativo   = f"BO{datetime.now().strftime('%d%m%y')}"
lista_destino = [d.strip() for d in DESTINATARIOS.split(",") if d.strip()]
nombre_pdf    = f"Boletin_Sipacate_{correlativo}.pdf"

asunto = f"Boletín de viento Sipacate {correlativo} – {fecha_hoy}"

cuerpo_html = f"""
<html>
<body style="font-family: Arial, sans-serif; font-size: 14px; color: #333;">

  <p>Estimados,</p>

  <p>
    Se adjunta el <strong>Boletín de Condiciones de Viento {correlativo}</strong>
    para el área de Sipacate, correspondiente al pronóstico del día <strong>{fecha_hoy}</strong>.
  </p>

  <p>El documento incluye:</p>
  <ul>
    <li>Gráficas de dirección del viento a 10 m y 100 m</li>
    <li>Gráficas de velocidad del viento a 10 m y 100 m</li>
    <li>Análisis automático de condiciones favorables y no favorables</li>
    <li>Análisis de velocidades por día</li>
  </ul>

  <p style="color: #555; font-size: 12px;">
    Este correo fue generado y enviado automáticamente por el sistema de pronóstico
    del ICC – CeH y SSP. Los datos provienen del modelo ECMWF IFS vía Open-Meteo API.
  </p>

  <p>Atentamente,<br>
  <strong>Sistema Automático de Pronóstico – Sipacate</strong><br>
  ICC: CeH y SSP</p>

</body>
</html>
"""

# ---------------------------------------------------------
# Construir mensaje
# ---------------------------------------------------------
msg            = MIMEMultipart()
msg['From']    = GMAIL_USER
msg['To']      = ", ".join(lista_destino)
msg['Subject'] = asunto

msg.attach(MIMEText(cuerpo_html, 'html'))

# Adjuntar PDF
with open(pdf_path, "rb") as f:
    adjunto = MIMEBase('application', 'pdf')
    adjunto.set_payload(f.read())
    encoders.encode_base64(adjunto)
    adjunto.add_header('Content-Disposition', f'attachment; filename="{nombre_pdf}"')
    msg.attach(adjunto)

# ---------------------------------------------------------
# Enviar por Gmail SSL
# ---------------------------------------------------------
print(f"📧 Enviando boletín a {len(lista_destino)} destinatario(s)...")

try:
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as servidor:
        servidor.login(GMAIL_USER, GMAIL_PASS)
        servidor.sendmail(GMAIL_USER, lista_destino, msg.as_string())

    print(f"✅ Correo enviado exitosamente.")
    print(f"   Asunto    : {asunto}")
    print(f"   Adjunto   : {nombre_pdf}")
    print(f"   Enviado a : {', '.join(lista_destino)}")

except smtplib.SMTPAuthenticationError:
    print("❌ Error de autenticación Gmail.")
    print("   Usa una 'Contraseña de aplicación' de Google, no tu contraseña normal.")
    exit(1)
except Exception as e:
    print(f"❌ Error al enviar correo: {e}")
    exit(1)
