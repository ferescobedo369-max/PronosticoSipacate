"""
Script 5: Envío automático del boletín por correo Gmail
- Lee Boletin_latest.docx desde 2_Results/
- Envía a la lista de destinatarios configurada en GitHub Secrets
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
# Configuración desde variables de entorno (GitHub Secrets)
# ---------------------------------------------------------
GMAIL_USER    = os.environ.get("GMAIL_USER", "")
GMAIL_PASS    = os.environ.get("GMAIL_PASSWORD", "")
DESTINATARIOS = os.environ.get("DESTINATARIOS", "")  # emails separados por coma

# Rutas
base_dir      = os.path.dirname(os.path.abspath(__file__))
results_dir   = os.path.join(base_dir, "2_Results")
boletin_path  = os.path.join(results_dir, "Boletin_latest.docx")

# ---------------------------------------------------------
# Verificaciones
# ---------------------------------------------------------
if not GMAIL_USER or not GMAIL_PASS:
    print("⚠️  GMAIL_USER o GMAIL_PASSWORD no configurados. Saltando envío de correo.")
    exit(0)

if not DESTINATARIOS:
    print("⚠️  DESTINATARIOS no configurado. Saltando envío de correo.")
    exit(0)

if not os.path.exists(boletin_path):
    print(f"⚠️  No se encontró el boletín en: {boletin_path}")
    exit(1)

# ---------------------------------------------------------
# Preparar correo
# ---------------------------------------------------------
fecha_hoy    = datetime.now().strftime("%d/%m/%Y")
correlativo  = f"BO{datetime.now().strftime('%d%m%y')}"
lista_destino = [d.strip() for d in DESTINATARIOS.split(",") if d.strip()]

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
msg = MIMEMultipart()
msg['From']    = GMAIL_USER
msg['To']      = ", ".join(lista_destino)
msg['Subject'] = asunto

msg.attach(MIMEText(cuerpo_html, 'html'))

# Adjuntar el boletín Word
with open(boletin_path, "rb") as f:
    adjunto = MIMEBase('application', 'vnd.openxmlformats-officedocument.wordprocessingml.document')
    adjunto.set_payload(f.read())
    encoders.encode_base64(adjunto)
    nombre_adjunto = f"Boletin_Sipacate_{correlativo}.docx"
    adjunto.add_header('Content-Disposition', f'attachment; filename="{nombre_adjunto}"')
    msg.attach(adjunto)

# ---------------------------------------------------------
# Enviar
# ---------------------------------------------------------
print(f"📧 Enviando boletín a {len(lista_destino)} destinatario(s)...")
print(f"   Destinatarios: {', '.join(lista_destino)}")

try:
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as servidor:
        servidor.login(GMAIL_USER, GMAIL_PASS)
        servidor.sendmail(GMAIL_USER, lista_destino, msg.as_string())

    print(f"✅ Correo enviado exitosamente.")
    print(f"   Asunto : {asunto}")
    print(f"   Adjunto: {nombre_adjunto}")

except smtplib.SMTPAuthenticationError:
    print("❌ Error de autenticación Gmail.")
    print("   Verifica que GMAIL_USER y GMAIL_PASSWORD estén correctos.")
    print("   Recuerda usar una 'Contraseña de aplicación', no tu contraseña normal.")
    exit(1)
except Exception as e:
    print(f"❌ Error al enviar correo: {e}")
    exit(1)
