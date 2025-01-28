import pandas as pd
from datetime import datetime, timedelta
import win32com.client as win32
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
import os

# 원본 파일 경로 설정
file_path = 'C:/Users/KMS-LGR01/Downloads/Quanten (stock.quant).xlsx'

# 새로운 파일 이름 생성
save_date = datetime.now().strftime('%d.%m.%Y')
new_file_path = f'C:/Users/KMS-LGR01/Downloads/{save_date}.xlsx'

# 원본 데이터 불러오기
df = pd.read_excel(file_path, sheet_name='Sheet1')

# 1. 'bestatigt' 열 추가
df.insert(0, 'bestatigt', '')

# 2. 날짜 형식 변경 (dd.mm.yyyy)
for col in ['Ablaufdatum', 'Los-/Seriennummer/Mindesthaltbarkeitsdatum']:
    if col in df.columns:
        df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%d.%m.%Y')

# 3. 'Produkt/Interne Referenz'에서 V 또는 v로 시작하는 행 삭제
df = df[~df['Produkt/Interne Referenz'].str.startswith(('V', 'v'), na=False)]

# 4. 'Los-/Seriennummer/Interne Referenz'에서 T로 시작하는 행 삭제
df = df[~df['Los-/Seriennummer/Interne Referenz'].str.startswith('T', na=False)]

# 5. Verfügbare Menge가 0 이하인 값 제거
df = df[df['Verfügbare Menge'] > 0]

# 6. Haltbarkeit unter 3 Monaten 시트용 데이터 필터링
current_date = datetime.now()
three_months_later = current_date + timedelta(days=90)

# 'Ablaufdatum' 다시 날짜 형식으로 변환 후 필터링
df['Ablaufdatum'] = pd.to_datetime(df['Ablaufdatum'], format='%d.%m.%Y', errors='coerce')
haltbarkeit_df = df[(df['Ablaufdatum'] <= three_months_later) & (df['Ablaufdatum'] >= current_date)]

# 필터링 후 가까운 날짜순으로 오름차순 정렬
haltbarkeit_df = haltbarkeit_df.sort_values(by='Ablaufdatum')

# 'Ablaufdatum'을 다시 dd.mm.yyyy 형식으로 변환
df['Ablaufdatum'] = df['Ablaufdatum'].dt.strftime('%d.%m.%Y')
haltbarkeit_df['Ablaufdatum'] = haltbarkeit_df['Ablaufdatum'].dt.strftime('%d.%m.%Y')

# 7. Bestand von 5 oder weniger 시트용 데이터 필터링
bestand_df = df[df['Verfügbare Menge'] <= 5]

# Los-/Seriennummer/Interne Referenz를 기준으로 오름차순 정렬
bestand_df = bestand_df.sort_values(by='Los-/Seriennummer/Interne Referenz', ascending=True)

# 8. 새로운 시트로 저장
with pd.ExcelWriter(new_file_path, engine='openpyxl') as writer:
    # bearbeitet 시트 저장
    df.to_excel(writer, sheet_name='bearbeitet', index=False)
    # Haltbarkeit unter 3 Monaten 시트 저장 (오름차순 정렬된 데이터)
    haltbarkeit_df.to_excel(writer, sheet_name='Haltbarkeit unter 3 Monaten', index=False)
    # Bestand von 5 oder weniger 시트 저장
    bestand_df.to_excel(writer, sheet_name='Bestand von 5 oder weniger', index=False)

print(f"작업이 완료되었습니다. 새로운 파일이 저장되었습니다: {new_file_path}")



"""
# 생성된 파일 경로
save_date = datetime.now().strftime('%d.%m.%Y')
file_path = f'C:/Users/KMS-LGR01/Downloads/{save_date}.xlsx'

# Ionos SMTP 설정
smtp_server = "smtp.ionos.com"
smtp_port = 587  # TLS 사용 (또는 465 사용 가능)
sender_email = ""  # Ionos 이메일 주소
password = ""  # Ionos 이메일 비밀번호

recipient_email = "heyho0929@gmail.com"

# 이메일 메시지 작성
msg = MIMEMultipart()
msg["From"] = sender_email
msg["To"] = recipient_email
msg["Subject"] = f"Die heutige Excel-Datei: {save_date}"
msg.attach(MIMEText("Hallo,\n\nbitte prüfen die angehängte Datei.\n\nVielen Dank.", "plain"))

# 파일 첨부
if os.path.exists(file_path):
    with open(file_path, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(file_path)}")
    msg.attach(part)

# SMTP 서버 연결 및 이메일 전송
try:
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()  # TLS 보안 연결 활성화
    server.login(sender_email, password)
    server.send_message(msg)
    server.quit()
    print(f"이메일이 {recipient_email}에게 성공적으로 전송되었습니다.")
except Exception as e:
    print(f"이메일 전송 중 오류 발생: {e}")
    """
