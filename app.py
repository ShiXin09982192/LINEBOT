import os
import re
import tempfile
from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.models import MessageEvent, TextMessage, TextSendMessage
from docx import Document
from docx2pdf import convert
import boto3
from datetime import datetime

# 環境變數設定
LINE_CHANNEL_TOKEN = os.getenv('LINE_CHANNEL_TOKEN')
LINE_CHANNEL_SECRET = os.getenv('LINE_CHANNEL_SECRET')
AWS_ACCESS_KEY = os.getenv('AWS_ACCESS_KEY')
AWS_SECRET_KEY = os.getenv('AWS_SECRET_KEY')
S3_BUCKET = os.getenv('S3_BUCKET')

line_bot_api = LineBotApi(LINE_CHANNEL_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

s3 = boto3.client(
    's3',
    aws_access_key_id=AWS_ACCESS_KEY,
    aws_secret_access_key=AWS_SECRET_KEY
)

app = Flask(__name__)

def parse_text(text):
    lines = text.splitlines()
    data = {'owner': '', 'address': '', 'items': []}
    for l in lines:
        if l.startswith('業主：'):
            data['owner'] = l.split('：',1)[1].strip()
        elif l.startswith('地址：'):
            data['address'] = l.split('：',1)[1].strip()
        m = re.match(r'([A-Za-z0-9]+)[\.\:\s]+(.+?)[\：\s]+([\d,]+)', l)
        if m:
            code, desc, amt = m.groups()
            amt = int(amt.replace(',', ''))
            data['items'].append((desc.strip(), amt))
    total = sum(amt for _, amt in data['items'])
    tax = int(total * 0.05)
    data.update({'total': total, 'tax': tax, 'grand_total': total + tax})
    return data

def fill_docx(data, out_docx_path):
    doc = Document('template.docx')
    table = doc.tables[0]
    for row in table.rows[1:]:
        for cell in row.cells:
            cell.text = ''
    for i, (desc, amt) in enumerate(data['items'], start=1):
        row = table.add_row().cells
        row[0].text = str(i)
        row[1].text = desc
        row[2].text = '1'
        row[3].text = '式'
        row[4].text = f'{amt:,}'
    doc.add_paragraph(f'合計：{data["total"]:,}')
    doc.add_paragraph(f'營業稅（5%）：{data["tax"]:,}')
    doc.add_paragraph(f'總計：{data["grand_total"]:,}')
    doc.save(out_docx_path)

def upload_and_get_url(file_path, key):
    s3.upload_file(file_path, S3_BUCKET, key, ExtraArgs={'ACL': 'public-read'})
    url = s3.generate_presigned_url(
        'get_object',
        Params={'Bucket': S3_BUCKET, 'Key': key},
        ExpiresIn=3600
    )
    return url

@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers.get('X-Line-Signature')
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except Exception as e:
        abort(400)
    return 'OK'

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    text = event.message.text
    data = parse_text(text)
    with tempfile.TemporaryDirectory() as tmp:
        docx_path = os.path.join(tmp, f'报价单_{datetime.now().strftime("%Y%m%d%H%M%S")}.docx')
        pdf_path = docx_path.replace('.docx', '.pdf')
        fill_docx(data, docx_path)
        convert(docx_path, pdf_path)
        key_docx = os.path.basename(docx_path)
        key_pdf = os.path.basename(pdf_path)
        url_docx = upload_and_get_url(docx_path, key_docx)
        url_pdf = upload_and_get_url(pdf_path, key_pdf)
    reply_text = f'報價單已生成：\n📄 PDF：{url_pdf}\n📝 Word：{url_docx}'
    line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply_text))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", 8000)))
