import os, re, tempfile
from flask import Flask, request
from linebot import LineBotApi, WebhookHandler
from linebot.models import TextSendMessage
from docx import Document
from docx2pdf import convert
import boto3
from datetime import datetime, timedelta

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
    # 簡易解析：抓「業主」、「地址」與品項與金額
    lines = text.splitlines()
    data = {'owner': '', 'address': '', 'items': []}
    for l in lines:
        if l.startswith('業主：'): data['owner'] = l.split('：',1)[1].strip()
        elif l.startswith('地址：'): data['address'] = l.split('：',1)[1].strip()
        m = re.match(r'([A-Za-z0-9]+)[\.\:\s]+(.+?)[\：\s]+([\d,]+)', l)
        if m:
            code, desc, amt = m.groups()
            amt = int(amt.replace(',',''))
            data['items'].append((desc.strip(), amt))
    # 計算合計與稅
    total = sum(amt for _,amt in data['items'])
    tax = int(total * 0.05)
    data.update({'total': total, 'tax': tax, 'grand_total': total+tax})
    return data

def fill_docx(data, out_docx_path):
    doc = Document('template.docx')
    # 替換書籤或特定表格位置
    # 假設第一段落是標題，表格在第二個
    table = doc.tables[0]
    # 清空表格中動態列，再填入
    for row in table.rows[1:]:
        for cell in row.cells: cell.text = ''
    for i, (desc, amt) in enumerate(data['items'], start=1):
        row = table.add_row().cells
        row[0].text = str(i)
        row[1].text = desc
        row[2].text = '1'
        row[3].text = '式'
        row[4].text = f'{amt:,}'
    # 寫合計、稅金、總額到文檔尾
    doc.add_paragraph(f'合計：{data["total"]:,}')
    doc.add_paragraph(f'營業稅（5%）：{data["tax"]:,}')
    doc.add_paragraph(f'總計：{data["grand_total"]:,}')
    doc.save(out_docx_path)

def upload_and_get_url(file_path, key):
    s3.upload_file(file_path, S3_BUCKET, key, ExtraArgs={'ACL':'public-read'})
    # 預簽名連結也可用
    url = s3.generate_presigned_url(
        'get_object',
        Params={'Bucket': S3_BUCKET, 'Key': key},
        ExpiresIn=3600
    )
    return url

@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)
    handler.handle(body, signature)
    return 'OK'

@handler.add('message')
def handle_message(event):
    user_id = event.source.user_id
    text = event.message.text
    data = parse_text(text)
    # 臨時檔案
    with tempfile.TemporaryDirectory() as tmp:
        docx_path = f'{tmp}/报价单_{datetime.now().strftime("%Y%m%d%H%M%S")}.docx'
        pdf_path = docx_path.replace('.docx','.pdf')
        fill_docx(data, docx_path)
        convert(docx_path, pdf_path)
        key_docx = os.path.basename(docx_path)
        key_pdf = os.path.basename(pdf_path)
        url_docx = upload_and_get_url(docx_path, key_docx)
        url_pdf = upload_and_get_url(pdf_path, key_pdf)
    # 傳回連結
    msg = (f'報價單已生成：\n'
           f'📄 PDF：{url_pdf}\n'
           f'📝 Word：{url_docx}')
    line_bot_api.reply_message(event.reply_token, TextSendMessage(text=msg))

if __name__ == "__main__":
    app.run(port=8000)
