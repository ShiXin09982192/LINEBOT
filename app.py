from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage

app = Flask(__name__)

# 填入你的 Access Token 和 Secret
line_bot_api = LineBotApi('Pe6bx9AYP13n2EW/2LRZjE+9iGYIB9cCJMBZRzLNsUKgXRCHVy9Xu39A76P12PZlGbWKin/J2LCy/MKhv+y/efExrBrcn1/Qd/kMzroF/agvzYYHX5kv9cutg/O2gAylzxZ/Xnj/TdAqo3xm7IvXqwdB04t89/1O/w1cDnyilFU=')
handler = WebhookHandler('787cbcef88855b0b27853dcb7a7b0651')

@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)

    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)

    return 'OK'

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    incoming = event.message.text

    # 簡單關鍵字判斷
    if incoming.lower() == "今天排程":
        reply = "今天要去：XX大樓、YY大廈維修！"
    elif incoming.lower().startswith("完成"):
        reply = "好的，已記錄完成。"
    else:
        reply = f"你說了：{incoming}"

    line_bot_api.reply_message(
        event.reply_token,
        TextSendMessage(text=reply)
    )

if __name__ == "__main__":
    app.run()
