from __future__ import unicode_literals
import os
import json
import random
import configparser
import pandas as pd
import openpyxl
from urllib import parse
from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import *

app = Flask(__name__)

# LINE Bot 的基本資料
config = configparser.ConfigParser()
config.read("config.ini")

line_bot_api = LineBotApi(config.get("line-bot", "channel_access_token"))
handler = WebhookHandler(config.get("line-bot", "channel_secret"))

# 接收 LINE 的資訊
@app.route("/callback", methods=["POST"])
def callback():
    # get X-Line-Signature header value
    signature = request.headers["X-Line-Signature"]

    # get request body as text
    body = request.get_data(as_text=True)
    app.logger.info("Request body: " + body)

    # create a rich menu
    rich_menu_to_create = RichMenu(
        size=RichMenuSize(width=2500, height=1686),
        selected=True,
        name="圖文選單",
        chat_bar_text="查看更多功能",
        areas=[
            RichMenuArea(
                bounds=RichMenuBounds(x=0, y=0, width=1250, height=843),
                action=MessageAction(label="抽角色", text="抽角色"),
            ),
            RichMenuArea(
                bounds=RichMenuBounds(x=1250, y=0, width=1250, height=843),
                action=MessageAction(label="抽桌布", text="抽桌布"),
            ),
            RichMenuArea(
                bounds=RichMenuBounds(x=0, y=843, width=1250, height=843),
                action=URIAction(
                    label="看動漫", uri="https://www.linetv.tw/drama/11941/eps/1"
                ),
            ),
            RichMenuArea(
                bounds=RichMenuBounds(x=1250, y=843, width=1250, height=843),
                action=MessageAction(label="角色列表", text="角色列表"),
            ),
        ],
    )
    rich_menu_id = line_bot_api.create_rich_menu(rich_menu=rich_menu_to_create)
    print(rich_menu_id)

    with open("menu.png", "rb") as f:
        line_bot_api.set_rich_menu_image(rich_menu_id, "image/png", f)
    line_bot_api.set_default_rich_menu(rich_menu_id)

    # handle webhook body
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)

    return "OK"


# 處理 client 訊息
@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    message = event.message.text
    print("received message:", message)

    # 名字的 dataframe
    name_action_df = pd.read_excel(
        "reply_messages.xlsx", usecols=["name", "action"], engine="openpyxl"
    )
    name_df = name_action_df[name_action_df.action == "intro"]

    # 隨機抽一個角色
    if message == "抽角色":
        r = random.randint(0, len(name_df) - 1)
        name = name_df.iloc[r, 0]
        line_bot_api.reply_message(
            event.reply_token,
            get_reply_message(name, "intro"),
        )

    # 抽桌布
    elif message == "抽桌布":
        wallpaper_df = name_action_df[name_action_df.action == "img"]
        r = random.randint(1, len(wallpaper_df))
        line_bot_api.reply_message(
            event.reply_token,
            get_reply_message("wallpaper" + str(r), "img"),
        )

    # 角色列表
    elif message == "角色列表":
        name_all = {"type": "text", "text": ""}
        for i in range(len(name_df)):
            name_all["text"] += name_df["name"].values[i]
            if i != len(name_df) - 1:
                name_all["text"] += "\n"
        messageObject = getMessageObject(name_all)
        line_bot_api.reply_message(event.reply_token, messageObject)

    # 抽指定角色
    elif message in name_df["name"].values:
        line_bot_api.reply_message(
            event.reply_token,
            get_reply_message(message, "intro"),
        )

    # 關鍵字搜尋
    else:
        for i in range(len(name_df)):
            if message in name_df["name"].values[i]:
                line_bot_api.reply_message(
                    event.reply_token,
                    get_reply_message(name_df["name"].values[i], "intro"),
                )
                break


# 處理 postback 訊息
@handler.add(PostbackEvent)
def handle_postback(event):
    data = event.postback.data
    querystring_dict = dict(parse.parse_qsl(data))
    print("received data:", querystring_dict)
    line_bot_api.reply_message(
        event.reply_token,
        get_reply_message(querystring_dict["name"], querystring_dict["action"]),
    )


# 回傳資料
def get_reply_message(name, action):
    reply_df = pd.read_excel("reply_messages.xlsx", engine="openpyxl")
    name_row = reply_df[reply_df.name == name]
    action_row = name_row[name_row.action == action]

    return_arr = []
    for i in range(5):
        if not pd.isnull(action_row["message" + str(i + 1)].values[0]):
            # string to dict
            jsonObject = json.loads(action_row["message" + str(i + 1)].values[0])
            messageObject = getMessageObject(jsonObject)
            return_arr.append(messageObject)

    print("generate reply messages from", name)
    return return_arr


# 把 JSON 轉成 LINE Bot Object 格式
def getMessageObject(jsonObject):
    message_type = jsonObject.get("type")
    if message_type == "text":
        return TextSendMessage.new_from_json_dict(jsonObject)
    elif message_type == "imagemap":
        return ImagemapSendMessage.new_from_json_dict(jsonObject)
    elif message_type == "template":
        return TemplateSendMessage.new_from_json_dict(jsonObject)
    elif message_type == "image":
        return ImageSendMessage.new_from_json_dict(jsonObject)
    elif message_type == "sticker":
        return StickerSendMessage.new_from_json_dict(jsonObject)
    elif message_type == "audio":
        return AudioSendMessage.new_from_json_dict(jsonObject)
    elif message_type == "location":
        return LocationSendMessage.new_from_json_dict(jsonObject)
    elif message_type == "flex":
        return FlexSendMessage.new_from_json_dict(jsonObject)
    elif message_type == "video":
        return VideoSendMessage.new_from_json_dict(jsonObject)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
