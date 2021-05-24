import requests
from datetime import datetime


myToken = "xoxb-2102757467876-2096814164755-AmrODakCtoHvQXm7UYAmkEwT"

# post_message(myToken, "#stock", "hjun-bot")

def post_message(token, channel, text):
    response = requests.post("https://slack.com/api/chat.postMessage",
                             headers={"Authorization": "Bearer " + token},
                             data={"channel": channel, "text": text}
                             )
    print(response)





def dbgout(message):
    """인자로 받은 문자열을 파이썬 셸과 슬랙으로 동시에 출력한다."""
    print(datetime.now().strftime('[%m/%d %H:%M:%S]'), message)
    strbuf = datetime.now().strftime('[%m/%d %H:%M:%S] ') + message
    post_message(myToken, "#stock", strbuf)
