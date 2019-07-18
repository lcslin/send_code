## Workplace 活動
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header

df_code_list = pd.read_excel('../0612.xlsx')


def fun_query_code(query_email):
    email_query = str(query_email).strip().lower()
    
    if len(df_code_list[df_code_list.email.str.lower() == email_query].head(1)) != 1:
        return False
    else:
        return df_code_list[df_code_list.email.str.lower() == email_query].iloc[0,2]
    

@app.route('/api/send_code/',defaults={'query_email' :''})
@app.route('/api/send_code/<query_email>')
def fun_send_code(query_email):
    to_user = "Luke/Leia Skywalker"
    to_email = '' # 收件者 email 
    

    code = fun_query_code(query_email)
    if code == False:
        return '煩請確認資料，重新輸入，謝謝。'
    else:
        to_email = str(query_email).lower().strip()
        pre_fill_url = f'https://docs.google.com/forms/d/e/1FAIpQLSf0mOcWRmG61NLXLkRk-j6Dv8kpR5VWT5JmD2AIsxujaTq4cQ/viewform?usp=pp_url&entry.781779439={code}'
        from_ = 'TAITRA WorkPlace <workplace@taitra.org.tw>'
        to_ = f'{to_user} <{to_email}>'

        html = f'''
        <!doctype html>
        <html>
        <head>
          <meta charset='utf-8'>
          <title>HTML mail</title>
        </head>
        <body>
        您好：<p>
        【TAITRA Workplace社團營運競賽】初選投票 - <font color="red"><strong>您的投票編號為  {code} </strong></font><br>
        說明如下：<br><br>

        1.投票時間：2019/6/14(五) 9:00 ~ 2019/6/20(四) 18:00<br><br>
        2.使用您收到通知信中的「投票編號」，點選「投票連結」，進行投票。<br><br>
        3.每位同仁可投6票，只能投一次(重複投票將以最新投票為準)。<br><br>
        4.您可以點選 <a href="{pre_fill_url}"> 這裏 </a> 進行投票，或經以下連結，輸入投票編號後進行投票·<br><br>
        投票連結：<a href="https://forms.gle/oJhDJGdzPQM3xWsP8">
        https://forms.gle/oJhDJGdzPQM3xWsP8</a><br><br>

        數位商務處 敬上
        </body>
        </html>
        '''
        server = smtplib.SMTP_SSL('smtp.taitra.org.tw', 465)
        server.login("", "") #smtp id/passwd

        mime = MIMEText(html, "html", "utf-8")

        mime["Subject"] = '【TAITRA Workplace社團營運競賽】初選投票'
        mime["From"] = from_
        mime["To"] = to_

        msg = mime.as_string()  

        server.sendmail("", to_,  msg) # from email
        

        
        server.quit()
        return '請至您的 Email 信箱收取代碼，謝謝。'


