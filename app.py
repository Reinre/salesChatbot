from flask import Flask, request, abort
from jinja2 import Environment, FileSystemLoader, select_autoescape
import json
from linebot import (
    LineBotApi, WebhookHandler
)
from linebot.exceptions import (
    InvalidSignatureError
)
from linebot.models import (
    MessageEvent, TextMessage, TextSendMessage, FlexSendMessage
)

from linebot.models import BubbleContainer

from datetime import datetime, timedelta
import pytz
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from config import Config
from sales_report import SalesReport
app = Flask(__name__)

# Channel Access Token
line_bot_api = LineBotApi(Config.CHANNEL_ACCESS_TOKEN)
# Channel Secret
handler = WebhookHandler(Config.CHANNEL_SECRET)
scope = ['https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
client = gspread.authorize(creds)
# Template format definition
# Plus, the template was made on the LINE Flex Message Simulator
template_env = Environment(
    loader=FileSystemLoader('templates'),
    autoescape=select_autoescape(['html', 'xml', 'json'])
)

@app.route("/callback", methods=['POST'])
def callback():
    # get X-Line-Signature header value
    signature = request.headers['X-Line-Signature']
    # get request body as text
    body = request.get_data(as_text=True)
    app.logger.info("Request body: " + body)
    # handle webhook body
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return 'OK'

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    print(event.source.user_id)
    #print(event.source.group_id)
    message = event.message.text
    date = datetime.now(tz=pytz.timezone('Japan'))
    ss = client.open(Config.SPREADSHEET_NAME)
    ws = ss.get_worksheet(date.month - 1)
    content = message.split()
    if content[0] == "/sales":
        if len(content) == 2:
            inputt = content[1]
            if inputt.isdigit():
                if date.hour >= 15 and date.hour < 17:
                    ws.update_acell('J' + str(date.day + 4), inputt)
                    response = "3 PM Sales updated."
                elif date.hour >= 18 and date.hour < 20:
                    ws.update_acell('K' + str(date.day + 4), inputt)
                    response = "6 PM Sales updated."
                elif date.hour >= 21 and date.hour < 23:
                    ws.update_acell('L' + str(date.day + 4), inputt)
                    response = "9 PM Sales updated"
                else:
                    response = "Input window currently closed. Try again when a new window opens."
            else: 
                response = "Sorry, please use only numbers after the command."
        else:
            response = "Sorry, you are missing the amount parameter"
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=response))
    elif content[0] == "/final":
        if len(content) == 2:
            inputt = message.split()[1]
            if inputt.isdigit():
                if date.hour >= 0 and date.hour < 1:
                    ws.update_acell('I' + str((date.day - 1) + 4), inputt)
                    response = "Yesterday final sales updated."
                elif date.hour >= 23:
                    ws.update_acell('I' + str(date.day + 4), inputt)
                    response = "Today's final sales updated."
                else: 
                    response = "Input window currently closed. Try again when a new window opens."
            else:
                response = "Sorry, please use only numbers after the command."
        else:
            response = "Sorry, you are missing the amount parameter"
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=response))
    elif content[0] == "/help":
        response = "Bot usage: \n/sales <current sales>: To update de database at 3pm, 6pm or 9pm\n/final <total sales>: To update the total amout within a day\n/help: To print this message"
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=response))
    elif content[0] == "/report":
        #The coordinates for monthObjList will always be the same regardless of current date
        monthObjList = ws.acell('G36').value
        objList = ws.acell('G' + str(date.day + 3)).value
        pastList = ws.acell('H' + str(date.day + 3)).value
        # Here we remove the commas from the values for proper integer manipulation
        objList = objList.replace(',','')
        pastList = pastList.replace(',','')
        monthObjList = monthObjList.replace(',','')
        objList = int(objList)
        pastList = int(pastList)
        # MConverting all of the values into integers   
        threepm = int(ws.acell('J' + str(date.day + 3)).value)
        sixpm = int(ws.acell('K' + str(date.day + 3)).value)
        ninepm = int(ws.acell('L' + str(date.day + 3)).value)
        totalSales = ws.acell('I' + str(date.day + 3)).value
        totalSales = totalSales.replace(',','')
        totalSales = int(totalSales)
        monthlyObj = int(monthObjList)
        lastYearSales = int(pastList)
        objective = int(objList)
        # Some math
        change = int(totalSales/lastYearSales*100-100)
        changetxt = ""

        # Increases or decreases in profit comparing to last year 
        if change > 1.00:
            changetxt = "+"
        elif change < 1.00:
            changetxt = ""
        else:
            changetxt = ""
        # Whether the objective was reached or not for that day sales
        objectivestatus = ""
        if totalSales > objective:
            objectivestatus = "REACHED" 
        elif totalSales < objective:
            objectivestatus = "NOT REACHED"
        else:
            objectivestatus = " error"  
        # Here we append today's sales to the month dictionary to work with the variables below
        current = int(ws.acell('I36').value)
        left = int(monthlyObj) - int(current)
        percentage = int(current / monthlyObj * 100)
        # Class output variables
        objectivestatus_o = str(objectivestatus)    
        date_o = (date - timedelta(days=1)).strftime("%Y-%m-%d")
        lastYearSales_o = str("{:,}".format(lastYearSales)) + "円"
        objective_o = str("{:,}".format(objective)) + "円"
        totalSales_o = str("{:,}".format(totalSales)) + "円"
        threepm_o = str("{:,}".format(threepm)) + "円"
        sixpm_o = str("{:,}".format(sixpm)) + "円"
        ninepm_o = str("{:,}".format(ninepm)) + "円"
        changetxt_o = str(changetxt)
        change_o = str((changetxt) + (change) + "%")
        monthlyObj_o = str("{:,}".format(monthlyObj)) +  "円"
        current_o = str("{:,}".format(current)) + "円"
        left_o = str("{:,}".format(left)) + "円"
        percentage_o = str(percentage) + "%"
        report = SalesReport(objectivestatus_o, '#1DB446', date_o, 
            lastYearSales_o, totalSales_o, objective_o, threepm_o, 
            sixpm_o, ninepm_o, change_o, monthlyObj_o, current_o, left_o, percentage_o)
        template = template_env.get_template('sales-report.json')
        data = template.render(dict(data=report))
        # Send the structured message. Note, the template needs to be parsed to 
        # an Flex MessageUI element, in this case a BubbleContainer
        msg = FlexSendMessage(alt_text = "Sales Report", contents=BubbleContainer.new_from_json_dict(json.loads(data)))
        line_bot_api.reply_message(event.reply_token, msg)

    elif content[0] == "/now":
        if date.hour >= 15 and date.hour < 17:
            
            objList = ws.acell('G' + str(date.day + 4)).value
            pastList = ws.acell('H' + str(date.day + 4)).value
            # Here we remove the commas from the values for proper integer manipulation
            objList = objList.replace(',','')
            pastList = pastList.replace(',','')
            
            objList = int(objList)
            pastList = int(pastList)

            # MConverting all of the values into integers   
            threepm = int(ws.acell('J' + str(date.day + 4)).value)
        
    
            lastYearSales = int(pastList)
            objective = int(objList)

     
            # Class output variables
            date_o = (date - timedelta(days=1)).strftime("%Y-%m-%d")
            lastYearSales_o = str("{:,}".format(lastYearSales)) + "円"
            objective_o = str("{:,}".format(objective)) + "円"
            threepm_o = str("{:,}".format(threepm)) + "円"
    
            
            report = SalesReport3('#1DB446', date_o, 
                lastYearSales_o, objective_o, threepm_o)
            template = template_env.get_template('sales-report3.json')
            data = template.render(dict(data=report))
            # Send the structured message. Note, the template needs to be parsed to 
            # an Flex MessageUI element, in this case a BubbleContainer
            msg = FlexSendMessage(alt_text = "Sales Report 3 PM", contents=BubbleContainer.new_from_json_dict(json.loads(data)))
            line_bot_api.reply_message(event.reply_token, msg)

        elif date.hour >= 18 and date.hour < 20:

            objList = ws.acell('G' + str(date.day + 4)).value
            pastList = ws.acell('H' + str(date.day + 4)).value
            # Here we remove the commas from the values for proper integer manipulation
            objList = objList.replace(',','')
            pastList = pastList.replace(',','')
            
            objList = int(objList)
            pastList = int(pastList)

            # MConverting all of the values into integers   
            threepm = int(ws.acell('J' + str(date.day + 4)).value)
            sixpm = int(ws.acell('K' + str(date.day + 4)).value)
        
    
            lastYearSales = int(pastList)
            objective = int(objList)

     
            # Class output variables
            date_o = (date - timedelta(days=1)).strftime("%Y-%m-%d")
            lastYearSales_o = str("{:,}".format(lastYearSales)) + "円"
            objective_o = str("{:,}".format(objective)) + "円"
            threepm_o = str("{:,}".format(threepm)) + "円"
            sixpm_o = str("{:,}".format(sixpm)) + "円"
    
            
            report = SalesReport6('#1DB446', date_o, 
                lastYearSales_o, objective_o, threepm_o, 
                sixpm_o)
            template = template_env.get_template('sales-report6.json')
            data = template.render(dict(data=report))
            # Send the structured message. Note, the template needs to be parsed to 
            # an Flex MessageUI element, in this case a BubbleContainer
            msg = FlexSendMessage(alt_text = "Sales Report 6 PM", contents=BubbleContainer.new_from_json_dict(json.loads(data)))
            line_bot_api.reply_message(event.reply_token, msg)

        elif date.hour >= 21 and date.hour < 23:


            objList = ws.acell('G' + str(date.day + 4)).value
            pastList = ws.acell('H' + str(date.day + 4)).value
            # Here we remove the commas from the values for proper integer manipulation
            objList = objList.replace(',','')
            pastList = pastList.replace(',','')
            
            objList = int(objList)
            pastList = int(pastList)

            # MConverting all of the values into integers   
            threepm = int(ws.acell('J' + str(date.day + 4)).value)
            sixpm = int(ws.acell('K' + str(date.day + 4)).value)
            ninepm = int(ws.acell('L' + str(date.day + 4)).value)
        
    
            lastYearSales = int(pastList)
            objective = int(objList)

     
            # Class output variables
            date_o = (date - timedelta(days=1)).strftime("%Y-%m-%d")
            lastYearSales_o = str("{:,}".format(lastYearSales)) + "円"
            objective_o = str("{:,}".format(objective)) + "円"
            threepm_o = str("{:,}".format(threepm)) + "円"
            sixpm_o = str("{:,}".format(sixpm)) + "円"
            ninepm_o = str("{:,}".format(ninepm)) + "円"
    
            
            report = SalesReport9('#1DB446', date_o, 
                lastYearSales_o, objective_o, threepm_o, 
                sixpm_o, ninepm_o)
            template = template_env.get_template('sales-report9.json')
            data = template.render(dict(data=report))
            # Send the structured message. Note, the template needs to be parsed to 
            # an Flex MessageUI element, in this case a BubbleContainer
            msg = FlexSendMessage(alt_text = "Sales Report 9 PM", contents=BubbleContainer.new_from_json_dict(json.loads(data)))
            line_bot_api.reply_message(event.reply_token, msg)






    
    
import os
if __name__ == "__main__":
    port = int(Config.PORT)
    app.run(host='0.0.0.0', port=port)