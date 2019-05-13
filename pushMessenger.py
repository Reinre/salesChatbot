#!/venv/bin/python

from jinja2 import Environment, FileSystemLoader, select_autoescape
import json
from linebot import LineBotApi

from linebot.models import FlexSendMessage, TextSendMessage

from linebot.models import BubbleContainer
import datetime
import pytz
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import config
from sales_report import SalesReport

line_bot_api = LineBotApi(config.COnfig.CHANNEL_ACCESS_TOKEN)
scope = ['https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('dummy.json', scope)
client = gspread.authorize(creds)

# Template format definition
# Plus, the template was made on the LINE Flex Message Simulator
template_env = Environment(
    loader=FileSystemLoader('templates'),
    autoescape=select_autoescape(['html', 'xml', 'json'])
)

def push_report(content):
        #The coordinates for monthObjList will always be the same regardless of current date
        monthObjList = ws.acell('G36').value
        objList = ws.acell('G' + str(date.day + 4)).value
        pastList = ws.acell('H' + str(date.day + 4)).value
        # Here we remove the commas from the values for proper integer manipulation
        objList = objList.replace(',','')
        pastList = pastList.replace(',','')
        monthObjList = monthObjList.replace(',','')
        objList = int(objList)
        pastList = int(pastList)
        # MConverting all of the values into integers	
        threepm = int(ws.acell('J' + str(date.day + 4)).value)
        sixpm = int(ws.acell('K' + str(date.day + 4)).value)
        ninepm = int(ws.acell('L' + str(date.day + 4)).value)
        totalSales = ws.acell('I' + str(date.day + 4)).value
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
            changetxt = " error"
        # Whether the objective was reached or not for that day sales
        objectivestatus = ""
        if totalSales >= objective:
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
        date_o = date.strftime("%Y-%m-%d")
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
        line_bot_api.push_message(config.Config.GROUP_ID, msg)

if __name__ == "__main__":
    push_report()