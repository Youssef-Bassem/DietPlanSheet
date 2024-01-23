import gspread
from oauth2client.service_account import ServiceAccountCredentials
from openai import OpenAI
import time
import re
import os

credentials       = ServiceAccountCredentials.from_json_keyfile_name("creds.json" )
client            = OpenAI(api_key=os.environ.get('openAiApiKey'))
gc                = gspread.authorize(credentials)
spreadsheet       = gc.open("DietPlan")
templateWorksheet = spreadsheet.worksheet('Template')
nutriFactsWs      = spreadsheet.worksheet('Saved Nutrition Facts')
dietWorksheet     = spreadsheet.worksheet('Diet')
savedNutriFacts   = nutriFactsWs.get_all_records()
data              = templateWorksheet.get_all_values()
macrosToExtract   = data[0][2:]

# call chatgpt
def callOpenAi(msgContent):
  chat_completion = client.chat.completions.create(
  messages=[
    {
      "role": "user",
      "content": msgContent,
    }
  ],
  model="gpt-3.5-turbo",
)
  return chat_completion.choices[0].message.content

# Summation of all macros in the last row in Diet sheet
def sumMacros():
  try:
    sumRow = ['Sum','']
    for j in range(2,len(data[0])):
      sum = 0
      for i in range(1,len(data)):
        sum += float(data[i][j])
      sumRow.append(sum)
    data.append(sumRow)
  except Exception as e:
    print(e)

########################################### GET MACROS ###########################################
for i in range(1,len(data)):
  print(data[i][1])
  try:
    for j,item in enumerate(savedNutriFacts):
      if data[i][1] == item['Food']:
        for k, macro in enumerate(macrosToExtract):
          if macro in savedNutriFacts[j]:
            data[i][k+2] = float(data[i][0]) * float(savedNutriFacts[j][macro])
          else:
            data[i][k+2] = 0
        break
    else:
      content = callOpenAi(f"What is the nutrition facts of {data[i][0] + ' ' + data[i][1]} and I want every macros for each of the following written in a new line {' '.join(macrosToExtract)}, I want to give me an integer rough number not a range.")
      print(content)
      for j, macro in enumerate(macrosToExtract):
        matchDigit = re.search(f'{macro}:.* (\d+)', content).group(1)
        if matchDigit.isnumeric():
          data[i][j+2] = matchDigit
        else:
          data[i][j+2] = 0
      time.sleep(15)
  except Exception as e:
    print("Couldn't set ", macro, " estimate nutrition fact ",e)


########################################### WRITE DIET PLAN ###########################################
sumMacros()
dietWorksheet.clear()
dietWorksheet.append_rows(data)
dietWorksheet.format(f'A1:Z1'                    , {'textFormat': {'bold': True},"horizontalAlignment": "CENTER"})
dietWorksheet.format(f'A{len(data)}:Z{len(data)}', {'textFormat': {'bold': True},"horizontalAlignment": "CENTER"})


# Mifflin-St.Jeor
# RMR = (9.99 * BM) + (6.25 * H) - (4.92*Age) + 5