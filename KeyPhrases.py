import httplib, urllib, base64, json
from openpyxl import Workbook
from openpyxl import load_workbook

# This program reads controls and use Azure Key phrases API to extract key phrases
# Name of the file with cloud security standard
cfname = 'NIST 800-53-controls.xlsx'
# Name of the new file with keyphrases
ofile = 'US_NIST_800-53-r4.xlsx'
# File with subscription key in json format
keyfile = 'Credential.json'
with open(keyfile,'r') as k:
    key = json.load(k)

wb = load_workbook(filename = cfname)
ws = wb['Sheet1']
headers = {
    'Content-Type': 'application/json',
    'Ocp-Apim-Subscription-Key': key['key'],
}
params = urllib.urlencode({

})

# Assume column B is id, F is text and I is for key phrases
row = 1
while not(ws['B'+str(row)].value is None and ws['F'+str(row)].value is None):
    row = row + 1
    id = ws['B'+str(row)].value
    text = ws['F'+str(row)].value
    body = json.dumps({
        "documents" : [
            {
                "id": id,
                "text": text,
            }
        ]
    })
    try:
        conn = httplib.HTTPSConnection('westus.api.cognitive.microsoft.com')
        conn.request("POST", "/text/analytics/v2.0/keyPhrases?%s" % params, body, headers)
        response = conn.getresponse()
        data = json.loads(response.read())
        documents = data['documents']
        print(str(row)+": "+documents[0]['id'])
        ws['I'+str(row)].value = ', '.join(documents[0]['keyPhrases'])
        #print(', '.join(documents[0]['keyPhrases']))
        conn.close()
    except Exception as e:
        print(str(row)+": "+e.message+"\n"),response.status, response.reason, response.msg

# Save final output file
wb.save(ofile)