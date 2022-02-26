import os
import sys
import urllib.request
import json
import pandas as pd

def init_request(client_id, client_secret):
    url = "https://openapi.naver.com/v1/papago/n2mt"
    request = urllib.request.Request(url)
    request.add_header("X-Naver-Client-Id", client_id)
    request.add_header("X-Naver-Client-Secret", client_secret)
    return request

def get_papago(request, source, source_lang, target_lang):
    encText = urllib.parse.quote(source)
    data = f"source={source_lang}&target={target_lang}&text=" + encText
    try:
        with urllib.request.urlopen(request, data=data.encode("utf-8")) as response:
            rescode = response.getcode()
            if (rescode == 200):
                response_body = response.read()
                res = json.loads(response_body.decode('utf-8'))
                return res['message']['result']['translatedText']
            else:
                print("Error Code:" + rescode)
                return None
    except urllib.request.HTTPError as e:
        print(e)
        return None

config = {}
with open("config.txt", "r", encoding = "utf-8") as f:
    for line in f.readlines():
        line = line.rstrip()
        key, data = line.split(':')
        if key == "SHEET_NAME":
            config[key] = data.split(',')
        else:
            config[key] = data

try:
    config_required = [
                "FILE_NAME",
                "SHEET_NAME",
                "INDEX_BEFORE",
                "INDEX_AFTER",
                "CLIENT_ID",
                "CLIENT_SECRET",
                "SOURCE_LANG",
                "TARGET_LANG"
            ]
    for c in config_required:
        config[c]
except KeyError:
    print("Error: No Required Configuration Value")
    exit()

if not config['FILE_NAME'].endswith(".xlsx"):
    print("Error: inappropriate excel file name")
    exit()

request = init_request(config['CLIENT_ID'], config['CLIENT_SECRET'])

try:
    with pd.ExcelFile(config['FILE_NAME']) as reader:
        for s in config["SHEET_NAME"]:
            try:
                xlsx = pd.read_excel(reader, sheet_name=s)
                if config['INDEX_BEFORE'] not in xlsx.columns:
                    print("Error: INDEX_BEFORE value not in file")
                    exit()
                elif config['INDEX_AFTER'] not in xlsx.columns:
                    print("Error: INDEX_AFTER value not in file")
                    exit()
                else:
                    for i in xlsx.index:
                        if not pd.isnull(xlsx.loc[i, config['INDEX_AFTER']]):
                            continue
                        data_before = xlsx.loc[i, config['INDEX_BEFORE']]
                        data_after = get_papago(request, data_before, 
                                config["SOURCE_LANG"], config['TARGET_LANG'])
                        if data_after is None:
                            print("Operation Interrupted due to Server Error")
                            xlsx.to_excel(f"{config['FILE_NAME'][:-5]}_translated.xlsx")
                            exit()
                        else:
                            xlsx.loc[i, config['INDEX_AFTER']] = data_after
            except ValueError:
                print("Error: No such sheet")
                exit()

        xlsx.to_excel(f"{config['FILE_NAME'][:-5]}_translated.xlsx", index=False)

except FileNotFoundError:
    print("Error: No such file")
    exit()
