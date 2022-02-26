import os
import sys
import urllib.request
import pandas as pd

def get_papago(source, source_lang, target_lang):
    url = "https://papago.naver.com/apis/dictionary/search"
    request = urllib.request.Request(url)
    encText = urllib.parse.quote(source)
    data = f"source={source_lang}&target={target_lang}&text={encText}&locale=en"
    try:
        with urllib.request.urlopen(request, data=data.encode("utf-8")) as response:
            rescode = response.getcode()
            if (rescode == 200):
                response_body = response.read().decode('utf-8')
                print(response_body)
                return response_body
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
                        if not xlsx.loc[i, config['INDEX_AFTER']]:
                            continue
                        data_before = xlsx.loc[i, config['INDEX_BEFORE']]
                        response = get_papago(data_before, 
                                config["SOURCE_LANG"], config['TARGET_LANG'])
                        if response is None:
                            print("Operation Interrupted due to Server Error")
                            xlsx.to_excel(f"{config['FILE_NAME'][:-5]}_translated.xlsx")
                            exit()
                        else:
                            data_after = response
                            xlsx.loc[i, config['INDEX_AFTER']] = data_after
            except ValueError:
                print("Error: No such sheet")
                exit()

        xlsx.to_excel(f"{config['FILE_NAME'][:-5]}_translated.xlsx")

except FileNotFoundError:
    print("Error: No such file")
    exit()
