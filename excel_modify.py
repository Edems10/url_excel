import concurrent.futures
import pandas as pd
import requests
import PyPDF2
from io import BytesIO
import multiprocessing
import http.cookiejar
from tkinter import  ttk
import math

CORRECT = 'leave as is'
INCORRECT = ''
TIMEOUT = "URL TIMEOUT"
LOCALIZATION = 'remove en-us'
ACCESS_FORBIDDEN = 'access fodbidden'
COMMENTS = 'DUB comment'
URL = 'link'
MICROSOFT_IN_PDF_LINK = 'PDF is Microsoft'

cookie_jar = http.cookiejar.CookieJar()

def check_pdf_url(url):
    try:
        with requests.Session() as session:
            session.cookies = cookie_jar
            response = session.get(url,timeout=10)
            if response.status_code == 200 and 'application/pdf' in response.headers.get('content-type', ''):
                pdf_content = BytesIO(response.content)
                pdf_reader = PyPDF2.PdfReader(pdf_content)
                num_pages = len(pdf_reader.pages)
                if 'microsoft' in url:
                    return 'microsoft'
                return True
            else:
                return False
    except PyPDF2.utils.PdfReadError as e:
        return False
    except requests.exceptions.Timeout:
        return False
    except requests.RequestException as e:
        return False

def check_localization(url):
    for localization in LOCALIZATION_TYPES:
        if localization in url:
            url_without_localization = url.replace(f"/{localization}", "")
            print(f'replaced {localization} in url: {url}')
            try:
                with requests.Session() as session:
                    session.cookies = cookie_jar
                    response = session.head(url_without_localization, allow_redirects=True,timeout=10)
                    if response.status_code == 200:
                        return (True, f'remove {localization}')
                    elif response.status_code == 301:
                        redirected_url = response.headers.get('Location')
                        try:
                            response = session.head(redirected_url, allow_redirects=True,timeout=10)
                            if response.status_code == 200:
                                return (True, f'remove {localization}')
                        except requests.exceptions.Timeout:
                            print('timeout for',url)
                            return (False, TIMEOUT)
                        except requests.RequestException as e:
                            return (False, 'Error') 
                    else:
                        return False
            except requests.RequestException as e:
                return (False, 'Error')
    return (False, 'no localization')

def check_url(url, curr_row, data):
    
    if url.startswith('mailto'):
        data.at[curr_row, COMMENTS] = CORRECT
        return True
    
    if CHECK_PDF:
        if url.lower().endswith('.pdf'):
            pdf_check_result = check_pdf_url(url)
            if pdf_check_result == 'microsoft':
                data.at[curr_row, COMMENTS] = MICROSOFT_IN_PDF_LINK
                return True
            elif pdf_check_result:
                data.at[curr_row, COMMENTS] = CORRECT
                return True
            else:
                data.at[curr_row, COMMENTS] = INCORRECT
                return False
    
    localization_outcome = check_localization(url)
    if localization_outcome[0]:
        data.at[curr_row, COMMENTS] = localization_outcome[1]
        return True
    elif localization_outcome[1] != 'no localization':
        return False
    
    try:
        with requests.Session() as session:
            session.cookies = cookie_jar
            response = session.head(url, allow_redirects=True,timeout=10)
            if response.status_code == 200:
                data.at[curr_row, COMMENTS] = CORRECT
                return True
            elif response.status_code == 301:
                redirected_url = response.headers.get('Location')
                check_url(redirected_url, curr_row, data)
            elif response.status_code == 403:
                data.at[curr_row, COMMENTS] = ACCESS_FORBIDDEN
                return False
            elif response.status_code == 405:
                data.at[curr_row, COMMENTS] = ACCESS_FORBIDDEN
                return False
            else:
                print(response.status_code)
                data.at[curr_row, COMMENTS] = INCORRECT
                return False
    except requests.exceptions.Timeout:
        print('timeout for',url)
        data.at[curr_row, COMMENTS] = TIMEOUT
        return False
    except requests.RequestException as e:
        data.at[curr_row, COMMENTS] = INCORRECT
        return False

def process_url(curr_row, address, data):
    positive_count = 0
    negative_count = 0

    if check_url(address, curr_row, data):
        positive_count += 1
    else:
        negative_count += 1
        print(f"Processing of row {curr_row} failed for address: {address}") 

    return positive_count, negative_count


def process_excel(file_path,cor,incor,forbiden,pdf,local_types,progress_bar:ttk.Progressbar):
    
    global CHECK_PDF
    global CORRECT
    global INCORRECT
    global ACCESS_FORBIDDEN
    global LOCALIZATION_TYPES
    
    FILE_PATH = file_path
    CORRECT =cor
    INCORRECT = incor
    ACCESS_FORBIDDEN = forbiden
    CHECK_PDF =pdf
    LOCALIZATION_TYPES = local_types
    max_cores = multiprocessing.cpu_count()-1
    data = pd.read_excel(FILE_PATH)
    list_addresses = data[URL].tolist()
    progress_bar_length = len(list_addresses)
    updated_data = data.copy()
    total_positive = 0
    total_negative = 0
    curr_progress = 0
    with concurrent.futures.ThreadPoolExecutor(max_workers=max_cores) as executor:
        future_to_row = {
            executor.submit(process_url, curr_row, address, updated_data): curr_row
            for curr_row, address in enumerate(list_addresses)
        }

        for future in concurrent.futures.as_completed(future_to_row):
            curr_row = future_to_row[future]
            try:
                positive_count, negative_count = future.result()
                total_positive += positive_count
                total_negative += negative_count
                curr_progress = (curr_row / progress_bar_length) * 100 if (curr_row / progress_bar_length) * 100 > curr_progress else curr_progress
                progress_bar['value']= curr_progress
                progress_bar.pack()
                print(f"Processed row {curr_row} successfully.")
            except Exception as e:
                print(f"Error processing row {curr_row}: {e}")

    data[COMMENTS] = updated_data[COMMENTS]

    data.to_excel(FILE_PATH, index=False)
    return total_positive, total_negative
