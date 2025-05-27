import requests
import os
import pandas as pd
import datetime

download_results = []
#end_point = 'pdf/print'
#document_type = 'pdf'
end_point = 'archive'
document_type = 'zip'


def download_document(file_name, document_id):
    url = f"https://edo.vchasno.ua/api/v2/documents/{document_id}/{end_point}"
    file_name = f"{file_name}.{document_type}"

    headers = {
        "Content-Type": f"application/{document_type}",
        "Authorization": "V0hx_b1XF6dWtqAc9EKToz_AjE3mfXknhMwh"
    }

    try:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            save_path = r"C:\Users\ykoli\Downloads\VchasnoDocs"

            if not os.path.exists(save_path):
                os.makedirs(save_path)
        
            full_save_path = os.path.join(save_path, file_name)

            with open(full_save_path, "wb") as f:
                    f.write(response.content)

            print(f"File successful downloaded as {full_save_path}")

            append_to_excel(file_name, document_id, full_save_path, "Ok")
    
        else:
            error_message = f"Error while downloading file: {response.status_code}"
            print(error_message)
            append_to_excel(file_name, document_id, error_message, "Error")
    except Exception as e:
        error_message = f"Exception occurred: {str(e)}"
        print(error_message)

        append_to_excel(file_name, document_id, error_message, "Error")

result_file_path = r"C:\Users\ykoli\Downloads\result.xlsx"

def append_to_excel(file_name, document_id, full_save_path, status):
    if not os.path.exists(result_file_path):
        df = pd.DataFrame(columns=["File Name", "Document ID", "Full File Path", "Status", "Date"])
        df.to_excel(result_file_path, index=False)

    df = pd.read_excel(result_file_path)
    current_date = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if ((df['Document ID'] == document_id) & (df['Status'] == "Error")).any():
        df.loc[(df['Document ID'] == document_id) & (df['Status'] == "Error"), :] = [file_name, document_id, full_save_path, status, current_date]
    else:
        new_row = pd.DataFrame([[file_name, document_id, full_save_path, status, current_date]], columns=df.columns)
        df = pd.concat([df, new_row], ignore_index=True)
    
    df.to_excel(result_file_path, index=False)

    print(f"Result for {file_name} saved to {result_file_path}")

def process_file(file_path):
    df = pd.read_excel(file_path)

    for index, row in df.iterrows():
        doc_link = row['Посилання на документ']
        document_id = doc_link.split('/')[-1]
        print(document_id)

        file_name = row['NEW_NAME']
        print(file_name)

        download_document(file_name, document_id)

process_file("C:/Users/ykoli/Downloads/Download_Text_File.xlsx")