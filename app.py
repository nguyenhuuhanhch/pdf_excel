from flask import Flask, request
import subprocess
import fitz  # PyMuPDF
import requests

app = Flask(__name__)

def download_file_from_pdf(pdf_path, keyword="Excel"):
    # Mở file PDF
    pdf_document = fitz.open(pdf_path)

    # Duyệt qua từng trang
    for page_num in range(pdf_document.page_count):
        page = pdf_document[page_num]
        text_instances = page.search_for(keyword)  # Tìm từ khóa "Excel" trên trang

        # Duyệt qua tất cả các vị trí mà từ khóa xuất hiện
        for inst in text_instances:
            links = page.get_links()
            for link in links:
                if link['uri'] and link['uri'].startswith("https://news.andi.vn/downloadFile.ashx?"):
                    uri = link['uri']
                    print(f"Found link: {uri}")

                    # Tải file về từ link đã tìm được
                    download_excel_from_link(uri)  # Gọi hàm tải file
                    return  # Thoát sau khi tìm được link phù hợp để tránh tải nhiều file

    pdf_document.close()

def download_excel_from_link(url, output_file='downloaded_file.xlsx'):
    try:
        response = requests.get(url, stream=True)
        response.raise_for_status()  # Kiểm tra lỗi HTTP
        
        # Ghi file xuống máy local
        with open(output_file, 'wb') as out_file:
            for chunk in response.iter_content(1024):
                out_file.write(chunk)
        print(f"File downloaded successfully as {output_file}")
    except requests.exceptions.RequestException as e:
        print(f"An error occurred while downloading: {e}")

@app.route('/run-script', methods=['POST'])
def run_script():
    # Đường dẫn tới file được upload
    file_path = request.json.get('file_path')

    # Chạy script xử lý PDF
    download_file_from_pdf(file_path, "Excel")
    
    return {'status': 'Script executed successfully'}

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
