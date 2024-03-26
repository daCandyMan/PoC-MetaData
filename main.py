import os.path
import os
import base64
import urllib.parse
from azure.storage.blob import BlobServiceClient
from azure.search.documents import SearchClient
from azure.core.credentials import AzureKeyCredential
from document_processing import extract_text_from_docx, extract_title_from_docx, extract_version_from_docx, summarize_text
from azure.search.documents import SearchClient
from tempfile import NamedTemporaryFile
from urllib.parse import quote


"""
Blob storage information
"""
connect_str = "..."
container_name = "policies"
blob_service_client = BlobServiceClient.from_connection_string(connect_str)
container_client = blob_service_client.get_container_client(container_name)

"""
Index documents with azure cognitive search
"""
endpoint = "..."
index_name = "policies-index"
key = "..."
search_client = SearchClient(endpoint, index_name, AzureKeyCredential(key))


def clean_string(input_string):
    # Replace smart quotes and any other characters as needed
    return input_string.replace(u"\u2019", "'").replace(u"\u2018", "'")


def process_and_upload_file(file_path):
    file_name = os.path.basename(file_path)
    policy_id = file_name.replace(' ', '_').replace('.docx', '')
    encoded_file_name = quote(file_name)
    blob_client = container_client.get_blob_client(file_name)

    # Check if the blob already exists
    if blob_client.exists():
        print(f"File {file_name} already exists. Skipping upload.")
        return

    # Upload the file
    print(f"{file_name} is being uploaded")
    with open(file_path, "rb") as data:
        blob_client.upload_blob(data)
        print(f"Uploaded {file_name}")

    storage_account_name = "famulus"
    document_url = f"https://{storage_account_name}.blob.core.windows.net/{container_name}/{encoded_file_name}"
    print(f"Document URL: {document_url}")

    # Process the document to extract metadata
    title = clean_string(extract_title_from_docx(file_path).strip())
    version = clean_string(extract_version_from_docx(file_path).strip().replace('\n', ' ').replace('\r', ' '))
    policy_text = clean_string(extract_text_from_docx(file_path).strip())
    summary = clean_string(summarize_text(policy_text).strip())

    # After confirming the blob is uploaded, then retrieve the URL
    try:
        document_url = blob_client.url
        print(f"Document URL: {document_url}")
    except Exception as e:
        print(f"Failed to retrieve URL for blob {file_name}: {e}")

    # Set the metadata on the blob
    metadata = {
        "title": title,
        "version": version,
        "summary": summary
    }
    blob_client.set_blob_metadata(metadata)


    # Index the document in Azure Cognitive Search (if needed)
    document = {
        "PolicyID": policy_id,  # Unique identifier for the policy
        "Title": title,  # Extracted title from the document
        "Content": policy_text,  # Full content of the document or a summary
        "version": version,  # Extracted version from the document
        "DocumentURL": document_url
    }
    try:
        search_client.upload_documents(documents=[document])
        print(f"Document {file_name} indexed successfully.")
    except Exception as e:
        print(f"Failed to index document {file_name}: {e}")

def upload_files_from_folder(folder_path):
    """
    Uploads all files from a specified folder to Azure Blob Storage.
    """
    print(f"Files in the directory {folder_path}: {os.listdir(folder_path)}")
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)
        print(f"Current file: {file_path}")
        if os.path.isfile(file_path):
            print(f"Processing file: {file_path}")
            process_and_upload_file(file_path)
        else:
            print(f"Skipping {file_path}, not a file.")

try:
    container_client = blob_service_client.get_container_client(container_name)
    blob_list = container_client.list_blobs()
    print("List of blobs:")
    for blob in blob_list:
        print(blob.name)
    print("Connection successful")
except Exception as e:
    print(f"An error occued: {e}")


# Main execution
if __name__ == "__main__":
    directory_path = "documents"
    upload_files_from_folder(directory_path)