import requests
from photoshop import Session
from shutil import copyfile
import time
import win32com.client
import json
import os
import dotenv

dotenv.load_dotenv()


# WooCommerce API credentials
api_key = os.getenv("API_KEY")
api_secret = os.getenv("API_SECRET")
store_url = os.getenv("STORE_URL")
desired_fields = ["id", "line_items"]

# Photoshop template folder path
# template_folder = r"C:\Users\Administrator\Pictures"
template_folder = "images"

# Define the API endpoint for orders
orders_endpoint = f"{store_url}orders"


def fetch_orders():
    response = requests.get(orders_endpoint, auth=(api_key, api_secret))
    if response.status_code == 200:
        return response.json()
    else:
        print(f"Error: {response.status_code} - {response.text}")
        return None


def process_order(order):
    product_name = order["line_items"][0]["name"]
    order_id = order["id"]
    original_psd_path = os.path.join(template_folder, f"{product_name}.psd")
    edited_psd_path = os.path.join(
        template_folder, f"{product_name}_order_{order_id}.psd"
    )
    copyfile(original_psd_path, edited_psd_path)

    if os.path.exists(edited_psd_path):
        psApp = win32com.client.Dispatch("Photoshop.Application")
        psApp.Open(edited_psd_path)
        doc = psApp.Application.ActiveDocument
        for meta_data in order["line_items"][0]["meta_data"]:
            key = meta_data["key"]
            value = meta_data["value"]
            try:
                layer = doc.ArtLayers[key]
                layer.TextItem.contents = value
            except Exception as e:
                print(f"Error replacing text in layer '{key}': {e}")

        png_path = edited_psd_path.replace(".psd", ".png")
        options = win32com.client.Dispatch("Photoshop.ExportOptionsSaveForWeb")
        options.Format = 13  # PNG
        options.PNG8 = False
        doc.Export(ExportIn=png_path, ExportAs=2, Options=options)
        doc.Close(2)

        downloadable_url = (
            f"https://designbytext.shop/wp-json/wc/v3/{os.path.basename(png_path)}"
        )
        product_id = create_downloadable_product(
            downloadable_url, f"{product_name}_order_{order_id}"
        )
        if product_id and grant_access_to_product(order_id, product_id):
            print(f"Access granted to order {order_id} for product {product_id}")
            if not mark_order_as_completed(order_id):
                print(f"Failed to mark order {order_id} as completed")
        else:
            print(f"Failed to grant access to order {order_id} or create product")
    else:
        print(f"Copied PSD file not found for product: {product_name}")


def create_downloadable_product(file_path, product_name):
    product_data = {
        "name": product_name,
        "type": "simple",
        "virtual": True,
        "downloadable": True,
        "catalog_visibility": "hidden",
        "downloads": [{"name": product_name, "file": file_path}],
        "download_limit": -1,  # Unlimited downloads
    }
    headers = {"Content-Type": "application/json"}
    response = requests.post(
        f"{store_url}products",
        auth=(api_key, api_secret),
        data=json.dumps(product_data),
        headers=headers,
    )
    if response.status_code == 201:
        return response.json()["id"]
    else:
        print(f"Failed to create product: {response.status_code} - {response.text}")
        return None


def grant_access_to_product(order_id, product_id):
    data = {"line_items": [{"id": product_id}]}
    headers = {"Content-Type": "application/json"}
    response = requests.put(
        f"{store_url}orders/{order_id}",
        auth=(api_key, api_secret),
        data=json.dumps(data),
        headers=headers,
    )
    return response.status_code == 200


def mark_order_as_completed(order_id):
    data = {"status": "completed"}
    headers = {"Content-Type": "application/json"}
    response = requests.put(
        f"{store_url}orders/{order_id}",
        auth=(api_key, api_secret),
        data=json.dumps(data),
        headers=headers,
    )
    return response.status_code == 200


# Poll for orders every 15 seconds
while True:
    orders = fetch_orders()
    if orders:
        for order in orders:
            process_order(order)
    time.sleep(15)
