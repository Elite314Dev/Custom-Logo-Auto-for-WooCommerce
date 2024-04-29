# Custom Logo Automation Script

This script is designed to automate the process of creating, managing, and delivering custom logos via a WooCommerce-powered website. It integrates directly with Adobe Photoshop to personalize logos based on customer orders and updates the WooCommerce orders to provide customers with downloadable access to the finished logos.

## Key Features

`Order Fetching:` Automatically fetches new orders from the WooCommerce store using the API.

`Logo Customization:` Uses Adobe Photoshop to customize logos based on customer specifications provided in the order details.

`File Management:` Generates and manages PSD and PNG files, ensuring each customer receives a personalized product.

`Product Creation:` Automatically creates a hidden, downloadable product in WooCommerce for each customized logo.

`Order Update:` Updates the customer's order to grant access to the newly created downloadable logo and marks the order as completed.

## Workflow

`Fetch Orders:` The script continuously polls for new orders from the WooCommerce API every 15 seconds.

`Process Orders: `For each new order, the script:
Retrieves the base logo template from a designated folder.
Personalizes the logo in Photoshop based on the order's metadata (e.g., customer's chosen text).
Saves the personalized logo as a PNG file.

`Manage WooCommerce Products:`
Creates a new hidden product with the generated logo file as a downloadable item.
Grants the ordering customer access to this downloadable product.
Marks the order as completed once access is granted.

### Security and Configuration

Uses environment variables to securely manage API keys and store URLs, preventing sensitive data from being hardcoded into the script.
Utilizes robust error handling to ensure smooth operation and provides clear error messages for troubleshooting.

### System Requirements

Adobe Photoshop must be installed on the server running the script.
The server must have internet access to interact with the WooCommerce API and handle file uploads/downloads.

### Installation and Setup

Install necessary Python packages: requests, win32com.client, python-dotenv.

Place the script in a secure directory on the server.
Configure the .env file with the necessary API credentials and store URL.

Ensure that the template PSD files are correctly named and placed in the designated folder.

### Running the Script

The script is designed to run continuously. It should be started in a background process or as a service on the server to ensure uninterrupted operation.
