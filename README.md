# Zapier Automation using Webhook

## Overview

This Flask application sets up a webhook endpoint (`/zap-webhook`) to receive data from Zapier, process the received JSON data, and replace placeholders in a `.docx` document with actual values from the received data. The modified document is then saved with the client's name.

## Requirements

- Python 3.7
- Flask
- python-docx
- python-dotenv

## Setup

1. **Install Dependencies:**

   ```bash
   pip install flask python-docx python-dotenv
   ```

2. **Environment Variables:**

   Create a `.env` file in the project directory. This file should contain:

   ```
   DOCUMENT_NAME=Path_to_your_template_document.docx
   ```

3. **Run the Application:**

   ```bash
   python app.py
   ```

   The application will run on ngrok endpoint.

## Endpoint

### `/zap-webhook`

#### Method: POST

#### Request Body

The endpoint expects a JSON object with the following structure:

```json
{
    "first_name": "John",
    "last_name": "Doe",
    "email": "john.doe@example.com",
    "firm_name": "Example Firm",
    "attorney_name": "Jane Smith",
    "amount": "10000",
    "home_phone": "123-456-7890",
    "address": "123 Main St",
    "city": "Anytown",
    "state": "CA",
    "zip": "12345"
}
```

#### Response

- **Success:** 

   ```json
   {
       "status": "success",
       "message": "Data received successfully!"
   }
   ```

- **Failure:**

   ```json
   {
       "status": "failure",
       "message": "No data received"
   }
   ```

## Functions

### `replace_word_in_paragraphs(paragraphs, old_word, new_word)`

Replaces instances of `old_word` with `new_word` in the provided paragraphs.

### `replace_word_in_tables(tables, old_word, new_word)`

Replaces instances of `old_word` with `new_word` in the provided tables.

### `find_data(json_data)`

Extracts and returns a list of personal data based on the received JSON data.

### `replace_data(personal_data)`

Replaces placeholders in the document with the actual values from `personal_data` and saves the modified document. 

## Example

Once the data is received at the `/zap-webhook` endpoint, the placeholders in the specified document (e.g., `<Client_Name>`, `<Client_Email>`) will be replaced with actual values (e.g., "John Doe", "john.doe@example.com"). The modified document will be saved with the client's name as the filename (e.g., `John Doe.docx`).

---

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
