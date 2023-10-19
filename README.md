```markdown
# PYEX

## Excel Data API

This is a Django API for creating, reading, updating, and deleting data in Excel files.

### Prerequisites

Before using this API, ensure you have the following:

- Python and Django installed.
- Required Python libraries: openpyxl, pandas.
- Excel file (main.xlsx) to store data.

### API Endpoints

#### Create Data

Create a new sheet and add data to it.

**Request:**

- Method: POST
- URL: `/api/excel/sheet_name`
- Body: JSON data with 'json_objects' containing an array of objects.

Example:
```json
{
  "json_objects": [
    {
      "name": "John Doe",
      "amount": 100,
      "pending": false
    },
    {
      "name": "Alice Smith",
      "amount": 200,
      "pending": true
    }
  ]
}
```

**Response:**

The API responds with a success message:

```json
{
  "message": "Data created successfully"
}
```

#### Read Data

Retrieve data from an existing sheet.

**Request:**

- Method: GET
- URL: `/api/excel/sheet_name`

Example: GET `/api/excel/sales`

**Response:**

The API responds with the data from the sheet:

```json
{
  "data": [
    {
      "sr_no": 1,
      "name": "John Doe",
      "amount": 100,
      "pending": 50
    },
    {
      "sr_no": 2,
      "name": "Alice Smith",
      "amount": 200,
      "pending": 30
    }
  ]
}
```

#### Update Data

Update data in an existing sheet.

**Request:**

- Method: PUT
- URL: `/api/excel/sheet_name`
- Body: JSON data with 'json_objects' containing an array of objects. Each object should have 'sr_no' for identifying the row.

Example:
```json
{
  "json_objects": [
    {
      "sr_no": 1,
      "name": "Updated Name",
      "amount": 300,
      "pending": 33
    }
  ]
}
```

**Response:**

The API responds with a success message:

```json
{
  "message": "Data updated successfully"
}
```

#### Delete Data

Clear data in an existing sheet.

**Request:**

- Method: DELETE
- URL: `/api/excel/sheet_name`

Example:
```json
{
   "sr_no" : 1
}
```

**Response:**

The API responds with a success message:

```json
{
  "message": "Data deleted successfully"
}
```

### File Path

Replace the `file_path` in views.py with the actual path to your Excel file.

### Usage

1. Run the Django development server.
2. Use an API client or a tool like `curl` to make requests to the API endpoints.
3. Monitor responses and handle errors.

### Error Handling

- HTTP 400: Bad Request - Missing or invalid data.
- HTTP 404: Not Found - The requested sheet does not exist.
- HTTP 500: Internal Server Error - Unexpected server issues.

### License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

### Acknowledgments

- [Django](https://www.djangoproject.com/)
- [openpyxl](https://openpyxl.readthedocs.io/en/stable/)
- [pandas](https://pandas.pydata.org/)