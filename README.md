# ShipScan v9

ShipScan v9 is a Flask-based web application that automates the process of scanning shipping-related emails, extracting shipment information, and generating Excel reports for easy tracking.

## Features

- Scan emails from an IMAP email account.
- Extract shipment-related information.
- Generate Excel reports automatically.
- Simple web interface built with Flask.
- Ready for deployment on Render.

---

## Project Structure

```
ShipScan/
│
├── app.py
├── requirements.txt
├── render.yaml
├── templates/
│   └── index.html
└── README.md
```

---

## Requirements

- Python 3.10 or later
- pip

---

## Installation

Clone the repository:

```bash
git clone https://github.com/your-username/ShipScan.git
cd ShipScan
```

Install dependencies:

```bash
pip install -r requirements.txt
```

---

## Run Locally

Start the application:

```bash
python app.py
```

Open your browser and visit:

```
http://localhost:5000
```

---

## Deploying to Render

### Build Command

```bash
pip install -r requirements.txt
```

### Start Command

```bash
gunicorn app:app
```

Render automatically detects the `PORT` environment variable, so no additional configuration is required.

---

## Files

| File | Description |
|------|-------------|
| app.py | Main Flask application |
| requirements.txt | Python dependencies |
| render.yaml | Render deployment configuration |
| templates/index.html | Web interface |

---

## Dependencies

- Flask
- Gunicorn
- IMAPClient
- OpenPyXL
- PyPDF

---

## Notes

- This application uses temporary storage while running on Render.
- Any files generated during execution are not permanently stored after the service restarts.
- For long-term file storage, consider integrating AWS S3, Google Cloud Storage, or another cloud storage provider.

---

## License

This project is intended for educational and internal use.
# ship_scan_v9
