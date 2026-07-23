# 📚 Open-Access Journal Auto-Downloader

An automated Python tool designed to streamline the literature review process by batch-downloading Open-Access academic papers. It reads a list of DOIs from an Excel file, intelligently queries multiple legal Open-Access APIs, and downloads the valid PDF files directly to your local drive. 

Whether you are compiling clinical research on MRI sequences or parsing large amounts of spatial-temporal data literature, this script automates the tedious manual download process.

## ✨ Features

- **Multi-Source Retrieval:** Sequentially queries Unpaywall, OpenAlex, DOAJ, Elsevier Article Retrieval API, and Crossref to find the best available Open-Access PDF.
- **Smart PDF Validation:** Verifies file signatures (`%PDF-`) and sizes to ensure the downloaded file is a valid document, avoiding corrupted files or HTML login pages.
- **Automated Landing Page Scraping:** Crawls journal landing pages and extracts public PDF links hidden in HTML metadata (`citation_pdf_url`, `wkhealth_pdf_url`, etc.).
- **Resilient Progress Tracking:** Automatically updates and saves the download status, source, and final URL back into a tracking Excel file (`Jauto_Status_OpenAccess.xlsx`). If the script stops, it resumes exactly where it left off.
- **Connection Retry Mechanism:** Built-in HTTP adapter with automatic retries for timeouts and bad gateways to handle unstable network conditions gracefully.

## 🚀 Prerequisites

Ensure you have Python 3.8+ installed. The following Python libraries are required:

```bash
pip install pandas requests beautifulsoup4 openpyxl urllib3
```

## ⚙️ Configuration & Setup

To maximize the success rate and respect API rate limits, it is highly recommended to configure the following environment variables:

- `SCHOLAR_EMAIL`: Your email address (used for Unpaywall and OpenAlex polite pools).
- `OPENALEX_API_KEY`: (Optional) Your OpenAlex API key.
- `ELSEVIER_API_KEY`: Your Elsevier API key (required to pull from Elsevier's API).

**Setting Environment Variables (Windows):**
```cmd
set SCHOLAR_EMAIL=your.email@example.com
set ELSEVIER_API_KEY=your_elsevier_key_here
```

**Setting Environment Variables (Linux/Mac):**
```bash
export SCHOLAR_EMAIL="your.email@example.com"
export ELSEVIER_API_KEY="your_elsevier_key_here"
```

## 📂 Usage

1. **Prepare your Input File:**
   Create an Excel file named `Jauto.xlsx` in the same directory as the script. Ensure it contains the following required column headers:
   - `Sub`: A short category or subject name (used for naming the file).
   - `Title`: The title of the paper.
   - `DOI`: The DOI of the paper (e.g., `10.1016/j.nicl.2018.10.011`).

2. **Run the Script:**
   ```bash
   python Jauto.py
   ```

3. **Check the Output:**
   - **Downloaded PDFs:** Found in the `Jurnal_Unduhan` folder.
   - **Download Report:** Open `Jauto_Status_OpenAccess.xlsx` to see the status, source, and notes for every DOI processed.

## 🛠️ How It Works

1. **Read & Normalize:** The script reads `Jauto.xlsx` and normalizes the DOI strings.
2. **Retrieve Candidates:** It requests URL candidates from multiple legal APIs.
3. **Download & Validate:** It attempts to download from the candidates, checking the chunk stream for the `%PDF-` signature.
4. **Log & Save:** The result is saved iteratively to prevent data loss.

## ⚠️ Disclaimer

This script strictly interacts with open-access and legal API sources. It does not integrate with shadow libraries (like Sci-Hub). Always ensure your use of APIs complies with the respective providers' terms of service.

### Dokumentasi / Hasil

![Screenshot 1](Screenshot%202026-07-23%20143859.png)

![Screenshot 2](Screenshot%202026-07-23%20143933.png)

![Screenshot 3](Screenshot%202026-07-23%20144127.png)
