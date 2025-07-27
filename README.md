# NMMS Attendance Downloader

This project contains a Python script to automatically download muster rolls and attendance data from the NREGA website for a specific region in Karnataka, India. It scrapes the website, navigates through several pages to find the correct data, and then generates multiple formatted Excel reports, including one with embedded photos of the worksites.

## Features

-   Navigates the NREGA website to find attendance data for a specific date and Panchayath.
-   Allows selection of all muster rolls or specific ones based on work codes.
-   Downloads attendance data and associated worksite photos.
-   Generates four different Excel reports:
    1.  **`nmr_{panchayath}_{date}.xlsx`**: A detailed report with attendance data and embedded photos. Cells are color-coded to indicate gender (Male/Female) and attendance status (Present/Absent).
    2.  **`nmr_images_{panchayath}_{date}.xlsx`**: A report containing only the muster roll numbers and the large-format worksite photos.
    3.  **`verification_format_{panchayath}_{date}.xlsx`**: A pre-formatted sheet for official verification purposes.
    4.  **`nmr_raw_{panchayath}_{date}.xlsx`**: A raw data dump of all attendance records in a simple tabular format, also with color-coding for status and gender.
-   Handles potential website structure changes and missing data gracefully to prevent crashes.

## Requirements

-   Python 3.x
-   The packages listed in `requirements.txt`.

## How to Run

1.  **Set up a virtual environment (recommended):**
    ```bash
    python -m venv venv
    source venv/bin/activate  # On Windows, use `venv\Scripts\activate`
    ```

2.  **Install the required packages:**
    ```bash
    pip install -r requirements.txt
    ```

3.  **Run the main script:**
    ```bash
    python attend_web.py
    ```

4.  **Follow the on-screen prompts:**
    -   The script will first display a list of available attendance dates. You will be prompted to **enter a date** from this list.
    -   Next, you will be asked to **enter the Panchayath name**.
    -   Finally, the script will show the available work codes for that Panchayath and ask you to choose between downloading all muster rolls (`all`) or only those for specific works (`work`). If you choose `work`, you will be prompted to enter the work codes.

## Output

After the script finishes, it will save four `.xlsx` files in the project directory, named according to the Panchayath and date you selected. For example:

-   `nmr_my_panchayath_18_07_2025.xlsx`
-   `nmr_images_my_panchayath_18_07_2025.xlsx`
-   `verification_format_my_panchayath_18_07_2025.xlsx`
-   `nmr_raw_my_panchayath_18_07_2025.xlsx`
