# macOS Accessibility Export and Transcript Conversion Tools

This guide provides instructions for setting up the necessary Python environment, cloning the project, running the `accessibility_export.py` script to capture accessibility data from macOS applications (like Zoom or Teams in a browser), and using the conversion scripts (`process_zoom_transcript.py`, `process_teams_transcript.py`) to transform the exported JSON data into readable text transcripts.

**System Requirements:**

* **macOS:** These scripts rely heavily on macOS-specific Accessibility APIs (`ApplicationServices` via `pyobjc`) and process management tools (`ps`, `pgrep`). They will **not** work on Windows or Linux.
* **Python 3:** Ensure Python 3.x is installed. These were developed and tested using `Python 3.13.2`.
* **Git:** You'll need Git installed to clone the repository.

## 1. Setup: Clone Repository, Python Virtual Environment, and Dependencies

Setting up a virtual environment ensures that the required packages don't interfere with other Python projects on your system.

### Check Python Installation

Verify you have Python 3 installed:

```bash
python3 --version
```
*(If you don't have Python 3, please install it from [python.org](https://www.python.org/))*

### Clone Repository and Set Up Virtual Environment

1.  **Clone the Project Repository:**
    Open your terminal, navigate to where you want to store the project, and run:
    ```bash
    git clone https://github.com/timbenroeck/macos_capture_transcripts
    cd macos_capture_transcripts
    ```
    This downloads the project code and moves you into the project directory.

2.  **Create the Virtual Environment:**
    While inside the `macos_capture_transcripts` directory, create a virtual environment. This command creates a subdirectory named `.venv` containing a private Python installation.
    ```bash
    python3 -m venv .venv
    ```

3.  **Activate the Virtual Environment:**
    Activating modifies your shell's PATH to prioritize the Python and pip installations within `.venv`.
    ```bash
    source .venv/bin/activate
    ```
    Your command prompt should now start with `(.venv)`, indicating the environment is active.

### Install Required Packages

With the virtual environment activated, install the necessary packages listed in the `requirements.txt` file:

```bash
pip install -r requirements.txt
```
This command reads the `requirements.txt` file included in the repository and installs the specified packages (including `pyobjc-framework-ApplicationServices` and any others you've added).


### Note on Running Scripts

* **Activation Method (Recommended):** Once the virtual environment is activated (`source .venv/bin/activate`), you *should* be able to run the scripts directly:
    ```bash
    python accessibility_export.py
    python process_zoom_transcript.py ...
    ```
* **Explicit Path Method (Fallback):** If, for some reason, activating the environment doesn't correctly resolve the Python interpreter or its packages in your shell, you can run the scripts by explicitly calling the Python interpreter inside the `.venv` directory. You *don't* need to activate the environment first if using this method:
    ```bash
    .venv/bin/python accessibility_export.py
    .venv/bin/python process_zoom_transcript.py ...
    ```

### Deactivating the Virtual Environment

When you're finished working with the scripts, you can deactivate the environment:

```bash
deactivate
```

---

## 2. Generating Accessibility Data (`accessibility_export.py`)

### Purpose

The `accessibility_export.py` script interacts with running macOS applications to capture their **Accessibility UI Tree** information. This is the raw data structure that assistive technologies use to understand application interfaces. The script can target specific applications and elements (like Zoom's transcript window or Teams' live captions) and export this data periodically into JSON files.

**These JSON files are the necessary input for the conversion scripts described in the next section.**

### Accessibility Permissions

This script **requires Accessibility permissions** to inspect UI elements of other applications.

1.  Run the script once. If permissions are not granted, it will print instructions and exit.
    ```bash
    # Make sure environment is active OR use the explicit path
    python accessibility_export.py
    # OR
    .venv/bin/python accessibility_export.py
    ```
2.  Go to: **System Settings > Privacy & Security > Accessibility**.
3.  Find the terminal application you are running the script from (e.g., `Terminal.app`, `iTerm.app`) in the list.
4.  **Enable the checkbox** next to your terminal application. You might need to unlock the settings pane with your password.
5.  Rerun the script.

### Usage

Run the script from your terminal (ensure the target application like Zoom or Teams in a browser is already running with the relevant window/feature active):

```bash
# Make sure environment is active OR use the explicit path
python accessibility_export.py
# OR include the verbose flag for detailed output:
python accessibility_export.py --verbose
```

The script is interactive and will guide you through:

1.  **Selecting the Application Context:** Choose the application and window you want to target (e.g., "Zoom Transcript Window", "Teams in Browser (Chrome/Prisma)", or manual PID/Window selection options).
2.  **Selecting the Serialization Target:** Choose *what* part of the accessibility tree to save (e.g., "Zoom Transcript Table", "Teams Live Captions Group", or the "Full Element Found"). Often, the default selection based on the context is appropriate.
3.  **Setting Export Parameters:**
    * **Depth:** How deep into the accessibility tree to explore (default usually 25).
    * **Interval:** How often (in seconds) to save a new JSON snapshot. Enter `0` for a single, one-time export. A common interval is `30` seconds.

### Output

* The script creates a base directory named `exports/` in the current working directory (where you ran the script).
* Inside `exports/`, it creates a timestamped subdirectory for each session, like `export_YYYY-MM-DD-HH-MM-SS/`.
* Within this timestamped subdirectory, it saves the accessibility data as individual JSON files, each with its own timestamp in the filename (e.g., `export_YYYY-MM-DD-HH-MM-SS.json`).
* **This timestamped subdirectory (e.g., `./exports/export_2025-04-25-10-30-00/`) is what you will use as input for the conversion scripts.**

Press `Ctrl+C` to stop the periodic export loop when you are finished capturing data.

---

## 3. Converting Exports to Text Transcripts

Once you have generated the JSON accessibility export using `accessibility_export.py`, you can use the following scripts to convert them into formatted text files.

### Output Directory Structure

Both conversion scripts will create an output directory named `processed_transcripts` in the **same directory where you run the conversion script**. The converted text files will be placed inside this directory.

* **Output Location:** `./processed_transcripts/`
* **Output Filename:** The name of the output `.txt` file will be based on the input filename (for Zoom) or the input directory name (for Teams).

### `process_zoom_transcript.py`

#### Purpose

This script processes a **single** JSON file exported by `accessibility_export.py` when targeting Zoom's transcript window. Zoom exports usually contain the full transcript history available at the time of export.

#### Input

* A single `.json` file located within one of the timestamped directories created by `accessibility_export.py` (e.g., inside `./exports/export_YYYY-MM-DD-HH-MM-SS/`). You typically only need the *latest* JSON file from a Zoom export session.

#### Usage

```bash
# Make sure environment is active OR use the explicit path
python process_zoom_transcript.py ./exports/export_2025-04-25-10-30-00/export_export_2025-04-25-10-35-00.json
```

Replace the path with the actual path to **one** of your Zoom export JSON files.

#### Output

* A `.txt` file will be created inside the `./processed_transcripts/` directory.
* The filename will match the input JSON file's name (e.g., `./processed_transcripts/export_export_2025-04-25-10-35-00.txt`).
* The format includes the speaker, timestamp, and dialogue:
    ```
    [Speaker Name] HH:MM:SS
    Dialogue text spoken by the speaker at that time.

    [Another Speaker Name] HH:MM:SS
    More dialogue text.

    ```

### `process_teams_transcript.py`

#### Purpose

This script processes **multiple** JSON files within a directory generated by `accessibility_export.py` when targeting Teams (usually in a browser). Teams exports often only contain recent captions, so this script sorts the JSON files chronologically (using the timestamp in the filename) and stitches them together, removing overlaps to create a single, coherent transcript.

#### Input

* The path to the **entire timestamped directory** created by `accessibility_export.py` containing multiple Teams JSON exports (e.g., `./exports/export_YYYY-MM-DD-HH-MM-SS/`).
* The script relies on the `YYYY-MM-DD-HH-MM-SS` timestamp pattern within the JSON filenames inside this directory to sort them correctly.

#### Usage

```bash
# Make sure environment is active OR use the explicit path
python process_teams_transcript.py ./exports/export_2025-04-25-10-45-15/
```

Replace the path with the actual path to the **directory** containing your Teams export JSON files.

#### Output

* A single `.txt` file will be created inside the `./processed_transcripts/` directory.
* The filename will be based on the name of the input directory (e.g., `./processed_transcripts/export_2025-04-25-10-45-15.txt`).
* The format includes the speaker and their combined dialogue, merging consecutive messages:
    ```
    [Speaker Name]
    Dialogue text from the speaker. This might span multiple original JSON file segments if they spoke consecutively.

    [Another Speaker Name]
    More dialogue text.

    ```
* The script prints progress messages while processing the files.

---
