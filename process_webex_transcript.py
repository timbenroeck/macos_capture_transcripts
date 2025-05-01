################################
# Still very much a WIP, I don't have any webex meetings.
################################

import json
import sys
import os

# --- Configuration ---
OUTPUT_BASE_DIR = "processed_transcripts"

# --- Helper Function ---

def _find_webex_table(node):
    """
    Recursively finds the first node with role "AXTable" in the Webex JSON hierarchy.

    Args:
        node (dict or list): The current node or list of nodes to search within.

    Returns:
        dict or None: The found AXTable dictionary, or None if not found.
    """
    if isinstance(node, dict):
        if node.get("role") == "AXTable":
            return node
        for child in node.get("children", []):
            result = _find_webex_table(child)
            if result:
                return result
    elif isinstance(node, list):
        for item in node:
            result = _find_webex_table(item)
            if result:
                return result
    return None

# --- Core Processing Functions ---

def parse_webex_json(data):
    """
    Parses the Webex JSON data structure to extract transcript parts.

    Args:
        data (dict): The loaded JSON data from the Webex dump file.

    Returns:
        list: A list of tuples, where each tuple contains
              (speaker, timestamp, dialogue). Returns empty list if no
              valid data or table is found.
    """
    transcript_parts = []

    table_node = _find_webex_table(data)
    if not table_node:
        print("Warning: Could not find AXTable structure in the Webex JSON.")
        return [] # Return empty list if table not found

    rows = table_node.get("children", [])

    for row in rows:
        if row.get("role") != "AXRow":
            continue
        cells = row.get("children", [])
        if not cells or cells[0].get("role") != "AXCell":
            continue

        cell_children = cells[0].get("children", [])
        speaker = None
        timestamp = None
        dialogue = None # Renamed from 'message' for consistency

        for child in cell_children:
            role = child.get("role")
            value = child.get("value", "") # Use value consistently

            if role == "AXStaticText":
                # Basic check for timestamp format (HH:MM or HH:MM:SS)
                if ":" in value and (value.count(":") == 1 or value.count(":") == 2):
                    timestamp = value.strip()
                else:
                    speaker = value.strip()
            elif role == "AXScrollArea":
                # Look for the text area within the scroll area
                for grandchild in child.get("children", []):
                    if grandchild.get("role") == "AXTextArea":
                        dialogue = grandchild.get("value", "").strip()
                        break # Assume first text area is the message

        # Only add if all parts were found
        if speaker and timestamp and dialogue:
            transcript_parts.append((speaker, timestamp, dialogue))
        # Optional: Add warning if parts are missing for a row?
        # else:
        #     print(f"Warning: Skipping row due to missing data. Found: Speaker='{speaker}', Timestamp='{timestamp}', Dialogue='{dialogue}'")


    return transcript_parts


def format_transcript(transcript_parts):
    """
    Formats the extracted transcript parts into the final text output string.
    (Identical to the Zoom script's format function)

    Args:
        transcript_parts (list): A list of (speaker, timestamp, dialogue) tuples.

    Returns:
        str: The formatted transcript as a single string.
    """
    output_lines = []
    for speaker, timestamp, dialogue in transcript_parts:
        output_lines.append(f"[{speaker}] {timestamp}")
        output_lines.append(f"{dialogue}\n") # Add extra newline for readability
    return "\n".join(output_lines)

# --- Main Function ---

def process_webex_file(input_file_path):
    """
    Main function to process a single Webex transcript JSON file.

    Args:
        input_file_path (str): The path to the input Webex JSON file.
    """
    print(f"Processing Webex transcript file: {input_file_path}")

    if not os.path.isfile(input_file_path):
        print(f"Error: Input file not found at '{input_file_path}'")
        sys.exit(1)

    # Determine output path
    base_name = os.path.basename(input_file_path)           # e.g., 'webex_dump.json'
    file_name_without_ext = os.path.splitext(base_name)[0]  # e.g., 'webex_dump'
    output_file_name = f"{file_name_without_ext}.txt"       # e.g., 'webex_dump.txt'

    # Ensure output directory exists
    try:
        os.makedirs(OUTPUT_BASE_DIR, exist_ok=True)
    except OSError as e:
        print(f"Error creating output directory '{OUTPUT_BASE_DIR}': {e}")
        sys.exit(1)

    output_file_path = os.path.join(OUTPUT_BASE_DIR, output_file_name)
    print(f"Output will be saved to: {output_file_path}")

    # Read and parse the JSON file
    try:
        with open(input_file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except json.JSONDecodeError as e:
        print(f"Error: Failed to decode JSON from '{input_file_path}': {e}")
        sys.exit(1)
    except Exception as e:
        print(f"Error reading file '{input_file_path}': {e}")
        sys.exit(1)

    # Extract transcript parts
    transcript_parts = parse_webex_json(data)
    if not transcript_parts:
        print("Warning: No transcript parts could be extracted.")
        # Write an empty file to indicate processing occurred but found nothing
        try:
            with open(output_file_path, 'w', encoding='utf-8') as out_file:
                out_file.write("")
            print(f"Empty transcript file written to {output_file_path}")
        except IOError as e:
             print(f"Error writing empty output file '{output_file_path}': {e}")
             sys.exit(1)
        return # Exit gracefully after writing empty file

    print(f"Successfully extracted {len(transcript_parts)} transcript segments.")

    # Format the transcript
    formatted_output = format_transcript(transcript_parts)

    # Write the output file
    try:
        with open(output_file_path, 'w', encoding='utf-8') as out_file:
            out_file.write(formatted_output)
        print(f"Successfully converted transcript to {output_file_path}")
    except IOError as e:
        print(f"Error writing output file '{output_file_path}': {e}")
        sys.exit(1)

# --- Script Execution ---

if __name__ == "__main__":
    if len(sys.argv) != 2:
        # Updated usage message
        print("Usage: python process_webex_transcript.py <path_to_webex_dump.json>")
        sys.exit(1)

    input_path = sys.argv[1]
    process_webex_file(input_path) # Call the refactored main function
