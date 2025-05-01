import json
import sys
import os

# --- Configuration ---
OUTPUT_BASE_DIR = "processed_transcripts"

# --- Core Processing Functions ---

def parse_zoom_json(data):
    """
    Parses the Zoom JSON data structure to extract transcript parts.

    Args:
        data (dict): The loaded JSON data from the Zoom export file.

    Returns:
        list: A list of tuples, where each tuple contains
              (speaker, timestamp, dialogue_text). Returns empty list if no
              valid data is found.
    """
    transcript_parts = []
    current_speaker = None
    current_timestamp = None
    current_text_buffer = []

    for row in data.get('children', []):
        cells = row.get('children', [])
        if not cells:
            continue  # Skip rows without cells

        cell_contents = cells[0].get('children', [])
        if not cell_contents:
            continue # Skip cells without content

        # Extract text values and check for speaker image marker
        cell_values = [item.get('value', '') for item in cell_contents if item.get('role') == 'AXTextArea']
        has_image = any(item.get('role') == 'AXImage' for item in cell_contents)

        if has_image:
            # Flush previous speaker block if any
            if current_speaker and current_text_buffer and current_timestamp:
                dialogue = ' '.join(current_text_buffer).strip()
                if dialogue:
                    transcript_parts.append((current_speaker, current_timestamp, dialogue))
                current_text_buffer = [] # Reset buffer
                current_timestamp = None # Reset timestamp

            # Start new speaker block
            current_speaker = cell_values[0].strip() if cell_values else "Unknown Speaker"

        elif cell_values and current_speaker: # Process timestamp and text lines
            # Check if the first value looks like a timestamp (HH:MM:SS)
            if len(cell_values) >= 2 and cell_values[0].count(':') == 2:
                # Flush previous text for the same speaker if timestamp changes
                if current_text_buffer and current_timestamp:
                    dialogue = ' '.join(current_text_buffer).strip()
                    if dialogue:
                        transcript_parts.append((current_speaker, current_timestamp, dialogue))
                    current_text_buffer = [] # Reset buffer

                # Assign new timestamp and start new text line
                current_timestamp = cell_values[0].strip()
                current_text_buffer.append(cell_values[1].strip())
            else:
                 # Append text lines to the current buffer (continuation)
                current_text_buffer.extend([val.strip() for val in cell_values])

    # Flush any remaining text after the loop
    if current_speaker and current_timestamp and current_text_buffer:
        dialogue = ' '.join(current_text_buffer).strip()
        if dialogue:
            transcript_parts.append((current_speaker, current_timestamp, dialogue))

    return transcript_parts


def format_transcript(transcript_parts):
    """
    Formats the extracted transcript parts into the final text output string.

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

def process_zoom_file(input_file_path):
    """
    Main function to process a single Zoom transcript JSON file.

    Args:
        input_file_path (str): The path to the input Zoom JSON file.
    """
    print(f"Processing Zoom transcript file: {input_file_path}")

    if not os.path.isfile(input_file_path):
        print(f"Error: Input file not found at '{input_file_path}'")
        sys.exit(1)

    # Determine output path
    base_name = os.path.basename(input_file_path)           # e.g., 'zoom_meeting_abc.json'
    file_name_without_ext = os.path.splitext(base_name)[0]  # e.g., 'zoom_meeting_abc'
    output_file_name = f"{file_name_without_ext}.txt"       # e.g., 'zoom_meeting_abc.txt'

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
    transcript_parts = parse_zoom_json(data)
    if not transcript_parts:
        print("Warning: No transcript parts could be extracted.")
        # Write an empty file to indicate processing occurred but found nothing
        with open(output_file_path, 'w', encoding='utf-8') as out_file:
            out_file.write("")
        print(f"Empty transcript file written to {output_file_path}")
        return # Exit gracefully

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
        print("Usage: python process_zoom_transcript.py <path_to_zoom_export.json>")
        sys.exit(1)

    input_path = sys.argv[1]
    process_zoom_file(input_path)
