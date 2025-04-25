import json
import sys
import os
import re
import datetime
import glob
from difflib import SequenceMatcher

# --- Configuration ---
OUTPUT_BASE_DIR = "converted_transcripts"
# How many items from the end of the previous transcript part list to check
# against the current part list for overlap detection.
OVERLAP_LOOKBACK_PREVIOUS = 30
# Minimum number of consecutive matching (speaker, text) tuples required
# to confirm an overlap between sequential files. Adjust if needed.
MIN_OVERLAP_LENGTH = 3
# Set to True for detailed print statements during overlap detection
DEBUG_OVERLAP = False

# --- Core Processing Functions ---

def find_transcript_parts_teams(node):
    """
    Recursively searches the Teams JSON data structure for transcript parts
    (speaker and text). Normalizes whitespace in the extracted text.

    Args:
        node: The current node (dict or list) in the JSON structure.

    Returns:
        list: A list of (speaker, text) tuples found under this node.
    """
    parts = []
    if isinstance(node, dict):
        # Specific structure indicating a speaker and their text in Teams JSON
        if (node.get("role") == "AXGroup" and
                len(node.get("children", [])) == 2):
            child1 = node["children"][0]
            child2 = node["children"][1]
            is_speaker_group = (
                isinstance(child1, dict) and
                child1.get("role") == "AXGroup" and
                len(child1.get("children", [])) == 1 and
                isinstance(child1["children"][0], dict) and
                child1["children"][0].get("role") == "AXStaticText" and
                "value" in child1["children"][0]
            )
            is_text_element = (
                isinstance(child2, dict) and
                child2.get("role") == "AXStaticText" and
                "value" in child2
            )
            if is_speaker_group and is_text_element:
                speaker = child1["children"][0]["value"]
                text = child2["value"]
                text = ' '.join(text.split()) # Normalize whitespace
                if text: # Only add if text is not empty after normalization
                    # Clean speaker name (remove potential trailing indicators like '(Guest)')
                    processed_speaker = re.sub(r'\s*\(.*\)\s*$', '', speaker).strip()
                    parts.append((processed_speaker, text))
                # Stop searching deeper within this matched structure
                return parts

        # Recursively search children if not the target structure
        if "children" in node:
            for child in node.get("children", []):
                parts.extend(find_transcript_parts_teams(child))
    elif isinstance(node, list):
        # Recursively search items in a list
        for item in node:
            parts.extend(find_transcript_parts_teams(item))
    return parts

def get_timestamp_from_filename(filename):
    """
    Extracts a datetime object from a filename containing a timestamp
    in the format YYYY-MM-DD-HH-MM-SS.

    Args:
        filename (str): The filename to parse.

    Returns:
        datetime.datetime | None: The extracted timestamp, or None if not found
                                   or not parsable.
    """
    # Regex to find the timestamp pattern
    match = re.search(r"(\d{4}-\d{2}-\d{2}-\d{2}-\d{2}-\d{2})", filename)
    if match:
        try:
            ts_str = match.group(1)
            return datetime.datetime.strptime(ts_str, "%Y-%m-%d-%H-%M-%S")
        except ValueError:
            if DEBUG_OVERLAP: print(f"Warning: Could not parse timestamp from '{ts_str}' in {filename}")
            return None
    if DEBUG_OVERLAP: print(f"Warning: Could not find timestamp pattern in filename {filename}")
    return None

def find_best_overlap_index(previous_parts, current_parts, lookback_prev, min_len):
    """
    Finds the best overlap between the end of previous_parts and the start
    of current_parts using SequenceMatcher to find the longest common subsequence.
    This is more robust to minor variations than exact matching.

    Args:
        previous_parts (list): List of (speaker, text) tuples from previous files.
        current_parts (list): List of (speaker, text) tuples from the current file.
        lookback_prev (int): Max items from the end of previous_parts to consider.
        min_len (int): Minimum length of matching block to consider valid overlap.

    Returns:
        int: The index in `current_parts` *after* the identified overlap.
             Returns 0 if no sufficient overlap is found, meaning all of
             `current_parts` should be considered new.
    """
    if not previous_parts or not current_parts or min_len <= 0:
        return 0 # No basis for overlap or no overlap needed

    len_prev = len(previous_parts)
    len_curr = len(current_parts)

    # Define the slices to compare: end of previous vs. whole of current
    # Use max(0, ...) to prevent negative indices if len_prev < lookback_prev
    prev_slice = previous_parts[max(0, len_prev - lookback_prev):]
    curr_slice = current_parts

    if not prev_slice or not curr_slice:
         return 0 # One of the slices is empty

    # Use SequenceMatcher to find the longest block of matching (speaker, text) tuples
    matcher = SequenceMatcher(None, prev_slice, curr_slice, autojunk=False)
    # Find the longest matching block within the specified ranges
    match = matcher.find_longest_match(0, len(prev_slice), 0, len(curr_slice))

    if DEBUG_OVERLAP:
        print(f"\n  Overlap Check (SequenceMatcher):")
        print(f"    Comparing last {len(prev_slice)} of prev ({len_prev} total) with {len(curr_slice)} of current.")
        print(f"    Longest match details: prev_idx={match.a}, curr_idx={match.b}, size={match.size}")
        # Optional: print matched content for debugging
        # if match.size > 0:
        #     print(f"      Match in prev_slice: {prev_slice[match.a : match.a + match.size]}")
        #     print(f"      Match in curr_slice: {curr_slice[match.b : match.b + match.size]}")

    # Check if the longest match found is sufficiently long
    if match.size >= min_len:
        # If a good match is found, assume the new content starts right after
        # this match ends *in the current_parts list*.
        new_content_start_index = match.b + match.size
        if DEBUG_OVERLAP:
            print(f"  ----> Valid overlap confirmed. Size: {match.size}. New content starts at index {new_content_start_index} in current_parts.")
        # Ensure index doesn't exceed current_parts length (shouldn't happen with find_longest_match logic)
        return min(new_content_start_index, len_curr)
    else:
        # No sufficiently long common subsequence found. Assume no overlap.
        if DEBUG_OVERLAP:
             print(f"  ----> No significant overlap found. Longest match ({match.size}) < min_len ({min_len}). Treating all current parts as new.")
        return 0 # Return 0 to indicate all of current_parts is new


def format_combined_transcript(all_parts):
    """
    Formats the combined list of transcript parts, merging consecutive messages
    from the same speaker.

    Args:
        all_parts (list): List of (speaker, text) tuples.

    Returns:
        str: The final formatted transcript text.
    """
    if not all_parts:
        return ""

    formatted_lines = []
    current_speaker = None
    current_text_buffer = []

    def flush_buffer():
        nonlocal current_speaker, current_text_buffer, formatted_lines
        if current_speaker and current_text_buffer:
            full_message = " ".join(current_text_buffer).strip()
            if full_message:
                # Add speaker tag and the accumulated text
                formatted_lines.append(f"[{current_speaker}]")
                formatted_lines.append(f"{full_message}\n") # Extra newline for readability
        current_text_buffer = [] # Reset buffer

    for speaker, text in all_parts:
        if speaker == current_speaker:
            # Same speaker, append text to buffer
            current_text_buffer.append(text)
        else:
            # Speaker change: flush previous buffer and start new one
            flush_buffer()
            current_speaker = speaker
            current_text_buffer.append(text)

    flush_buffer() # Flush the last speaker's text

    return "\n".join(formatted_lines)

# --- Main Function ---

def process_teams_directory(input_dir_path):
    """
    Main function to process a directory of Teams transcript JSON files.

    Args:
        input_dir_path (str): The path to the input directory containing Teams JSON files.
    """
    print(f"Processing Teams transcript directory: {input_dir_path}")

    if not os.path.isdir(input_dir_path):
        print(f"Error: Input directory not found at '{input_dir_path}'")
        sys.exit(1)

    # Find and sort JSON files by timestamp in filename
    json_files = []
    print("Scanning for JSON files...")
    for root, _, files in os.walk(input_dir_path):
        for file in files:
            if file.lower().endswith(".json"):
                full_path = os.path.join(root, file)
                timestamp = get_timestamp_from_filename(file)
                if timestamp:
                    json_files.append((timestamp, full_path))
                else:
                    print(f"Warning: Skipping file due to missing/unparsable timestamp: {full_path}")

    if not json_files:
        print(f"Error: No JSON files with valid timestamps found in '{input_dir_path}'")
        sys.exit(1)

    json_files.sort() # Sort by timestamp (first element of tuple)
    print(f"Found {len(json_files)} JSON files with timestamps. Processing in chronological order...")

    # Determine output path
    # Use directory name for the output file name
    dir_name = os.path.basename(os.path.normpath(input_dir_path))
    if not dir_name: dir_name = "teams_export" # Fallback name
    output_file_name = f"{dir_name}.txt"

    # Ensure output directory exists
    try:
        os.makedirs(OUTPUT_BASE_DIR, exist_ok=True)
    except OSError as e:
        print(f"Error creating output directory '{OUTPUT_BASE_DIR}': {e}")
        sys.exit(1)

    output_file_path = os.path.join(OUTPUT_BASE_DIR, output_file_name)
    print(f"Output will be saved to: {output_file_path}")

    # Process files chronologically and stitch them
    all_transcript_parts = []
    last_progress_msg_len = 0 # Track length for clearing progress line

    for i, (timestamp, file_path) in enumerate(json_files):
        progress_msg = f"Processing file {i+1}/{len(json_files)}: {os.path.basename(file_path)}..."
        # Print progress message, overwriting previous one
        print(progress_msg + " " * (last_progress_msg_len - len(progress_msg)), end='\r', flush=True)
        last_progress_msg_len = len(progress_msg)
        if DEBUG_OVERLAP: print(f"\n{progress_msg}") # Print clearly on new line in debug mode

        try:
            with open(file_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            current_parts = find_transcript_parts_teams(data)
            if not current_parts:
                 if DEBUG_OVERLAP: print("  No transcript parts found in this file.")
                 continue # Skip empty files

        except json.JSONDecodeError as e:
            print(f"\nWarning: Skipping file {os.path.basename(file_path)} due to JSON decode error: {e}")
            continue
        except Exception as e:
            # Print warning on a new line so it doesn't get overwritten by progress
            print(f"\nWarning: Skipping file {os.path.basename(file_path)} due to read/parse error: {e}")
            continue

        if not all_transcript_parts: # This is the first file with content
            all_transcript_parts.extend(current_parts)
            if DEBUG_OVERLAP: print(f"  Added {len(current_parts)} parts from the first file.")
        else:
            # Find where new content starts in current_parts based on overlap with the end of all_transcript_parts
            new_content_start_idx = find_best_overlap_index(
                all_transcript_parts,
                current_parts,
                lookback_prev=OVERLAP_LOOKBACK_PREVIOUS,
                min_len=MIN_OVERLAP_LENGTH
            )

            # Add only the non-overlapping parts from the current file
            if new_content_start_idx < len(current_parts):
                new_parts_to_add = current_parts[new_content_start_idx:]
                all_transcript_parts.extend(new_parts_to_add)
                if DEBUG_OVERLAP: print(f"  Overlap detected. Added {len(new_parts_to_add)} new parts (from index {new_content_start_idx}).")
            elif DEBUG_OVERLAP:
                # This means the heuristic determined the entire current file overlapped
                 print(f"  All {len(current_parts)} items considered overlap (start index {new_content_start_idx}). Nothing new added.")

    # Clear the final progress indicator line before printing summary
    print(" " * last_progress_msg_len, end='\r')

    print(f"\nProcessing complete. Total unique transcript parts collected: {len(all_transcript_parts)}")

    if not all_transcript_parts:
        print("Warning: No transcript content found after processing all files.")
        # Write an empty file
        with open(output_file_path, "w", encoding="utf-8") as f:
             f.write("")
        print(f"Empty transcript file written to {output_file_path}")
        return # Exit gracefully

    # Format the combined transcript
    formatted_output = format_combined_transcript(all_transcript_parts)

    # Write the final output
    try:
        with open(output_file_path, "w", encoding="utf-8") as f:
             f.write(formatted_output)
        print(f"Combined transcript successfully written to {output_file_path}")
    except IOError as e:
        print(f"An error occurred while writing the output file '{output_file_path}': {e}")
        sys.exit(1)

# --- Script Execution ---

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python process_teams_transcript.py <path_to_teams_export_directory>")
        sys.exit(1)

    input_directory = sys.argv[1]
    process_teams_directory(input_directory)
