#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
macOS Accessibility UI Element Exporter

This script uses the macOS Accessibility API (via ApplicationServices)
to serialize the UI element tree of running applications. It is primarily
designed for extracting text content like captions or transcripts from
meeting clients (e.g., Zoom, Teams).

It includes presets for common applications and manual modes for exploring
the UI tree of arbitrary applications. Exports are saved as timestamped
JSON files.
"""

# --- Imports ---
import subprocess
import time
import os
import json
from datetime import datetime
from ApplicationServices import (
    AXUIElementCreateApplication,
    AXUIElementCopyAttributeValue,
    AXIsProcessTrusted,
    kAXChildrenAttribute,
    kAXTitleAttribute,
    kAXValueAttribute,
    kAXRoleAttribute,           # Explicitly import if used via constant AX_ROLE
    kAXSubroleAttribute,        # Explicitly import if used via constant AX_SUBROLE
    kAXDescriptionAttribute,    # Explicitly import if used via constant AX_DESCRIPTION
    kAXHelpAttribute,           # Explicitly import if used via constant AX_HELP
    kAXLabelValueAttribute,     # Explicitly import if used via constant AX_LABEL_VALUE
    kAXWindowRole,              # Explicitly import if used via string "AXWindow"
    kAXRowRole                  # Explicitly import if used via string "AXRow"
)
import copy
import signal # For checking PID existence
import argparse # For verbosity flag
import traceback # For detailed error printing

# --- Accessibility Attribute Constants ---
AX_ROLE = kAXRoleAttribute
AX_SUBROLE = kAXSubroleAttribute
AX_DESCRIPTION = kAXDescriptionAttribute
AX_HELP = kAXHelpAttribute
AX_LABEL_VALUE = kAXLabelValueAttribute
AX_VALUE = kAXValueAttribute
AX_TITLE = kAXTitleAttribute
AX_CHILDREN = kAXChildrenAttribute
AX_WINDOW_ROLE = kAXWindowRole
AX_ROW_ROLE = kAXRowRole

# --- Configuration ---
BASE_EXPORT_DIR_NAME = "exports" # Base directory name for saving exports

# --- App Context Presets ---
# Define how to find the initial target application or window.
APP_CONTEXTS = {
    "Zoom Transcript Window": {
        "app_label": "Zoom",
        "find_method": "pid_and_exact_window", # Find specific PID, then exact window title
        "cmd_path": "/Applications/zoom.us.app/Contents/MacOS/zoom.us",
        "window_title": "Transcript",
        "default_serialization_preset_name": "Zoom Transcript Table"
    },
    "Teams in Browser (Chrome/Prisma)": {
        "app_label": "TeamsInBrowser", # Will be refined based on actual browser found
        "find_method": "specific_pids_containing_window", # Find specific PIDs, then windows containing title fragment
        "cmd_paths": [ # List of target browser executables
            "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
            "/Applications/Prisma Access Browser.app/Contents/MacOS/Prisma Access Browser"
        ],
        "window_title_fragment": "teams", # Case-insensitive search term in window title
        "default_serialization_preset_name": "Teams Live Captions Group"
    },
    # --- Manual Options ---
     "Enter PID Directly": {
         "app_label": "ManualPIDEntry",
         "find_method": "manual_pid_direct", # Prompt user for PID
     },
     "Manual PID Search": {
         "app_label": "ManualPIDSearch",
         "find_method": "manual_pid_search_only", # Prompt user for process name/query, select PID
     },
      "Manual Window Search": {
         "app_label": "ManualWindowSearch",
         "find_method": "manual_window_search_only", # Prompt user for window title fragment, search all apps
     }
}

# --- Serialization Presets ---
# Define what to serialize once the initial element is found.
# --- Serialization Presets ---
SERIALIZATION_PRESETS = {
    "Full Element Found": {
        "description": "Serialize the entire App/Window element found by the context.",
        "target_criteria": None,
        "default_depth": None,
        "default_interval": None,
        "text_line_roles": [] # Nothing specific to count here by default
    },
    "Zoom Transcript Table": {
        "description": "Finds and serializes the AXTable containing Zoom transcript text.",
        "target_criteria": {"role": "AXTable", "description": "Transcript list"},
        "applicable_contexts": ["Zoom Transcript Window"],
        "default_depth": 25,
        "default_interval": 30,
        "text_line_roles": ["AXTextArea"] # Count AXTextArea elements in Zoom transcript
    },
    "Teams Live Captions Group": {
        "description": "Finds and serializes the AXGroup containing Teams Live Captions.",
        "target_criteria": {"role": "AXGroup", "description": "Live Captions"},
        "default_depth": 50,
        "default_interval": 30,
        "text_line_roles": ["AXStaticText"] # Count AXStaticText elements in Teams captions
    },
}

# --- Helper Functions ---

def verbose_print(args, *print_args, **print_kwargs):
    """Prints only if the verbose flag is set."""
    if args.verbose:
        print(*print_args, **print_kwargs)

def get_attribute(element, attr):
    """
    Safely retrieves an accessibility attribute value from an AXUIElement.

    Args:
        element: The AXUIElement to query.
        attr: The accessibility attribute constant (e.g., AX_TITLE).

    Returns:
        The attribute value if successful, otherwise None.
    """
    if not element or not attr:
        return None
    result, value = AXUIElementCopyAttributeValue(element, attr, None)
    if result == 0:
        return value
    # Return None for any error, including attribute not supported
    return None

def serialize_ax_element(element, depth=0, max_depth=50, text_roles_to_count=None):
    """
    Recursively serializes an accessibility element and its children into a dictionary.

    Args:
        element: The starting AXUIElement.
        depth: Current recursion depth.
        max_depth: Maximum recursion depth.
        text_roles_to_count: Optional list of AXRole strings to count as relevant text elements.

    Returns:
        A tuple containing:
        - A dictionary representing the element and its children (or None if empty/error).
        - An integer count of elements whose role matched one in text_roles_to_count.
    """
    if not element or depth > max_depth:
        # Return 0 count if max depth reached or element is invalid
        return ({"error": f"<Max depth {max_depth} reached>"} if depth > max_depth else None), 0

    # Initialize count for this level
    text_element_count = 0
    data = {}
    # Ensure text_roles_to_count is a list/set for efficient lookup, even if None was passed
    roles_to_count_set = set(text_roles_to_count) if text_roles_to_count else set()

    # Attributes to attempt to serialize for the current element
    attributes_to_check = {
        "role": AX_ROLE, "subrole": AX_SUBROLE, "title": AX_TITLE,
        "value": AX_VALUE, "description": AX_DESCRIPTION, "help": AX_HELP,
        "label": AX_LABEL_VALUE,
    }

    current_element_role = None # Store role for counting check

    for key, attr_name in attributes_to_check.items():
        try:
            value = get_attribute(element, attr_name)
            if value is not None:
                # Store simple types or represent complex ones safely
                if isinstance(value, str) and value:
                    data[key] = value
                elif isinstance(value, (int, float, bool)):
                    data[key] = value
                elif not isinstance(value, str): # Handle non-string, non-simple types
                    try:
                        data[key] = repr(value) # Fallback representation
                    except Exception:
                        data[key] = "<Unrepresentable CFType>"

                # Store the role if found
                if key == "role" and isinstance(value, str):
                   current_element_role = value

        except Exception as e:
            # Record error if attribute fetching fails unexpectedly
            data[f"{key}_error"] = f"Error fetching {attr_name}: {e}"

    # --- New Counting Logic ---
    # Check if the current element's role is one we should count
    if current_element_role and current_element_role in roles_to_count_set:
        text_element_count += 1
        # Optional: Add a marker to the serialized data itself?
        # data["is_counted_text"] = True

    # Recursively serialize children
    try:
        children = get_attribute(element, AX_CHILDREN)
        if children:
            children_data = []
            for child in children:
                # Recursive call, passing down the roles to count
                child_data, child_count = serialize_ax_element(
                    child,
                    depth + 1,
                    max_depth,
                    text_roles_to_count=text_roles_to_count # Pass the original list/None
                )
                if child_data:  # Append only if child serialization returned data
                    children_data.append(child_data)
                # Accumulate count from children
                text_element_count += child_count
            if children_data:
                data["children"] = children_data
    except Exception as e:
        data["children_error"] = f"Error fetching children: {e}"

    # Only return data if it contains meaningful information beyond errors
    has_real_data = any(not k.endswith("_error") for k in data if k != "children") or "children" in data
    # Return data and the accumulated count
    return data if has_real_data else None, text_element_count


def find_element_by_criteria(start_element, criteria, args, current_depth=0, max_search_depth=50):
    """
    Recursively searches (Depth First Search) for the first element matching all specified criteria.

    Args:
        start_element: The AXUIElement to start searching from.
        criteria: A dictionary where keys are attribute names (e.g., "role", "description")
                  and values are the expected values.
        args: Command line arguments (for verbose_print).
        current_depth: Current search depth (for recursion control).
        max_search_depth: Maximum depth to search.

    Returns:
        The first matching AXUIElement found, or None if not found or max depth reached.
    """
    if not start_element or not criteria or current_depth > max_search_depth:
        return None

    verbose_print(args, f"{'  ' * current_depth} Searching element at depth {current_depth} for {criteria}...")

    match = True
    element_attrs = {}

    # Map criteria keys to Accessibility API constants
    criteria_map = {
        "role": AX_ROLE, "subrole": AX_SUBROLE, "title": AX_TITLE,
        "description": AX_DESCRIPTION, "value": AX_VALUE, "help": AX_HELP,
        "label": AX_LABEL_VALUE
        # Add other attributes here if needed for criteria matching
    }

    # Check if the current element matches all criteria
    for key, expected_value in criteria.items():
        attr_name = criteria_map.get(key)
        if not attr_name:
            print(f"Warning: Unknown criteria key '{key}'")
            match = False
            break # Cannot match an unknown attribute

        actual_value = get_attribute(start_element, attr_name)
        element_attrs[key] = actual_value # Store for logging
        verbose_print(args, f"{'  ' * current_depth}   Attr '{key}': Expected='{expected_value}', Actual='{actual_value}'")

        # Perform comparison (handle potential type differences if necessary, though usually strings)
        if actual_value != expected_value:
            match = False
            break # Stop checking criteria for this element if one fails

    if match:
        verbose_print(args, f"{'  ' * current_depth}   🎉 Match found!")
        return start_element # Found the target element

    verbose_print(args, f"{'  ' * current_depth}   No match at this level. Checking children...")

    # If not matched, search children recursively
    children = get_attribute(start_element, AX_CHILDREN)
    if children:
        verbose_print(args, f"{'  ' * current_depth}   Found {len(children)} children.")
        for child in children:
            found_element = find_element_by_criteria(child, criteria, args, current_depth + 1, max_search_depth)
            if found_element:
                return found_element # Propagate the found element up the recursion chain

    verbose_print(args, f"{'  ' * current_depth}   No match found in this branch (depth {current_depth}).")
    return None # Not found in this element or its descendants


def find_process_by_cmd(cmd_path, args):
    """
    Finds a process ID (PID) by its command path using pgrep and ps for verification.

    Args:
        cmd_path: The full path to the executable.
        args: Command line arguments (for verbose_print).

    Returns:
        The integer PID if found and verified, otherwise None.
    """
    basename = os.path.basename(cmd_path)
    pid_found = None

    # 1. Try pgrep with the full command path
    verbose_print(args, f"   Running: pgrep -f '{cmd_path}'")
    result_full = subprocess.run(["pgrep", "-f", cmd_path], capture_output=True, text=True)

    potential_pids_full = []
    if result_full.returncode == 0 and result_full.stdout.strip():
        potential_pids_full = [line.strip() for line in result_full.stdout.strip().splitlines()]
        verbose_print(args, f"   pgrep -f found potential PIDs: {potential_pids_full}")

    # Verify PIDs found by full path using `ps`
    for pid_str in potential_pids_full:
        try:
            pid = int(pid_str)
            ps_result = subprocess.run(["ps", "-o", "command=", "-p", str(pid)], capture_output=True, text=True)
            if ps_result.returncode == 0:
                ps_command = ps_result.stdout.strip()
                # Check if the exact cmd_path is in the command AND it's not the pgrep process itself
                if cmd_path in ps_command and 'pgrep' not in ps_command.lower():
                    verbose_print(args, f"   Verified PID {pid} via ps for '{cmd_path}'.")
                    pid_found = pid
                    break # Found a verified PID, use the first one
        except (ValueError, subprocess.CalledProcessError):
            continue # Ignore invalid PIDs or ps errors

    if pid_found:
        return pid_found

    # 2. If full path failed, try pgrep with just the basename
    verbose_print(args, f"   Full path search failed or no verified PID. Trying basename: pgrep -fl '{basename}'")
    result_base = subprocess.run(["pgrep", "-fl", basename], capture_output=True, text=True)

    if result_base.returncode == 0 and result_base.stdout.strip():
        verbose_print(args, f"   Found potential matches by basename:\n{result_base.stdout.strip()}")
        for line in result_base.stdout.strip().splitlines():
            try:
                pid_str, potential_cmd = line.strip().split(" ", 1)
                pid = int(pid_str)
                # Verify this PID using ps and check if the command contains the *original* full cmd_path
                ps_result = subprocess.run(["ps", "-o", "command=", "-p", str(pid)], capture_output=True, text=True)
                if ps_result.returncode == 0:
                     ps_command = ps_result.stdout.strip()
                     if cmd_path in ps_command and 'pgrep' not in ps_command.lower():
                         verbose_print(args, f"   Found PID {pid} via basename search, verified full path '{cmd_path}' in ps output: {ps_command}")
                         pid_found = pid
                         break # Found a verified PID
            except (ValueError, subprocess.CalledProcessError):
                continue # Ignore lines that don't split or ps errors

    if pid_found:
        return pid_found
    else:
        verbose_print(args, f"   No verified PID found for '{cmd_path}' using either full path or basename search.")
        return None


def choose_pid_manually(query, args):
    """
    Allows manual selection of a PID from processes matching a query string.

    Args:
        query: The string to search for in process command lines using pgrep.
        args: Command line arguments (for verbose_print).

    Returns:
        The selected integer PID, or None if no processes found or selection cancelled.
    """
    verbose_print(args, f"   Running: pgrep -fl '{query}'")
    result = subprocess.run(["pgrep", "-fl", query], capture_output=True, text=True)
    processes = [] # List of tuples: (pid, cmd)
    if result.returncode == 0 and result.stdout:
        for line in result.stdout.strip().splitlines():
            try:
                pid_str, cmd = line.strip().split(" ", 1)
                # Filter out the pgrep command itself to avoid confusion
                if 'pgrep -fl' not in cmd:
                    processes.append((int(pid_str), cmd))
            except ValueError:
                continue # Ignore lines that don't split correctly

    if not processes:
        print(f"❌ No running processes found matching '{query}'.")
        return None

    print("\nMatching processes:")
    for idx, (pid, cmd) in enumerate(processes):
        print(f"[{idx: >2}] PID: {pid: <6} | CMD: {cmd}")

    while True:
        try:
            choice = input(f"Select index [0-{len(processes)-1}] (or press Enter to cancel): ").strip()
            if choice == "":
                print("   Selection cancelled.")
                return None
            idx = int(choice)
            if 0 <= idx < len(processes):
                return processes[idx][0] # Return the selected PID
            else:
                print("Invalid index.")
        except ValueError:
            print("Invalid input. Please enter a number.")


def search_window_titles_across_apps(title_fragment, args):
    """
    Searches all running applications for AXWindows containing a title fragment.

    Args:
        title_fragment: The case-insensitive string to search for in window titles.
        args: Command line arguments (for verbose_print).

    Returns:
        A list of tuples: (pid, window_title, command_name, window_element).
        Returns an empty list if no matches found or errors occur.
    """
    verbose_print(args, "   Getting list of running processes: ps -axo pid=,comm=")
    result = subprocess.run(["ps", "-axo", "pid=,comm="], capture_output=True, text=True)
    matches = [] # Store (pid, title, cmd, element)
    processed_pids = set() # Avoid processing the same PID multiple times

    if not result.stdout:
        verbose_print(args, "   ps command returned no output.")
        return matches

    for line in result.stdout.strip().splitlines():
        try:
            pid_str, cmd = line.strip().split(maxsplit=1)
            pid = int(pid_str)
            if pid in processed_pids: continue
            processed_pids.add(pid)

            verbose_print(args, f"   Checking PID {pid} ({cmd})...")
            app_element = AXUIElementCreateApplication(pid)
            if not app_element:
                # Common if process lacks GUI or permissions are insufficient
                verbose_print(args, f"   Skipping PID {pid}: Could not create AXUIElement.")
                continue

            # Get direct children, which often include windows
            children = get_attribute(app_element, AX_CHILDREN)
            if children:
                verbose_print(args, f"   Found {len(children)} potential children for PID {pid}.")
                for i, child_element in enumerate(children):
                    # Check if the child is actually a window
                    role = get_attribute(child_element, AX_ROLE)
                    if role != AX_WINDOW_ROLE:
                         # verbose_print(args, f"      Child {i} is not AXWindow (Role: {role}), skipping.")
                         continue

                    # Get the title and check for the fragment
                    title = get_attribute(child_element, AX_TITLE)
                    if isinstance(title, str) and title_fragment.lower() in title.lower():
                         verbose_print(args, f"      Match found: Window {i}, Title: '{title}'")
                         matches.append((pid, title.strip(), cmd, child_element)) # Store the window element

            # else: verbose_print(args, f"   No children found or error getting children for PID {pid}.")

        except ValueError:
            verbose_print(args, f"   Skipping line (format error): {line.strip()}")
        except Exception as e:
            # Catch potential errors during AX interaction for a specific PID
            verbose_print(args, f"   Error processing PID {pid_str} ({cmd}): {e}")
            continue # Continue to the next process

    verbose_print(args, f"   Found {len(matches)} total window matches across all PIDs.")
    return matches


def find_window_by_title(app_element, title_query, args, match_type="contains"):
    """
    Finds the FIRST AXWindow within a specific app element matching the title query.

    Args:
        app_element: The AXUIElement of the application to search within.
        title_query: The string to search for in the window title.
        args: Command line arguments (for verbose_print).
        match_type: "contains" (case-insensitive) or "exact" (case-sensitive).

    Returns:
        The first matching AXUIElement (window) found, or None.
    """
    if not app_element: return None

    children = get_attribute(app_element, AX_CHILDREN)
    if not children:
        verbose_print(args, f"   find_window_by_title: No children found for the app element.")
        return None

    verbose_print(args, f"   find_window_by_title: Searching {len(children)} children for title '{title_query}' (match: {match_type})...")
    for i, child_element in enumerate(children):
        # Check if it's a window first
        role = get_attribute(child_element, AX_ROLE)
        if role != AX_WINDOW_ROLE:
             # verbose_print(args, f"      Child {i} is not AXWindow (Role: {role}), skipping.")
             continue

        # Get title and perform match
        title = get_attribute(child_element, AX_TITLE)
        if isinstance(title, str):
            title_str = title.strip()
            matches = False
            if match_type == "exact":
                matches = (title_str == title_query)
            elif match_type == "contains":
                matches = (title_query.lower() in title_str.lower())
            else: # Default to contains
                matches = (title_query.lower() in title_str.lower())

            # verbose_print(args, f"      Checking window {i}: Title='{title_str}', Matches={matches}")
            if matches:
                verbose_print(args, f"      Found first match: Window {i} with Title '{title_str}'")
                return child_element # Return the first matching window element

    verbose_print(args, f"   find_window_by_title: No matching window found.")
    return None


def find_all_windows_by_title(app_element, title_query, args, match_type="contains"):
    """
    Finds ALL AXWindows within a specific app element matching the title query.

    Args:
        app_element: The AXUIElement of the application to search within.
        title_query: The string to search for in the window title.
        args: Command line arguments (for verbose_print).
        match_type: "contains" (case-insensitive) or "exact" (case-sensitive).

    Returns:
        A list of tuples: (window_title, window_element). Returns an empty list if no matches.
    """
    matches = [] # Store (title, element) tuples
    if not app_element: return matches

    children = get_attribute(app_element, AX_CHILDREN)
    if not children:
        verbose_print(args, f"   find_all_windows_by_title: No children found for the app element.")
        return matches

    verbose_print(args, f"   find_all_windows_by_title: Searching {len(children)} children for title '{title_query}' (match: {match_type})...")
    for i, child_element in enumerate(children):
        # Check if it's a window first
        role = get_attribute(child_element, AX_ROLE)
        if role != AX_WINDOW_ROLE:
             # verbose_print(args, f"      Child {i} is not AXWindow (Role: {role}), skipping.")
             continue

        # Get title and perform match
        title = get_attribute(child_element, AX_TITLE)
        if isinstance(title, str):
            title_str = title.strip()
            does_match = False
            if match_type == "exact":
                does_match = (title_str == title_query)
            else: # Default contains
                does_match = (title_query.lower() in title_str.lower())

            if does_match:
                verbose_print(args, f"      Found match: Window {i}, Title: '{title_str}'")
                matches.append((title_str, child_element)) # Store title and element

    verbose_print(args, f"   find_all_windows_by_title: Found {len(matches)} matching windows.")
    return matches


def pid_exists(pid):
    """ Check For the existence of a unix pid. """
    if pid is None or not isinstance(pid, int) or pid <= 0:
        return False
    try:
        os.kill(pid, 0) # Signal 0 doesn't kill, just checks existence/permissions
    except OSError:
        # errno 3: No such process
        # errno 1: Operation not permitted (process exists but owned by different user)
        # Either way, we can't reliably interact with it or it's gone.
        return False
    else:
        return True # Process exists and we have permission to signal it


# --- Export Function ---

# --- Updated Export Function ---
def export_to_json(element, base_export_dir, depth, args, serialization_config):
    """
    Serializes the given element and saves it to a timestamped JSON file.

    Args:
        element: The AXUIElement to serialize.
        base_export_dir: The directory where the JSON file will be saved.
        depth: The maximum serialization depth.
        args: Command line arguments (for verbose_print).
        serialization_config: The dictionary for the selected Serialization Preset.
    """
    if not element:
        print(f"❌ Cannot export, element is None.")
        return
    if not serialization_config:
         print(f"❌ Cannot export, serialization_config is missing.")
         # Assign default empty list if config is missing, though this indicates a bug elsewhere
         serialization_config = {}

    # Try to get a meaningful name for logging/status
    element_desc = get_attribute(element, AX_DESCRIPTION) or \
                   get_attribute(element, AX_TITLE) or \
                   get_attribute(element, AX_ROLE) or \
                   "UnknownElement"
    print(f"⏳ Serializing element tree starting from '{element_desc}' (max depth: {depth})...")

    # Get the list of roles to count from the serialization config
    text_roles = serialization_config.get("text_line_roles", []) # Default to empty list
    if text_roles:
         verbose_print(args, f"   Counting elements with roles: {text_roles}")
    else:
         verbose_print(args, f"   No specific text element roles defined for counting.")


    start_time = time.time()
    # Pass the roles to count to the serialization function
    serialized_data, text_element_count = serialize_ax_element(
        element,
        max_depth=depth,
        text_roles_to_count=text_roles
    )
    end_time = time.time()
    print(f"⏱️ Serialization finished in {end_time - start_time:.2f} seconds.")

    if not serialized_data:
        print("❓ Serialization resulted in empty data (possibly due to depth limit or inaccessible elements). Nothing to save.")
        return

    # Generate timestamp and filename
    timestamp = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
    filename = f"export_{timestamp}.json"
    filepath = os.path.join(base_export_dir, filename)

    try:
        # Ensure the directory exists
        os.makedirs(base_export_dir, exist_ok=True)
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(serialized_data, f, indent=2, ensure_ascii=False)
        # Update the success message to use the new count
        count_desc = f"{text_element_count} relevant text elements found" if text_roles else "Count not applicable"
        print(f"✅ Export saved to: {filepath} ({count_desc})")
    except Exception as e:
        print(f"❌ Error saving JSON file '{filepath}': {e}")
        verbose_print(args, traceback.format_exc())


# --- Main Workflow Functions ---
def get_initial_element_from_context(context_config, args):
    """
    Finds the initial Application or Window AXUIElement based on the selected App Context configuration.

    Args:
        context_config: The dictionary defining the chosen application context.
        args: Command line arguments (for verbose_print).

    Returns:
        A tuple containing:
        - The found AXUIElement (can be an App or Window element, or None on failure/cancel).
        - The refined application label (e.g., including browser name, or None).
        - The process ID (PID) of the application (or None).
    """
    find_method = context_config.get("find_method")
    app_label = context_config.get("app_label", "UnknownContext") # Initial label
    pid = None
    app_element = None
    window_element = None # Represents the final element to be returned if it's a window

    verbose_print(args, f"--- Finding Element using Method: {find_method} ---")

    # --- Methods involving finding a PID first ---
    if find_method in ["pid_and_exact_window", "pid_and_containing_window",
                       "manual_pid_search_only", "manual_pid_then_containing_window",
                       "manual_pid_direct"]:

        # --- Sub-methods for Finding the PID ---
        if find_method == "manual_pid_direct":
             while pid is None:
                 try:
                     pid_input = input("Enter the PID number (or press Enter to cancel): ").strip()
                     if not pid_input: print("   PID entry cancelled."); return None, None, None
                     pid_val = int(pid_input)
                     if pid_exists(pid_val):
                         pid = pid_val
                         print(f"   Using PID: {pid}")
                     else:
                         print(f"❌ Process with PID {pid_val} not found or inaccessible.")
                 except ValueError: print("❌ Invalid PID. Please enter a number.")
                 except Exception as e: print(f"❌ Error checking PID: {e}")
                 if pid is None: # Ask to retry if PID not found/valid
                     retry = input("Try entering PID again? (Y/n): ").strip().lower()
                     if retry.startswith('n'): return None, None, None

        elif find_method == "manual_pid_search_only":
            while pid is None:
                 query = input(f"Enter process name/query to find PID (leave empty to cancel): ").strip()
                 if not query: print("   PID search cancelled."); return None, None, None
                 pid = choose_pid_manually(query, args)
                 if pid is None: return None, None, None # choose_pid_manually handles cancellation message

        elif find_method == "manual_pid_then_containing_window":
             search_hint = context_config.get("pid_search_hint", "application name")
             while pid is None:
                 query = input(f"Enter query to find PID (hint: {search_hint}, leave empty to cancel): ").strip()
                 if not query: print("   PID search cancelled."); return None, None, None
                 pid = choose_pid_manually(query, args)
                 if pid is None: return None, None, None # choose_pid_manually handles cancellation message

        elif find_method in ["pid_and_exact_window", "pid_and_containing_window"]:
            cmd_path = context_config.get("cmd_path")
            if not cmd_path: print(f"❌ Context '{app_label}' misconfigured: Needs 'cmd_path'."); return None, None, None
            print(f"   Finding PID for '{cmd_path}'...")
            pid = find_process_by_cmd(cmd_path, args)
            if not pid: print(f"❌ PID not found for '{cmd_path}'. Is the application running?"); return None, None, None
            print(f"   Found PID: {pid}")

        # --- Common logic block after a PID has been found ---
        if pid:
            print(f"   Creating application element for PID: {pid}...")
            app_element = AXUIElementCreateApplication(pid)
            if not app_element:
                print(f"❌ Failed to create application element for PID {pid}. Check Accessibility permissions and application state.")
                return None, None, None
            verbose_print(args, "   Application element created successfully.")

            # If the method requires finding a specific window *within* this PID
            if find_method in ["pid_and_exact_window", "pid_and_containing_window", "manual_pid_then_containing_window"]:
                window_title_query = "" # The title/fragment used for searching
                match_type = "contains" # Default match type

                if find_method == "manual_pid_then_containing_window":
                    # Prompt user for window title fragment
                    window_title_hint = context_config.get("window_title_hint", "partial window title")
                    current_window_element = None
                    while current_window_element is None:
                         title_query_input = input(f"Enter window title fragment (hint: {window_title_hint}, leave empty to cancel): ").strip()
                         if not title_query_input: print("   Window search cancelled."); return None, None, None
                         window_title_query = title_query_input # Store query for messages
                         current_window_element = find_window_by_title(app_element, window_title_query, args, "contains")
                         if current_window_element is None:
                              print(f"❌ Window containing '{window_title_query}' not found in PID {pid}.")
                              try_again = input("Try searching again? (y/N): ").strip().lower()
                              if not try_again.startswith('y'): return None, None, None
                         else:
                              window_element = current_window_element # Assign to final variable
                              break # Exit loop, window found

                else: # pid_and_exact_window or pid_and_containing_window
                     # Get title/fragment from config
                     window_title_config = context_config.get("window_title") or context_config.get("window_title_fragment")
                     if not window_title_config: print(f"❌ Context '{app_label}' misconfigured: Needs 'window_title' or 'window_title_fragment'."); return None, None, None
                     window_title_query = window_title_config
                     match_type = "exact" if find_method == "pid_and_exact_window" else "contains"
                     print(f"   Finding window '{window_title_query}' (match: {match_type})...")
                     window_element = find_window_by_title(app_element, window_title_query, args, match_type)

                # Check result of window finding attempts
                if window_element:
                    actual_title = get_attribute(window_element, AX_TITLE) or "Unknown Title"
                    print(f"✅ Found initial window element: '{actual_title}'.")
                    return window_element, app_label, pid # Return the specific WINDOW element
                else:
                    print(f"❌ Window matching '{window_title_query}' not found for PID {pid}.")
                    return None, None, None

            else:
                # Methods like manual_pid_direct, manual_pid_search_only return the APP element
                print(f"✅ Found initial application element (PID: {pid}).")
                return app_element, app_label, pid # Return the APP element

        else:
            # This path should only be reached if PID finding failed AND returned None gracefully.
            verbose_print(args, "PID was not found or user cancelled.")
            return None, None, None
    # --- End of PID-first methods ---


    # --- Method: Find specific browser PIDs, then find window containing fragment ---
    elif find_method == "specific_pids_containing_window":
        cmd_paths = context_config.get("cmd_paths")
        window_title_fragment = context_config.get("window_title_fragment")

        # Basic config validation
        if not cmd_paths or not isinstance(cmd_paths, list):
            print(f"❌ Context '{app_label}' misconfigured: Needs 'cmd_paths' (list)."); return None, None, None
        if not window_title_fragment or not isinstance(window_title_fragment, str):
            print(f"❌ Context '{app_label}' misconfigured: Needs 'window_title_fragment' (string)."); return None, None, None

        matching_windows = [] # Store dicts: {"pid", "title", "cmd_path", "browser_name", "element"}
        print(f"   Searching for windows containing '{window_title_fragment}' in specific browser processes:")

        # --- Find all potentially matching windows across configured browsers ---
        for cmd_path in cmd_paths:
            try: # Try to derive a user-friendly browser name from the path
                browser_name = os.path.basename(os.path.dirname(os.path.dirname(cmd_path))).replace('.app', '')
            except: browser_name = os.path.basename(cmd_path) # Fallback
            print(f"   - Checking {browser_name} ('{cmd_path}')...")

            pid_found = find_process_by_cmd(cmd_path, args)
            if pid_found:
                print(f"     Found PID: {pid_found}")
                current_app_element = AXUIElementCreateApplication(pid_found)
                if current_app_element:
                    verbose_print(args, f"     Searching windows in PID {pid_found} for '{window_title_fragment}'...")
                    # Find *all* matching windows within this specific browser process
                    found_in_pid = find_all_windows_by_title(current_app_element, window_title_fragment, args, match_type="contains")
                    for title, win_element in found_in_pid:
                        matching_windows.append({
                            "pid": pid_found,
                            "title": title,
                            "cmd_path": cmd_path,
                            "browser_name": browser_name,
                            "element": win_element # Store the actual window AXUIElement
                        })
                    print(f"     Found {len(found_in_pid)} matching window(s) in {browser_name} (PID: {pid_found}).")
                else:
                    print(f"     ❌ Could not create application element for PID {pid_found}.")
            else:
                print(f"     Process not found or running for {browser_name}.")
        # --- End of window finding loop ---

        # --- Process the results of the browser search ---
        num_matches = len(matching_windows)

        if num_matches == 0:
            # No windows found matching the fragment in any specified browser
            browser_names = [os.path.basename(p) for p in cmd_paths]
            print(f"❌ No windows containing '{window_title_fragment}' found in the specified browser processes ({', '.join(browser_names)}).")
            return None, None, None

        elif num_matches == 1:
            # Exactly one window found - this is the ideal case
            match = matching_windows[0]
            pid = match["pid"]
            window_element = match["element"]
            # Refine the app label to include the browser where it was found
            app_label = f"{app_label} ({match['browser_name']})"
            print(f"✅ Found unique window: '{match['title']}' (PID: {pid}, Browser: {match['browser_name']})")
            return window_element, app_label, pid # Return the specific WINDOW element

        else: # num_matches > 1
            # --- MODIFIED LOGIC FOR MULTIPLE MATCHES ---
            print(f"\nℹ️ Found multiple ({num_matches}) windows containing '{window_title_fragment}'. Attempting to auto-select the one containing the target element...")
            for idx, match in enumerate(matching_windows):
                 print(f"  Candidate [{idx}] PID: {match['pid']:<6} | Browser: {match['browser_name']:<20} | Title: '{match['title']}'")

            # Try to find the target element defined by the default serialization preset
            target_criteria = None
            default_preset_name = context_config.get("default_serialization_preset_name")
            if default_preset_name and default_preset_name in SERIALIZATION_PRESETS:
                serialization_config = SERIALIZATION_PRESETS[default_preset_name]
                target_criteria = serialization_config.get("target_criteria")
                verbose_print(args, f"   Using target criteria from default preset '{default_preset_name}': {target_criteria}")
            else:
                verbose_print(args, f"   No default serialization preset or target criteria found for context. Cannot auto-select based on content.")

            selected_match = None
            if target_criteria:
                # Search within each candidate window for the target criteria
                for idx, match in enumerate(matching_windows):
                    print(f"   Searching within Candidate [{idx}]: '{match['title']}'...")
                    # Use find_element_by_criteria to check if this window contains the target
                    found_target_in_window = find_element_by_criteria(match['element'], target_criteria, args, max_search_depth=15) # Limit depth for speed?
                    if found_target_in_window:
                        print(f"✅ Found target element ({target_criteria}) within Candidate [{idx}]. Selecting this window.")
                        selected_match = match
                        break # Stop searching once the target is found in a window
                    else:
                        verbose_print(args, f"   Target element not found in Candidate [{idx}].")

                if not selected_match:
                    print(f"⚠️ Target element ({target_criteria}) not found in any of the {num_matches} candidate windows.")
            else:
                 print(f"⚠️ Cannot auto-select based on target criteria (not defined for this context).")


            # --- Fallback/Final Selection ---
            if selected_match:
                # We found the target in one of the windows
                pid = selected_match["pid"]
                window_element = selected_match["element"]
                app_label = f"{app_label} ({selected_match['browser_name']})" # Refine label
                print(f"✅ Auto-selected window: '{selected_match['title']}' (PID: {pid}) as it contains the target.")
                return window_element, app_label, pid
            else:
                while True: # Loop for user selection
                    try:
                        choice = input(f"Select index [0-{num_matches-1}] (or press Enter to cancel): ").strip()
                        if choice == "": print("   Selection cancelled."); return None, None, None
                        idx = int(choice)
                        if 0 <= idx < num_matches:
                            selected_match = matching_windows[idx]
                            pid = selected_match["pid"]
                            window_element = selected_match["element"]
                            app_label = f"{app_label} ({selected_match['browser_name']})" # Refine label
                            print(f"✅ Selected window: '{selected_match['title']}' (PID: {pid})")
                            return window_element, app_label, pid # Return selected WINDOW element
                        else:
                            print("Invalid index.")
                    except ValueError:
                        print("Invalid input. Please enter a number.")
    # --- End of specific_pids_containing_window method ---

    # --- Method: Manual search for window title across ALL running apps ---
    elif find_method == "manual_window_search_only":
         while True: # Loop for searching with different fragments
             fragment = input("Enter window title fragment to search ALL apps (leave empty to cancel): ").strip()
             if not fragment: print("   Window search cancelled."); return None, None, None

             # search_window_titles_across_apps returns list of (pid, title, cmd, element)
             matches = search_window_titles_across_apps(fragment, args)
             if not matches:
                 print(f"❌ No windows found containing '{fragment}'.")
                 retry_frag = input("Search again with a different fragment? (y/N): ").strip().lower()
                 if not retry_frag.startswith('y'): return None, None, None
                 continue # Go back to fragment input

             # Display matches
             print("\nFound window title matches:")
             for i, (mpid, title, cmd, _) in enumerate(matches):
                 # Try to get a cleaner command name
                 cmd_name_short = os.path.basename(cmd) if cmd else "UnknownCmd"
                 print(f"[{i: >2}] PID: {mpid: <6} | Title: '{title}' | Cmd: {cmd_name_short}")

             # Loop for selection from the found matches
             while True:
                 try:
                     sel = input(f"Select index [0-{len(matches)-1}] (or press Enter to search again): ").strip()
                     if sel == "": break # Break inner loop to re-prompt for fragment
                     idx = int(sel)
                     if 0 <= idx < len(matches):
                         pid_sel, title_sel, cmd_name_sel, found_window = matches[idx]
                         if found_window:
                             print(f"✅ Selected initial window element: '{title_sel}'.")
                             # Use command name for a generic label if possible
                             context_label = os.path.basename(cmd_name_sel).split('.')[0] if '.' in os.path.basename(cmd_name_sel) else os.path.basename(cmd_name_sel)
                             if not context_label: context_label = app_label # Fallback
                             return found_window, context_label, pid_sel # Return WINDOW element
                         else:
                              # This shouldn't happen if search_window_titles_across_apps worked correctly
                              print("❌ Internal error: Selected match did not contain a valid window element.")
                              return None, None, None # Treat as failure
                     else: print("Invalid index.")
                 except ValueError: print("Invalid input. Please enter a number.")

             # If inner loop was broken (Enter pressed), the outer loop continues to ask for fragment
    # --- End of manual_window_search_only method ---

    else:
        print(f"❌ Unknown find_method '{find_method}' configured in context '{app_label}'.")
        return None, None, None

    # Fallback return - should ideally not be reached if logic above is sound
    return None, None, None


def run_periodic_export(pid, context_config, serialization_config, depth, interval, base_export_dir, args):
    """
    Periodically finds the target UI element and exports its tree to JSON.
    Handles re-finding the application/window element and searching within multiple
    candidate windows if necessary.

    Args:
        pid: The PID of the target application process.
        context_config: The dictionary for the selected App Context.
        serialization_config: The dictionary for the selected Serialization Preset.
        depth: Maximum serialization depth.
        interval: Time in seconds between export cycles.
        base_export_dir: The base directory to save export files.
        args: Command line arguments (for verbose_print).
    """
    if not pid or not context_config or not serialization_config:
        print("❌ Cannot run periodic export: Missing PID, context, or serialization config.")
        return

    app_label = context_config.get("app_label", "UnknownLoop")
    target_criteria = serialization_config.get("target_criteria") # e.g., {"role": "AXTable", ...}
    find_method = context_config.get("find_method")

    # Determine if this context requires finding a specific window each cycle
    needs_window_refind = find_method in [
        "pid_and_exact_window", "pid_and_containing_window",
        "manual_pid_then_containing_window", "manual_window_search_only",
        "specific_pids_containing_window"
    ]

    # Get window title/fragment needed for re-finding, if applicable
    window_title_or_fragment = context_config.get("window_title") or context_config.get("window_title_fragment")
    match_type = "exact" if find_method == "pid_and_exact_window" else "contains"

    verbose_print(args, f"Starting periodic export loop: PID={pid}, Interval={interval}s, TargetCriteria={target_criteria}, NeedsWindowRefind={needs_window_refind}")

    while True:
        verbose_print(args, f"--- Loop Cycle Start [{datetime.now().strftime('%H:%M:%S')}] ---")
        # 1. Check if the target process still exists
        if not pid_exists(pid):
            print(f"   ⚠️ Process with PID {pid} no longer exists. Stopping export loop.")
            break

        print(f"--- [{datetime.now().strftime('%H:%M:%S')}] Running export cycle for '{app_label}' (PID: {pid}) ---")
        element_to_export = None      # Reset element to export each cycle
        app_element_current = None    # Refreshed app element each cycle

        try:
            # 2. Re-acquire the Application Element using the PID
            verbose_print(args, f"   Re-creating application element for PID: {pid}...")
            app_element_current = AXUIElementCreateApplication(pid)
            if not app_element_current:
                print(f"   ⚠️ Failed to re-create application element for PID {pid}. Skipping cycle.")
                time.sleep(interval)
                continue
            verbose_print(args, "   Application element re-created successfully.")

            # --- Steps 3 & 4 Combined: Determine the specific element to export ---
            if needs_window_refind:
                # This context requires finding a specific window first
                if window_title_or_fragment:
                    verbose_print(args, f"   Re-finding window(s) matching '{window_title_or_fragment}' (match: {match_type})...")
                    candidate_windows = find_all_windows_by_title(app_element_current, window_title_or_fragment, args, match_type)

                    if not candidate_windows:
                        # No windows match the title this cycle
                        print(f"   ⚠️ Window matching '{window_title_or_fragment}' not found this cycle.")
                        element_to_export = None # Cannot proceed

                    elif len(candidate_windows) == 1:
                        # Exactly one window matches - the ideal case
                        win_title, win_element = candidate_windows[0]
                        print(f"   Found unique matching window: '{win_title}'")
                        if target_criteria:
                            # Search within this unique window for the target sub-element
                            print(f"   Searching within window for target: {target_criteria}...")
                            start_search_time = time.time()
                            element_to_export = find_element_by_criteria(win_element, target_criteria, args)
                            end_search_time = time.time()
                            verbose_print(args, f"   Target search took {end_search_time - start_search_time:.2f} seconds.")
                            if element_to_export:
                                print(f"   ✅ Found target sub-element within the window.")
                            else:
                                print(f"   ⚠️ Target sub-element not found within the unique window.")
                                # Keep element_to_export as None
                        else:
                            # No target criteria specified, so export the window itself
                            print("   No target criteria specified, exporting the window element.")
                            element_to_export = win_element

                    else: # Multiple candidate windows found
                        print(f"   ℹ️ Found {len(candidate_windows)} candidate windows matching '{window_title_or_fragment}'.")
                        if target_criteria:
                            # Search within each candidate for the target criteria
                            print(f"   Searching {len(candidate_windows)} candidates for target: {target_criteria}...")
                            found_target_in_any_candidate = False
                            for idx, (win_title, win_element) in enumerate(candidate_windows):
                                verbose_print(args, f"      Searching Candidate [{idx}]: '{win_title}'...")
                                start_search_time = time.time()
                                # Check if this candidate contains the target
                                found_target = find_element_by_criteria(win_element, target_criteria, args)
                                end_search_time = time.time()
                                verbose_print(args, f"      Search in candidate [{idx}] took {end_search_time - start_search_time:.2f} seconds.")
                                if found_target:
                                    print(f"   ✅ Found target element in Candidate [{idx}]: '{win_title}'. Using this for export.")
                                    element_to_export = found_target # Export the target element found
                                    found_target_in_any_candidate = True
                                    break # Stop searching once found in the first candidate
                            if not found_target_in_any_candidate:
                                print(f"   ⚠️ Target element ({target_criteria}) not found in any of the {len(candidate_windows)} candidate windows.")
                                element_to_export = None # Ensure it's None if not found
                        else:
                            # Multiple windows, but no target criteria - export the first window found
                            print(f"   ⚠️ No target criteria specified. Exporting the first candidate window found: '{candidate_windows[0][0]}'")
                            element_to_export = candidate_windows[0][1] # Export the container

                else: # Needs window refind, but no title/fragment provided (config error)
                     print(f"   ⚠️ Configuration error: Window re-find needed, but no window title/fragment specified. Cannot determine element.")
                     element_to_export = None

            else: # Not needs_window_refind (e.g., context targets the app element directly)
                 container_element = app_element_current
                 print("   Using Application element as base for export/search.")
                 if target_criteria:
                     # Search within the main app element
                     print(f"   Searching within application element for target: {target_criteria}...")
                     start_search_time = time.time()
                     element_to_export = find_element_by_criteria(container_element, target_criteria, args)
                     end_search_time = time.time()
                     verbose_print(args, f"   Target search took {end_search_time - start_search_time:.2f} seconds.")
                     if element_to_export:
                          print(f"   ✅ Found target sub-element within the application element.")
                     else:
                          print(f"   ⚠️ Target sub-element not found within the application element.")
                          # Keep element_to_export as None
                 else:
                     # No target criteria, export the app element itself
                     print("   No target criteria specified, exporting the application element.")
                     element_to_export = container_element

            # --- Step 5: Perform the Export ---
            if element_to_export:
                 export_to_json(element_to_export, base_export_dir, depth, args, serialization_config)
            else:
                 # This occurs if no matching window/app was found, or if target_criteria were specified but not met
                 print("   Skipping export this cycle (required element not found or identified).")

        except Exception as e:
            print(f"⚠️ An unexpected error occurred during the export cycle: {e}")
            if args.verbose: # Print traceback only if verbose
                 traceback.print_exc()
            # Allow the loop to continue to the next interval

        # --- Step 6: Wait for the next interval ---
        print(f"--- Waiting {interval} seconds ---")
        time.sleep(interval)
        # --- End of Loop Cycle ---

    verbose_print(args, "Exited periodic export loop.")


# --- Main Execution ---

def main():
    # --- ASCII Art Intro ---
    print(r"""
               @@@@@@                     @@
             @@@@@@@@@@                  @@@@
            @@@      @@@                 @@@@@
            @@        @@                @@  @@
           @@@        @@             @@@      @@@
           @@@        @@             @@        @@
           @@@        @@            @@@        @@@
     @@     @@        @@            @@          @@
     @@     @@@      @@@           @@            @@
     @@@     @@@@@@@@@@           @@@            @@@
      @@@      @@@@@@             @@@@@@@@@@@@@@@@@@
       @@@@                      @@@@@@@@@@@@@@@@@@@@
         @@@@@                  @@@                @@
            @@@@@@@@@@@         @@                  @@
                 @@            @@@                  @@@
                 @@            @@                    @@
                 @@           @@                     @@@
                 @@           @@                      @@

       macOS Accessibility Export and Transcript Conversion Tools
    """)

    # --- Argument Parsing for Verbosity ---
    parser = argparse.ArgumentParser(description="macOS Accessibility Tree Export Tool")
    parser.add_argument('-v', '--verbose', action='store_true', help='Enable verbose/debug output')
    args = parser.parse_args()

    if args.verbose:
        print("--- Verbose Mode Enabled ---")

    # --- Accessibility Permissions Check ---
    if not AXIsProcessTrusted():
        print("\n" + "="*60)
        print(" Accessibility Permissions Required ".center(60, "="))
        print("\nThis script requires Accessibility permissions to inspect UI elements.")
        print("Please grant permissions in:")
        print("  System Settings > Privacy & Security > Accessibility")
        print("\nEnsure the terminal application you are running this script from")
        print("(e.g., Terminal.app, iTerm.app) is listed and checked.")
        print("\nYou may need to restart the terminal or the script after granting.")
        print("="*60 + "\n")
        return # Exit if permissions are not granted

    # --- State Variables ---
    initial_element = None # The App or Window element found in Stage 1
    app_label = "UnknownApp" # User-friendly label, potentially refined
    pid = None               # PID of the target process
    chosen_context_name = None # Name of the selected context preset/manual option
    context_config = None    # Configuration dictionary for the chosen context

    # --- Stage 1: Select App Context & Find Initial Element ---
    while initial_element is None: # Loop until an element is found or user quits
        print("\n--- Stage 1: Select Application Context ---")

        # Separate contexts into presets and manual for display
        preset_keys = []
        manual_keys = []
        for key, config in APP_CONTEXTS.items():
            find_method = config.get("find_method", "")
            if find_method.startswith("manual_"):
                manual_keys.append(key)
            else:
                preset_keys.append(key)

        # Build the ordered list for selection and display menu
        all_options_keys = [] # Stores the keys in the order they are presented
        current_index = 0

        print("--- Presets ---")
        if preset_keys:
            for key in sorted(preset_keys): # Sort for consistent display order
                print(f"  [{current_index}] {key}")
                all_options_keys.append(key)
                current_index += 1
        else:
            print("  (No presets defined)")

        print("\n--- Manual Options ---")
        if manual_keys:
            for key in sorted(manual_keys): # Sort for consistent display order
                print(f"  [{current_index}] {key}")
                all_options_keys.append(key)
                current_index += 1
        else:
             print("  (No manual options defined)")

        print("--------------------")
        print("  [q] Quit")

        choice = input("\nEnter your choice: ").strip().lower()
        if choice == 'q':
            print("Exiting.")
            return

        try:
            chosen_index = int(choice)
            if 0 <= chosen_index < len(all_options_keys):
                chosen_context_name = all_options_keys[chosen_index]
                # Get a fresh copy of the config to avoid modifications affecting subsequent runs
                context_config = copy.deepcopy(APP_CONTEXTS[chosen_context_name])
                print(f"\nSelected Context: {chosen_context_name}. Attempting to find element...")

                # This function attempts to find the element based on the config
                initial_element, refined_app_label, pid = get_initial_element_from_context(context_config, args)

                if initial_element is None:
                     print("❌ Failed to find the initial App/Window element for this context, or operation cancelled.")
                     # Reset potentially modified variables if find failed/cancelled
                     pid = None
                     app_label = "UnknownApp"
                     context_config = None
                     # Loop will continue to re-prompt for context selection
                else:
                     # Success! Store the results and update the config with the refined label
                     app_label = refined_app_label # Use the potentially refined label (e.g., includes browser)
                     context_config['app_label'] = app_label # Update the config dict for potential later use
                     verbose_print(args, f"Successfully found initial element. PID={pid}, AppLabel='{app_label}'")
                     # Break the while loop as we found the element and can proceed
                     break
            else:
                print("❌ Invalid choice index.")
        except ValueError:
             print("❌ Invalid input. Please enter a number or 'q'.")
        # Loop continues if element wasn't found or input was invalid

    # Exit if loop terminated because initial_element is still None (should only happen via 'q')
    if initial_element is None:
        print("Exiting.")
        return

    # --- Stage 2: Select Serialization Preset ---
    selected_serialization_config = None
    serialization_name = ""
    # Use the potentially updated context_config which has the refined app_label
    default_preset_name = context_config.get("default_serialization_preset_name")

    # Try to use the default preset if specified and valid
    if default_preset_name and default_preset_name in SERIALIZATION_PRESETS:
        serialization_name = default_preset_name
        selected_serialization_config = copy.deepcopy(SERIALIZATION_PRESETS[serialization_name])
        print("\n--- Stage 2: Serialization Target ---")
        print(f"✅ Automatically selected serialization based on context: {serialization_name}")
        verbose_print(args, f"Using default serialization preset: {serialization_name}")
    else:
        # Prompt user if no valid default is set
        verbose_print(args, f"No valid default serialization preset found ({default_preset_name}). Prompting user.")
        print("\n--- Stage 2: Select Serialization Target ---")
        serialization_keys = sorted(list(SERIALIZATION_PRESETS.keys())) # Sort for consistent order
        print("Available Serialization Presets:")
        for i, name in enumerate(serialization_keys):
            desc = SERIALIZATION_PRESETS[name].get('description', '(No description)')
            print(f"  [{i}] {name} - {desc}")
        print("  [q] Quit")

        while selected_serialization_config is None: # Loop until valid selection or quit
            choice = input("\nEnter your choice: ").strip().lower()
            if choice == 'q': print("Exiting."); return

            try:
                chosen_index = int(choice)
                if 0 <= chosen_index < len(serialization_keys):
                    serialization_name = serialization_keys[chosen_index]
                    selected_serialization_config = copy.deepcopy(SERIALIZATION_PRESETS[serialization_name])
                    print(f"\nSelected Serialization: {serialization_name}")
                    verbose_print(args, f"User selected serialization preset: {serialization_name}")
                    break # Exit selection loop
                else:
                    print("❌ Invalid choice index.")
            except ValueError:
                 print("❌ Invalid input. Please enter a number or 'q'.")

    # Exit if serialization selection was cancelled
    if selected_serialization_config is None:
        print("Exiting.")
        return

    # --- Stage 3: Get Export Parameters (Depth and Interval) ---
    depth = None
    interval = None

    print("\n--- Stage 3: Set Export Parameters ---")

    # Determine Serialization Depth
    default_depth_value = selected_serialization_config.get("default_depth")
    if isinstance(default_depth_value, int) and default_depth_value > 0:
        depth = default_depth_value
        print(f"Using default serialization depth from preset: {depth}")
        verbose_print(args, f"Using default depth: {depth}")
    else:
        # Prompt user for depth if no valid default
        while True:
            try:
                depth_input = input(f"Enter serialization depth (e.g., 25, default is 25): ").strip()
                if not depth_input: depth = 25; print(f"   Using default depth: {depth}")
                else: depth = int(depth_input)

                if depth <= 0: print("❌ Depth must be a positive integer.")
                else: break # Valid depth entered
            except ValueError: print("❌ Invalid number for depth.")
        verbose_print(args, f"User set depth: {depth}")

    # Determine Export Interval
    default_interval_value = selected_serialization_config.get("default_interval")
    if isinstance(default_interval_value, int) and default_interval_value >= 0:
        interval = default_interval_value
        interval_desc = f"{interval} seconds" if interval > 0 else "Single export"
        print(f"Using default export interval from preset: {interval_desc}")
        verbose_print(args, f"Using default interval: {interval}")
    else:
        # Prompt user for interval if no valid default
         while True:
            try:
                interval_input = input("Export every N seconds (0 for single export) [default 30]: ").strip()
                if not interval_input: interval = 30; print(f"   Using default interval: {interval} seconds")
                else: interval = int(interval_input)

                if interval < 0: print("❌ Interval cannot be negative.")
                else: break # Valid interval entered
            except ValueError: print("❌ Invalid number for interval.")
         verbose_print(args, f"User set interval: {interval}")


    # --- Stage 4: Create Base Directory and Execute Export ---
    parent_timestamp = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
    # Create a filesystem-safe name from the (potentially refined) app label
    base_export_dir = os.path.join(BASE_EXPORT_DIR_NAME, f"export_{parent_timestamp}")

    try:
        os.makedirs(base_export_dir, exist_ok=True)
        print(f"\n--- Creating Base Export Directory ---")
        print(f"All exports for this session will be saved under: ./{base_export_dir}")
    except OSError as e:
        print(f"❌ CRITICAL ERROR: Could not create base export directory '{base_export_dir}': {e}")
        return # Cannot proceed without export directory

    # --- Execute Export (Single or Periodic) ---
    print("\n--- Starting Export ---")
    # Use chosen_context_name for user display, but context_config['app_label'] for details
    print(f"App Context: {chosen_context_name} (Label: {context_config['app_label']}, PID: {pid})")
    print(f"Serialization Target: {serialization_name}")
    print(f"Max Depth: {depth}")
    print(f"Interval: {'Single Export' if interval <= 0 else f'Every {interval} seconds'}")

    target_criteria = selected_serialization_config.get("target_criteria")

    if interval <= 0:
        # --- Single Export ---
        print("\nPerforming a single export...")
        element_to_export = initial_element # Start with the element found in Stage 1

        # If a specific sub-element is desired, try to find it
        if target_criteria:
             verbose_print(args, f"Searching within initial element for target: {target_criteria}...")
             found_sub_element = find_element_by_criteria(initial_element, target_criteria, args)
             if found_sub_element:
                  print("✅ Found target sub-element for single export.")
                  element_to_export = found_sub_element # Export the specific sub-element
             else:
                  print(f"⚠️ Target sub-element ({target_criteria}) not found. Exporting the initial element instead.")
                  # Keep element_to_export as initial_element (the container)
        else:
             verbose_print(args, "Exporting the initial element found by App Context (no target criteria).")

        # Perform the actual export
        if element_to_export:
             export_to_json(element_to_export, base_export_dir, depth, args, selected_serialization_config)
             print("✅ Single export complete.")
        else:
             # This case should ideally not happen if initial_element was valid
             print("❌ Error: No valid element was identified for single export.")

    else:
        # --- Periodic Export Loop ---
        print(f"\nStarting periodic export every {interval} seconds. Press Ctrl+C to stop.")
        # Ensure all necessary components are available before starting the loop
        if pid and context_config and selected_serialization_config:
             try:
                  # Pass the necessary configs and parameters to the loop function
                  run_periodic_export(pid, context_config, selected_serialization_config,
                                      depth, interval, base_export_dir, args)
             except KeyboardInterrupt:
                  print("\n🛑 Ctrl+C detected. Stopping export loop.")
             except Exception as e:
                  print(f"\n❌ An unexpected error occurred during the periodic export loop: {e}")
                  if args.verbose:
                       traceback.print_exc()
             finally:
                  print("Export loop finished.")
        else:
             # Should not happen if stages completed correctly, but safeguard
             print("❌ Error: Missing necessary information (PID, context, or serialization config) to start periodic export.")
             verbose_print(args, f"   Debug Info: PID={pid}, Context Config Present={context_config is not None}, Ser Config Present={selected_serialization_config is not None}")

    verbose_print(args, "--- End of main function ---")


if __name__ == "__main__":
    main()
