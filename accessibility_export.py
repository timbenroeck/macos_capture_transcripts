# Imports
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
)
import copy
import signal # For checking PID existence
import argparse # For verbosity flag
import traceback # For detailed error printing

# --- Attribute Strings ---
AX_ROLE = "AXRole"
AX_SUBROLE = "AXSubrole"
AX_DESCRIPTION = "AXDescription"
AX_HELP = "AXHelp"
AX_LABEL_VALUE = "AXLabelValue"
AX_VALUE = kAXValueAttribute
AX_TITLE = kAXTitleAttribute
AX_CHILDREN = kAXChildrenAttribute

# --- App Context Presets (Added "Teams in Browser") ---
APP_CONTEXTS = {
    "Zoom Transcript Window": {
        "app_label": "Zoom",
        "find_method": "pid_and_exact_window",
        "cmd_path": "/Applications/zoom.us.app/Contents/MacOS/zoom.us",
        "window_title": "Transcript",
        "default_serialization_preset_name": "Zoom Transcript Table"
    },
    "Teams in Browser (Chrome/Prisma)": {
        "app_label": "TeamsInBrowser", # Will be refined based on actual browser found
        "find_method": "specific_pids_containing_window",
        "cmd_paths": [ # List of target browser executables
            "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
            "/Applications/Prisma Access Browser.app/Contents/MacOS/Prisma Access Browser"
        ],
        "window_title_fragment": "teams", # Case-insensitive search term
        "default_serialization_preset_name": "Teams Live Captions Group"
    },
    # --- Manual Options ---
     "Enter PID Directly": {
         "app_label": "ManualPIDEntry",
         "find_method": "manual_pid_direct",
     },
     "Manual PID Search": {
         "app_label": "ManualPIDSearch",
         "find_method": "manual_pid_search_only",
     },
      "Manual Window Search": {
         "app_label": "ManualWindowSearch",
         "find_method": "manual_window_search_only",
     }
}

# --- Serialization Presets (Unchanged) ---
SERIALIZATION_PRESETS = {
    "Full Element Found": {
        "description": "Serialize the entire App/Window element found by the context.",
        "target_criteria": None,
        "export_suffix": "FullElement", # Kept for logic, but not used in filename
        "default_depth": None,
        "default_interval": None
    },
    "Zoom Transcript Table": {
        "description": "Finds and serializes the AXTable containing Zoom transcript text.",
        "target_criteria": {"role": "AXTable", "description": "Transcript list"},
        "export_suffix": "ZoomTranscriptTable", # Kept for logic, but not used in filename
        "applicable_contexts": ["Zoom Transcript Window"], # Could add TeamsInBrowser here if needed, but default handles it
        "default_depth": 25,
        "default_interval": 30
    },
    "Teams Live Captions Group": {
        "description": "Finds and serializes the AXGroup containing Teams Live Captions.",
        "target_criteria": {"role": "AXGroup", "description": "Live Captions"},
        "export_suffix": "TeamsLiveCaptions", # Kept for logic, but not used in filename
        # applicable_contexts not strictly necessary if relying on default_serialization_preset_name
        "default_depth": 25,
        "default_interval": 30
    },
}

# --- Helper Functions (Mostly unchanged, except print statements and find_all_windows_by_title) ---

def verbose_print(args, *print_args, **print_kwargs):
    """Prints only if the verbose flag is set."""
    if args.verbose:
        print(*print_args, **print_kwargs)

def get_attribute(element, attr):
    """Safely retrieves an accessibility attribute value."""
    result, value = AXUIElementCopyAttributeValue(element, attr, None)
    if result == 0: return value
    return None

def serialize_ax_element(element, depth=0, max_depth=25):
    """Recursively serializes an accessibility element and its children."""
    if not element or depth > max_depth:
        return ({"error": f"<Max depth {max_depth} reached>"} if depth > max_depth else None), 0
    data = {}
    row_count = 0
    attributes_to_check = {
        "role": AX_ROLE, "subrole": AX_SUBROLE, "title": AX_TITLE,
        "value": AX_VALUE, "description": AX_DESCRIPTION, "help": AX_HELP,
        "label": AX_LABEL_VALUE,
    }
    for key, attr_name in attributes_to_check.items():
        try:
            value = get_attribute(element, attr_name)
            if value is not None:
                if isinstance(value, str) and value: data[key] = value
                elif not isinstance(value, str):
                     if isinstance(value, (int, float, bool)): data[key] = value
                     else:
                         try: data[key] = repr(value)
                         except Exception: data[key] = "<Unrepresentable CFType>"
                if key == "role" and value == "AXRow": row_count += 1
        except Exception as e: data[f"{key}_error"] = f"Error fetching {attr_name}: {e}"
    try:
        children = get_attribute(element, AX_CHILDREN)
        if children:
            children_data = []
            for child in children:
                child_data, child_rows = serialize_ax_element(child, depth + 1, max_depth)
                if child_data: children_data.append(child_data)
                row_count += child_rows
            if children_data: data["children"] = children_data
    except Exception as e: data["children_error"] = f"Error fetching children: {e}"
    has_real_data = any(not k.endswith("_error") for k in data)
    return data if has_real_data else None, row_count

def find_element_by_criteria(start_element, criteria, args, current_depth=0, max_search_depth=15):
    """Recursively searches for the first element matching all criteria (DFS)."""
    if not start_element or current_depth > max_search_depth: return None
    match = True
    current_criteria = copy.deepcopy(criteria)
    element_attrs = {}
    for key, _ in current_criteria.items():
        attr_name = None
        if key == "role": attr_name = AX_ROLE
        elif key == "subrole": attr_name = AX_SUBROLE
        elif key == "title": attr_name = AX_TITLE
        elif key == "description": attr_name = AX_DESCRIPTION
        elif key == "value": attr_name = AX_VALUE
        elif key == "help": attr_name = AX_HELP
        elif key == "label": attr_name = AX_LABEL_VALUE
        if attr_name: element_attrs[key] = get_attribute(start_element, attr_name)
        else: print(f"Warning: Unknown criteria key '{key}'"); element_attrs[key] = None
    for key, expected_value in current_criteria.items():
        actual_value = element_attrs.get(key)
        if actual_value != expected_value: match = False; break
    if match: return start_element
    children = get_attribute(start_element, AX_CHILDREN)
    if children:
        for child in children:
            found_element = find_element_by_criteria(child, criteria, args, current_depth + 1, max_search_depth)
            if found_element: return found_element
    return None

def find_process_by_cmd(cmd_path, args):
    """Finds a PID by the exact command path."""
    basename = os.path.basename(cmd_path)
    verbose_print(args, f"   Running: pgrep -f '{cmd_path}'")
    result = subprocess.run(["pgrep", "-f", cmd_path], capture_output=True, text=True)
    if result.returncode != 0 or not result.stdout.strip():
         verbose_print(args, f"   '{cmd_path}' not found directly, trying basename: pgrep -fl '{basename}'")
         result_base = subprocess.run(["pgrep", "-fl", basename], capture_output=True, text=True)
         if result_base.returncode == 0 and result_base.stdout.strip():
             verbose_print(args, f"   Found potential matches by basename:\n{result_base.stdout.strip()}")
             for line in result_base.stdout.strip().splitlines():
                 try:
                     pid_str, potential_cmd = line.strip().split(" ", 1)
                     # Ensure it's the target process, not just the pgrep command itself finding its path
                     if cmd_path in potential_cmd and 'pgrep' not in potential_cmd.lower():
                         verbose_print(args, f"   Found matching full path in line: {line.strip()}")
                         return int(pid_str)
                 except ValueError: continue
         verbose_print(args, f"   No exact match found by full path or basename for '{cmd_path}'.")
         return None
    try:
        # Handle multiple PIDs returned, only return if it matches the command path closely
        pids_found = []
        for pid_str in result.stdout.strip().splitlines():
            try:
                pid = int(pid_str)
                # Verify the command line for the PID
                ps_result = subprocess.run(["ps", "-o", "command=", "-p", str(pid)], capture_output=True, text=True)
                if ps_result.returncode == 0 and cmd_path in ps_result.stdout:
                    pids_found.append(pid)
            except (ValueError, subprocess.CalledProcessError):
                continue
        if pids_found:
             pid = pids_found[0] # Take the first one usually
             verbose_print(args, f"   Found and verified PID {pid} directly for '{cmd_path}'.")
             return pid
        else:
             verbose_print(args, f"   pgrep found PIDs, but verification failed for '{cmd_path}'.")
             return None

    except (ValueError, IndexError):
        verbose_print(args, f"   Error parsing pgrep output or no process found for '{cmd_path}'.")
        return None

def choose_pid_manually(query, args):
    """Allows manual selection of a PID from processes matching a query."""
    verbose_print(args, f"   Running: pgrep -fl '{query}'")
    result = subprocess.run(["pgrep", "-fl", query], capture_output=True, text=True)
    processes = []
    if result.stdout:
        for line in result.stdout.strip().splitlines():
            try:
                pid, cmd = line.strip().split(" ", 1)
                if 'pgrep' not in cmd.lower(): # Filter out the pgrep command itself
                    processes.append((int(pid), cmd))
            except ValueError: continue
    if not processes: print(f"❌ No running processes found matching '{query}'."); return None
    print("\nMatching processes:")
    for idx, (pid, cmd) in enumerate(processes): print(f"[{idx: >2}] PID: {pid: <6} | CMD: {cmd}")
    while True:
        try:
            choice = input(f"Select index [0-{len(processes)-1}] (or press Enter to cancel): ").strip()
            if choice == "": return None
            idx = int(choice)
            if 0 <= idx < len(processes): return processes[idx][0]
            else: print("Invalid index.")
        except ValueError: print("Invalid input. Please enter a number.")

def search_window_titles_across_apps(title_fragment, args):
    """Searches all running applications for windows containing a title fragment."""
    verbose_print(args, "   Running: ps -axo pid=,comm=")
    result = subprocess.run(["ps", "-axo", "pid=,comm="], capture_output=True, text=True)
    matches = []
    processed_pids = set()
    for line in result.stdout.strip().splitlines():
        try:
            pid_str, cmd = line.strip().split(maxsplit=1)
            pid = int(pid_str)
            if pid in processed_pids: continue
            verbose_print(args, f"   Checking PID {pid} ({cmd})...")
            app = AXUIElementCreateApplication(pid)
            processed_pids.add(pid)
            if not app:
                verbose_print(args, f"   Skipping PID {pid}: Could not create AXUIElement.")
                continue

            result_code, windows = AXUIElementCopyAttributeValue(app, AX_CHILDREN, None)
            if result_code == 0 and windows:
                verbose_print(args, f"   Found {len(windows)} potential children for PID {pid}.")
                for i, win in enumerate(windows):
                    title = get_attribute(win, AX_TITLE)
                    if isinstance(title, str) and title_fragment.lower() in title.lower():
                         verbose_print(args, f"      Match found: Window {i}, Title: '{title}'")
                         matches.append((pid, title.strip(), cmd, win)) # Add window element
                    # else:
                    #     verbose_print(args, f"      Window {i}, Title: '{title}' (No match)")
            # else:
            #     verbose_print(args, f"   No windows found or error for PID {pid} (Result code: {result_code}).")

        except Exception as e:
            verbose_print(args, f"   Error processing PID {pid_str}: {e}")
            continue
    verbose_print(args, f"   Found {len(matches)} total matches across all PIDs.")
    return matches

def find_window_by_title(app, title_query, args, match_type="contains"):
    """Finds the FIRST window within a specific app element matching the title."""
    if not app: return None # Safety check
    result, windows = AXUIElementCopyAttributeValue(app, AX_CHILDREN, None)
    if result != 0 or not windows:
        verbose_print(args, f"   find_window_by_title: No windows found or error getting children (Result: {result}).")
        return None
    verbose_print(args, f"   find_window_by_title: Searching {len(windows)} windows for title '{title_query}' (match: {match_type})...")
    for i, win in enumerate(windows):
        title = get_attribute(win, AX_TITLE)
        if isinstance(title, str):
            title_str = title.strip()
            matches = False
            if match_type == "exact": matches = (title_str == title_query)
            elif match_type == "contains": matches = (title_query.lower() in title_str.lower())
            else: matches = (title_query.lower() in title_str.lower()) # Default to contains

            # verbose_print(args, f"      Checking window {i}: Title='{title_str}', Matches={matches}")
            if matches:
                verbose_print(args, f"      Found first match: Window {i}")
                return win # Return the first match
    verbose_print(args, f"   find_window_by_title: No matching window found.")
    return None

# --- NEW Helper Function ---
def find_all_windows_by_title(app, title_query, args, match_type="contains"):
    """Finds ALL windows within a specific app element matching the title."""
    matches = []
    if not app: return matches
    result, windows = AXUIElementCopyAttributeValue(app, AX_CHILDREN, None)
    if result != 0 or not windows:
        verbose_print(args, f"   find_all_windows_by_title: No windows or error (Result: {result}).")
        return matches
    verbose_print(args, f"   find_all_windows_by_title: Searching {len(windows)} windows for title '{title_query}' (match: {match_type})...")
    for i, win in enumerate(windows):
        title = get_attribute(win, AX_TITLE)
        if isinstance(title, str):
            title_str = title.strip()
            does_match = False
            if match_type == "exact": does_match = (title_str == title_query)
            else: does_match = (title_query.lower() in title_str.lower()) # Default contains

            if does_match:
                verbose_print(args, f"      Found match: Window {i}, Title: '{title_str}'")
                matches.append((title_str, win)) # Store title and element
    verbose_print(args, f"   find_all_windows_by_title: Found {len(matches)} matching windows.")
    return matches


def pid_exists(pid):
    """ Check For the existence of a unix pid. """
    if pid is None or pid <= 0: return False
    try:
        os.kill(pid, 0)
    except OSError:
        return False
    else:
        return True

# --- Refactored Export Function (Unchanged) ---
def export_to_json(element, base_export_dir, depth, args):
    """Serializes the element and saves it to a timestamped JSON file within the base export directory."""
    if not element:
        print(f"❌ Cannot export, element is None.")
        return
    element_role_or_title = get_attribute(element, AX_ROLE) or get_attribute(element, AX_TITLE) or "UnknownElement"
    print(f"⏳ Serializing element tree starting from '{element_role_or_title}' (max depth: {depth})...")

    start_time = time.time()
    serialized, row_count = serialize_ax_element(element, max_depth=depth)
    end_time = time.time()
    print(f"⏱️ Serialization finished in {end_time - start_time:.2f} seconds.")

    if not serialized:
        print("❓ Serialization resulted in empty data. Nothing to save.")
        return

    # Generate timestamp and filename according to new format
    timestamp = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
    filename = f"export_{timestamp}.json"
    filepath = os.path.join(base_export_dir, filename)

    try:
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(serialized, f, indent=2, ensure_ascii=False)
        print(f"✅ Export saved to: {filepath} ({row_count} AXRow elements found)")
    except Exception as e:
        print(f"❌ Error saving JSON file '{filepath}': {e}")


# --- Main Workflow Functions (Passing `args` and `base_export_dir`) ---

def get_initial_element_from_context(context_config, args):
    """ Finds the initial App or Window element based on the App Context config. """
    find_method = context_config.get("find_method")
    app_label = context_config.get("app_label", "UnknownContext")
    pid = None
    app_element = None
    window_element = None

    verbose_print(args, f"--- Finding Element using Method: {find_method} ---")

    # --- Methods that find PID first ---
    if find_method in ["pid_and_exact_window", "pid_and_containing_window", "manual_pid_search_only", "manual_pid_then_containing_window", "manual_pid_direct"]:
        if find_method == "manual_pid_direct":
             while pid is None:
                 try:
                     pid_input = input("Enter the PID number: ").strip()
                     pid_val = int(pid_input)
                     if pid_exists(pid_val):
                         pid = pid_val
                         print(f"   Using PID: {pid}")
                     else:
                         print(f"❌ Process with PID {pid_val} not found.")
                 except ValueError:
                     print("❌ Invalid PID. Please enter a number.")
                 except Exception as e:
                     print(f"❌ Error checking PID: {e}")
                 if pid is None:
                     retry = input("Try entering PID again? (Y/n): ").strip().lower()
                     if retry.startswith('n'): return None, None, None

        elif find_method == "manual_pid_search_only":
            while pid is None:
                 query = input(f"Enter query to find PID: ").strip()
                 if not query: print("Query cannot be empty."); continue
                 pid = choose_pid_manually(query, args)
                 if pid is None: return None, None, None # User cancelled

        elif find_method == "manual_pid_then_containing_window":
             search_hint = context_config.get("pid_search_hint", "application name")
             while pid is None:
                 query = input(f"Enter query to find PID (hint: {search_hint}): ").strip()
                 if not query: print("Query cannot be empty."); continue
                 pid = choose_pid_manually(query, args)
                 if pid is None: return None, None, None # User cancelled PID selection

        elif find_method in ["pid_and_exact_window", "pid_and_containing_window"]:
            cmd_path = context_config.get("cmd_path")
            if not cmd_path: print(f"❌ Context '{app_label}' misconfigured: Needs 'cmd_path'."); return None, None, None
            print(f"   Finding PID for {cmd_path}...")
            pid = find_process_by_cmd(cmd_path, args)
            if not pid: print(f"❌ PID not found for {cmd_path}. Is the app running?"); return None, None, None

        # --- Common logic after getting PID ---
        if pid:
            print(f"   Creating app element for PID: {pid}...")
            app_element = AXUIElementCreateApplication(pid)
            if not app_element: print(f"❌ Failed to create app element for PID {pid}. Check permissions."); return None, None, None
            verbose_print(args, "   App element created successfully.")

            # If method involves finding a window
            if find_method in ["pid_and_exact_window", "pid_and_containing_window", "manual_pid_then_containing_window"]:
                window_title = ""
                match_type = "contains" # default
                if find_method == "manual_pid_then_containing_window":
                    window_title_hint = context_config.get("window_title_hint", "partial window title")
                    while window_element is None:
                         title_query = input(f"Enter window title fragment (hint: {window_title_hint}): ").strip()
                         if not title_query: print("Title fragment cannot be empty."); continue
                         window_element = find_window_by_title(app_element, title_query, args, "contains")
                         if window_element is None:
                              try_again = input("Window not found. Try again? (y/N): ").strip().lower()
                              if not try_again.startswith('y'): return None, None, None # User gave up
                         else:
                              window_title = title_query # Store what we searched for
                              break # Found window
                else: # pid_and_exact_window or pid_and_containing_window
                     window_title = context_config.get("window_title")
                     if not window_title: print(f"❌ Context '{app_label}' misconfigured: Needs 'window_title'."); return None, None, None
                     match_type = "exact" if find_method == "pid_and_exact_window" else "contains"
                     print(f"   Finding window '{window_title}' (match: {match_type})...")
                     window_element = find_window_by_title(app_element, window_title, args, match_type)

                if window_element:
                    print(f"✅ Found initial window element.")
                    return window_element, app_label, pid # Return element, label, AND PID
                else:
                    print(f"❌ Window '{window_title}' not found for PID {pid}.")
                    return None, None, None
            else: # manual_pid_direct or manual_pid_search_only - return app element
                print(f"✅ Found initial app element.")
                return app_element, app_label, pid # Return element, label, AND PID

    # --- Method for specific PIDs and contained window ---
    elif find_method == "specific_pids_containing_window":
        cmd_paths = context_config.get("cmd_paths")
        window_title_fragment = context_config.get("window_title_fragment")

        if not cmd_paths or not isinstance(cmd_paths, list):
            print(f"❌ Context '{app_label}' misconfigured: Needs 'cmd_paths' (list)."); return None, None, None
        if not window_title_fragment or not isinstance(window_title_fragment, str):
            print(f"❌ Context '{app_label}' misconfigured: Needs 'window_title_fragment' (string)."); return None, None, None

        matching_windows = [] # Store tuples of (pid, title, cmd_path, window_element)
        print(f"   Searching for windows containing '{window_title_fragment}' in specific browsers:")

        for cmd_path in cmd_paths:
            browser_name = os.path.basename(os.path.dirname(os.path.dirname(cmd_path))) # e.g., Google Chrome.app
            print(f"   - Checking {browser_name}...")
            pid = find_process_by_cmd(cmd_path, args)
            if pid:
                print(f"     Found PID: {pid}")
                app_element = AXUIElementCreateApplication(pid)
                if app_element:
                    verbose_print(args, f"     Searching windows in PID {pid} for '{window_title_fragment}'...")
                    # Use find_all_windows_by_title
                    found_in_pid = find_all_windows_by_title(app_element, window_title_fragment, args, match_type="contains")
                    for title, win_element in found_in_pid:
                        matching_windows.append({
                            "pid": pid,
                            "title": title,
                            "cmd_path": cmd_path,
                            "browser_name": browser_name,
                            "element": win_element
                        })
                    print(f"     Found {len(found_in_pid)} matching window(s) in {browser_name}.")
                else:
                    print(f"     ❌ Could not create app element for PID {pid}.")
            else:
                print(f"     Process not found or running.")

        # --- Process results ---
        if not matching_windows:
            print(f"❌ No windows containing '{window_title_fragment}' found in the specified browsers ({[os.path.basename(p) for p in cmd_paths]}).")
            return None, None, None

        elif len(matching_windows) == 1:
            match = matching_windows[0]
            pid = match["pid"]
            window_element = match["element"]
            # Refine app_label based on the browser found
            app_label = f"Teams in {match['browser_name']}"
            print(f"✅ Found unique window: '{match['title']}' (PID: {pid})")
            return window_element, app_label, pid

        else: # Multiple matches
            print(f"\n❓ Found multiple windows containing '{window_title_fragment}'. Please select one:")
            for idx, match in enumerate(matching_windows):
                print(f"  [{idx}] PID: {match['pid']:<6} | Browser: {match['browser_name']:<25} | Title: '{match['title']}'")

            while True:
                try:
                    choice = input(f"Select index [0-{len(matching_windows)-1}] (or press Enter to cancel): ").strip()
                    if choice == "": return None, None, None # User cancelled
                    idx = int(choice)
                    if 0 <= idx < len(matching_windows):
                        selected_match = matching_windows[idx]
                        pid = selected_match["pid"]
                        window_element = selected_match["element"]
                        app_label = f"Teams in {selected_match['browser_name']}"
                        print(f"✅ Selected window: '{selected_match['title']}' (PID: {pid})")
                        return window_element, app_label, pid
                    else:
                        print("Invalid index.")
                except ValueError:
                    print("Invalid input. Please enter a number.")


    # --- Method for manual window search across ALL apps ---
    elif find_method == "manual_window_search_only":
         while True:
             fragment = input("Enter window title fragment to search ALL apps: ").strip()
             if not fragment: print("Fragment cannot be empty."); continue
             # search_window_titles_across_apps returns list of (pid, title, cmd, element)
             matches = search_window_titles_across_apps(fragment, args)
             if not matches: print(f"❌ No windows found containing '{fragment}'."); continue

             print("\nFound window title matches:")
             # Display matches: pid, title, cmd
             for i, (mpid, title, cmd, _) in enumerate(matches): print(f"[{i: >2}] PID: {mpid: <6} | Title: '{title}' | CMD: {cmd}")

             while True:
                 try:
                     sel = input(f"Select index [0-{len(matches)-1}] (or press Enter to cancel): ").strip()
                     if sel == "": break
                     idx = int(sel)
                     if 0 <= idx < len(matches):
                         # Retrieve the selected match including the element
                         pid_sel, title_sel, cmd_name, found_window = matches[idx]
                         # No need to re-find, we already have the element
                         if found_window:
                             print(f"✅ Selected initial window element.")
                             context_label = os.path.basename(cmd_name).split('.')[0] if '.' in os.path.basename(cmd_name) else os.path.basename(cmd_name)
                             if not context_label: context_label = app_label # Fallback
                             return found_window, context_label, pid_sel # Return element, label, PID
                         else:
                              # This case should technically not happen if matches list is built correctly
                              print("❌ Internal error: Selected match did not contain window element.")
                              break
                     else: print("Invalid index.")
                 except ValueError: print("Invalid input. Please enter a number.")
             # If inner loop broken, re-prompt for fragment or user cancelled
             retry_frag = input("Search again with a different fragment? (Y/n): ").strip().lower()
             if retry_frag.startswith('n'): return None, None, None

    else:
        print(f"❌ Unknown find_method '{find_method}' in context '{app_label}'.")
        return None, None, None

    return None, None, None # Fallback


def run_periodic_export(pid, context_config, serialization_config, depth, interval, base_export_dir, args):
    """ Periodically finds the target element and exports it, refreshing references. """
    verbose_print(args, f"DEBUG: *** Entered run_periodic_export. PID={pid}, Interval={interval}, BaseDir={base_export_dir} ***")

    if not pid:
        print("❌ Cannot run loop without a valid PID.")
        verbose_print(args, "DEBUG: Exiting run_periodic_export early: PID is invalid.")
        return

    app_label = context_config.get("app_label", "UnknownLoop")
    # --- Logic to get window title for re-finding ---
    # If the original method found a specific window, we need its title to re-find it.
    # The specific_pids_containing_window method returns an element, but we might need
    # the original title used for selection if multiple were present.
    # For simplicity in the loop, we *assume* the title attribute of the INITIAL element
    # is sufficient for re-finding, though this might be fragile if titles change slightly.
    # A more robust approach would pass the initially selected title through.
    # Let's try getting the title from the *initial* element if possible.
    # HOWEVER, the initial element isn't passed here. We only have the context config.
    # Let's rely on the context config's window_title or fragment.

    window_title_or_fragment = context_config.get("window_title") or context_config.get("window_title_fragment")
    match_type = "exact" if context_config.get("find_method") == "pid_and_exact_window" else "contains"

    target_criteria = serialization_config.get("target_criteria")
    find_method = context_config.get("find_method")

    verbose_print(args, f"DEBUG: run_periodic_export setup complete. Title/Fragment='{window_title_or_fragment}', MatchType='{match_type}'. Entering while True loop...")

    while True:
        verbose_print(args, f"DEBUG: Top of while loop. Checking PID {pid} existence...")
        if not pid_exists(pid):
            print(f"   ⚠️ Process with PID {pid} no longer exists. Stopping loop.")
            verbose_print(args, f"DEBUG: Breaking loop because pid_exists({pid}) returned False.")
            break

        print(f"--- [{datetime.now().strftime('%H:%M:%S')}] Starting export cycle for '{app_label}' (PID: {pid}) ---")
        element_to_export = None
        container_element = None
        app_element_current = None

        try:
            # 1. Re-acquire App Element
            verbose_print(args, f"   Re-creating app element for PID: {pid}...")
            app_element_current = AXUIElementCreateApplication(pid)
            if not app_element_current:
                print(f"   ⚠️ Failed to create app element for PID {pid}. Skipping cycle.")
                verbose_print(args, f"DEBUG: Skipping cycle, AXUIElementCreateApplication failed for PID {pid}")
                time.sleep(interval)
                continue
            verbose_print(args, f"   App element re-created.")

            # 2. Determine container: Find Window (if applicable) or use App
            # If the original method resulted in a window, try to re-find it.
            # Methods targeting windows: pid_and_exact_window, pid_and_containing_window,
            # manual_pid_then_containing_window, manual_window_search_only, specific_pids_containing_window
            needs_window_refind = find_method in [
                "pid_and_exact_window", "pid_and_containing_window",
                "manual_pid_then_containing_window", "manual_window_search_only",
                "specific_pids_containing_window"
            ]

            if needs_window_refind:
                if window_title_or_fragment:
                    verbose_print(args, f"   Re-finding window '{window_title_or_fragment}' (match: {match_type})...")
                    # Use find_window_by_title (finds first match)
                    container_element = find_window_by_title(app_element_current, window_title_or_fragment, args, match_type)
                    if not container_element:
                         print(f"   ⚠️ Window matching '{window_title_or_fragment}' not found this cycle.")
                         verbose_print(args, f"DEBUG: Window not found, container_element is None.")
                    else:
                         actual_title = get_attribute(container_element, AX_TITLE) or "Unknown Title"
                         verbose_print(args, f"   Window re-found (Title: '{actual_title}').")
                else:
                     print(f"   ⚠️ Context implies window target, but no window title/fragment available for re-finding. Using App element as container.")
                     verbose_print(args, f"DEBUG: Cannot re-find window (no title/fragment). Defaulting to app element.")
                     container_element = app_element_current # Fallback to app if no title known

            else: # Methods like manual_pid_direct, manual_pid_search_only target the app initially
                verbose_print(args, "   Using App element as container.")
                container_element = app_element_current

            # 3. Find Target Element within Container (if criteria exist)
            if container_element:
                if target_criteria:
                    verbose_print(args, f"   Searching within container for: {target_criteria}...")
                    start_search_time = time.time()
                    element_to_export = find_element_by_criteria(container_element, target_criteria, args)
                    end_search_time = time.time()
                    verbose_print(args, f"   Sub-element search took {end_search_time - start_search_time:.2f} seconds.")

                    if element_to_export:
                        verbose_print(args, f"   ✅ Found target sub-element.")
                    else:
                        print(f"   ⚠️ Target sub-element ({target_criteria}) not found this cycle. Exporting the container instead.")
                        element_to_export = container_element # Fallback
                else:
                    verbose_print(args, "   Exporting the container element (no sub-element criteria).")
                    element_to_export = container_element

            # 4. Perform Export
            if element_to_export:
                 export_to_json(element_to_export, base_export_dir, depth, args)
            else:
                 print("   Skipping export this cycle (container or target not found).")
                 verbose_print(args, "DEBUG: Skipping export_to_json call because element_to_export is None.")


        except Exception as e:
            print(f"⚠️ Error during export cycle: {e}")
            verbose_print(args, f"DEBUG: Caught exception in run_periodic_export's try-except block.")
            if args.verbose: # Print traceback only if verbose
                 traceback.print_exc()

        print(f"--- Waiting {interval} seconds ---")
        verbose_print(args, f"DEBUG: Bottom of while loop. About to sleep for {interval}s.")
        time.sleep(interval)
        verbose_print(args, f"DEBUG: Woke up from sleep.")

    verbose_print(args, f"DEBUG: Exited while True loop in run_periodic_export.")


# --- Main Function (Unchanged structure, just needs to handle new context) ---
def main():
    # --- ASCII Art Intro ---
    print(r"""

               @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
          @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
        @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
      @@@@@@@@@                               @@@@@@@@@
     @@@@@@                                       @@@@@@
   @@@@                                              @@@@@
   @@@@           @@@@@@@@@@@@@@@@@@@@@@@@@           @@@@
   @@@@           @@@@@@@@@@@@@@@@@@@@@@@@@           @@@@
   @@@@            @@@@@@@@@@@@@@@@@@@@@@@            @@@@
   @@@@                                               @@@@
                  @@@@@@@@@@@@@@@@@@                  @@@@
                  @@@@@@@@@@@@@@@@@@@                 @@@@
                   @@@@@@@@@@@@@@@@                   @@@@
                                                     @@@@@
     @@@@@@@                                     @@@@@@@
      @@@@@@@@                                 @@@@@@@@
        @@@@@@@@@@@@@@@@@@         @@@@@@@@@@@@@@@@@@
           @@@@@@@@@@@@@@@@       @@@@@@@@@@@@@@@@
                  @@@@@@@@@@@   @@@@@@@@@@@
                        @@@@@@ @@@@@@
                         @@@@@@@@@@@
                          @@@@@@@@@
                            @@@@@

       macOS Accessibility Export and Transcript Conversion Tools
    """)
    # --- End ASCII Art ---

    # --- Argument Parsing for Verbosity ---
    parser = argparse.ArgumentParser(description="macOS Accessibility Tree Export Tool")
    parser.add_argument('-v', '--verbose', action='store_true', help='Enable verbose/debug output')
    args = parser.parse_args()

    if args.verbose:
        print("--- Verbose Mode Enabled ---")

    if not AXIsProcessTrusted():
        print("\n" + "="*60)
        print(" Accessibility Permissions Required ".center(60, "="))
        print("\nThis script requires Accessibility permissions to inspect UI elements.")
        print("Please grant permissions in:")
        print("  System Settings > Privacy & Security > Accessibility")
        print("\nEnsure the terminal application you are running this script from")
        print("(e.g., Terminal.app, iTerm.app) is listed and checked.")
        print("\nYou may need to restart the terminal or the script after granting.")
        print("="*60)
        return

    # print("\n=== Accessibility Tree Export Tool ===") # Replaced by ASCII art

    initial_element = None
    app_label = "UnknownApp"
    pid = None
    chosen_context_name = None
    context_config = None

    # --- Stage 1: Select App Context & Find Initial Element ---
    while initial_element is None:
        print("\n--- Stage 1: Select Application Context ---")

        # Separate contexts into presets and manual
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
            for key in preset_keys:
                print(f"  [{current_index}] {key}")
                all_options_keys.append(key)
                current_index += 1
        else:
            print("  (No presets defined)")

        print("\n--- Manual Options ---")
        if manual_keys:
            for key in manual_keys:
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

        if choice.isdigit() and 0 <= int(choice) < len(all_options_keys):
            chosen_index = int(choice)
            chosen_context_name = all_options_keys[chosen_index] # Get key from the ordered list
            context_config = APP_CONTEXTS[chosen_context_name]
            print(f"\nSelected Context: {chosen_context_name}. Attempting to find element...")
            # Pass args to the function
            initial_element, app_label, pid = get_initial_element_from_context(context_config, args)

            if initial_element is None:
                 print("❌ Failed to find the initial App/Window element for this context.")
                 pid = None # Reset PID if finding failed
                 verbose_print(args, "DEBUG: get_initial_element_from_context returned None for initial_element.")
            else:
                 verbose_print(args, f"DEBUG: Found initial element. PID={pid}, AppLabel='{app_label}'")
                 # Update context_config in memory with potentially refined app_label
                 context_config['app_label'] = app_label

        else:
            print("❌ Invalid choice.")

    # --- Stage 2: Select Serialization Preset ---
    # (No changes needed from here downwards for this request)
    selected_serialization_config = None
    serialization_name = ""
    # Use the potentially updated context_config which might have a refined app_label
    default_preset_name = context_config.get("default_serialization_preset_name")

    if default_preset_name and default_preset_name in SERIALIZATION_PRESETS:
        serialization_name = default_preset_name
        selected_serialization_config = SERIALIZATION_PRESETS[serialization_name]
        print("\n--- Stage 2: Serialization Target ---")
        print(f"✅ Automatically selected serialization based on context: {serialization_name}")
        verbose_print(args, f"DEBUG: Using default serialization preset: {serialization_name}")
    else:
        verbose_print(args, f"DEBUG: No valid default serialization preset found ({default_preset_name}). Prompting user.")
        print("\n--- Stage 2: Select Serialization Target ---")
        serialization_keys = list(SERIALIZATION_PRESETS.keys())
        print("Available Serialization Presets:")
        for i, name in enumerate(serialization_keys):
            desc = SERIALIZATION_PRESETS[name].get('description', '')
            print(f"  [{i}] {name} - {desc}")
        print("  [q] Quit / Go Back")

        while selected_serialization_config is None:
            choice = input("\nEnter your choice: ").strip().lower()
            if choice == 'q': print("Exiting."); return

            if choice.isdigit() and 0 <= int(choice) < len(serialization_keys):
                serialization_name = serialization_keys[int(choice)]
                selected_serialization_config = SERIALIZATION_PRESETS[serialization_name]
                print(f"\nSelected Serialization: {serialization_name}")
                verbose_print(args, f"DEBUG: User selected serialization preset: {serialization_name}")
            else:
                print("❌ Invalid choice.")

    if selected_serialization_config is None:
        print("❌ Serialization target selection was cancelled or failed. Exiting.")
        return

    # --- Stage 3: Get Export Parameters ---
    depth = None
    interval = None

    print("\n--- Stage 3: Set Export Parameters ---")

    # --- Determine Depth ---
    default_depth_value = selected_serialization_config.get("default_depth")
    if default_depth_value is not None and isinstance(default_depth_value, int) and default_depth_value > 0:
        depth = default_depth_value
        print(f"Using default serialization depth from preset: {depth}")
        verbose_print(args, f"DEBUG: Using default depth: {depth}")
    else:
        verbose_print(args, f"DEBUG: No valid default depth found ({default_depth_value}). Prompting user.")
        while True:
            try:
                depth_input = input(f"Enter serialization depth (default 25): ").strip()
                if not depth_input: depth = 25
                else: depth = int(depth_input)
                if depth <= 0: print("❌ Depth must be positive."); continue
                break
            except ValueError: print("❌ Invalid number for depth.")
        verbose_print(args, f"DEBUG: User set depth: {depth}")

    # --- Determine Interval ---
    default_interval_value = selected_serialization_config.get("default_interval")
    if default_interval_value is not None and isinstance(default_interval_value, int) and default_interval_value >= 0:
        interval = default_interval_value
        interval_desc = f"{interval} seconds" if interval > 0 else "Single export"
        print(f"Using default export interval from preset: {interval_desc}")
        verbose_print(args, f"DEBUG: Using default interval: {interval}")
    else:
         verbose_print(args, f"DEBUG: No valid default interval found ({default_interval_value}). Prompting user.")
    if interval is None: # Check if interval still needs to be set
        while True:
            try:
                interval_input = input("Export every N seconds (0 for single export) [default 30]: ").strip()
                if not interval_input: interval = 30
                else: interval = int(interval_input)
                if interval < 0: print("❌ Interval cannot be negative."); continue
                break
            except ValueError: print("❌ Invalid number for interval.")
        verbose_print(args, f"DEBUG: User set interval: {interval}")


    # --- Stage 4: Create Base Directory and Execute Export ---
    parent_timestamp = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
    # Use the app_label determined during context selection for a more descriptive directory
    # Replace spaces/special chars in app_label for filesystem safety
    safe_app_label = "".join(c if c.isalnum() else "_" for c in app_label)
    base_export_dir = f"exports/{safe_app_label}_{parent_timestamp}" # Incorporate app label
    try:
        os.makedirs(base_export_dir, exist_ok=True)
        print(f"\n--- Creating Base Export Directory ---")
        print(f"All exports for this session will be saved under: ./{base_export_dir}")
        verbose_print(args, f"DEBUG: Created base directory: {base_export_dir}")
    except OSError as e:
        print(f"❌ CRITICAL ERROR: Could not create base export directory '{base_export_dir}': {e}")
        return

    # Use the potentially updated app_label from context_config
    print("\n--- Starting Export ---")
    # Use the original chosen_context_name for user display, but the refined context_config['app_label'] for details
    print(f"App Context: {chosen_context_name} (Label: {context_config['app_label']}, PID: {pid})")
    print(f"Serialization Target: {serialization_name}")
    print(f"Max Depth: {depth}")
    print(f"Interval: {'Single Export' if interval <= 0 else f'Every {interval} seconds'}")

    target_criteria = selected_serialization_config.get("target_criteria")

    if interval <= 0:
        # Single Export
        print("Performing a single export...")
        element_to_export = initial_element

        if target_criteria:
             verbose_print(args, f"Searching within initial element for: {target_criteria}...")
             found_sub_element = find_element_by_criteria(initial_element, target_criteria, args)
             if found_sub_element:
                  print("✅ Found target sub-element.")
                  element_to_export = found_sub_element
             else:
                  print("⚠️ Target sub-element not found. Exporting the initial element instead.")
        else:
             verbose_print(args, "Exporting the initial element found by App Context (no target criteria).")

        if element_to_export:
             export_to_json(element_to_export, base_export_dir, depth, args)
             print("✅ Single export complete.")
        else:
             print("❌ Error: No valid element to export.")
             verbose_print(args, "DEBUG: Single export failed because element_to_export is None.")
    else:
        # Periodic Export Loop
        print(f"Starting periodic export every {interval} seconds. Press Ctrl+C to stop.")
        verbose_print(args, f"DEBUG: Preparing to call run_periodic_export.")
        verbose_print(args, f"DEBUG: PID={pid}, Context Present={context_config is not None}, Ser. Present={selected_serialization_config is not None}")

        if pid and context_config and selected_serialization_config:
             try:
                  verbose_print(args, f"DEBUG: Calling run_periodic_export({pid}, ..., depth={depth}, interval={interval}, base_export_dir='{base_export_dir}', args=...).")
                  # Pass the potentially updated context_config
                  run_periodic_export(pid, context_config, selected_serialization_config, depth, interval, base_export_dir, args)
                  verbose_print(args, "DEBUG: run_periodic_export function finished execution.")

             except KeyboardInterrupt:
                  print("\n🛑 Ctrl+C detected. Stopping export loop.")
                  verbose_print(args, "DEBUG: KeyboardInterrupt caught in main.")
             except Exception as e:
                  print(f"\n❌ An unexpected error occurred during the periodic export loop: {e}")
                  verbose_print(args, f"DEBUG: Caught unexpected exception in main's periodic export section.")
                  if args.verbose:
                       traceback.print_exc()
             finally:
                  verbose_print(args, "DEBUG: Reached 'finally' block in main's periodic export section.")
        else:
             print("❌ Error: Missing necessary information (PID, context, or serialization config) to start periodic export.")
             verbose_print(args, f"   Debug Info: PID={pid}, Context Config Present={context_config is not None}, Ser Config Present={selected_serialization_config is not None}")

    verbose_print(args, "DEBUG: End of main function reached.")


if __name__ == "__main__":
    main()
