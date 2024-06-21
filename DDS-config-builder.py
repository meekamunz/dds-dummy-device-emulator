import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import xml.etree.ElementTree as ET
import logging
from logger_config import configure_logging  # Import configure_logging function
from datetime import datetime
import os

# Global variables to store the DataFrames and extracted data
df_device_names = None
df_source_ports = None
df_destination_ports = None
df_source_flows = None
flow_a_index = 0
flow_b_index = 0
processing_window = None

# Configure logging
configure_logging()

# Dictionary for Flow Type replacements
flow_type_replacements = {
    "ST 2110-20": "rfc_4175",
    "ST 2110-30": "audio_pcm",
    "ST 2110-40": "metadata",
    "ST 2022-6": "smpte2022_6"
}

def replace_nan_with_empty_string(df):
    # Replace nan values with empty strings in the DataFrame
    return df.fillna('')

def open_file():
    global df_device_names, df_source_ports, df_destination_ports, df_source_flows, create_xml_button  # Declare global variables
    filepath = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )
    if not filepath:
        return
    
    try:
        show_processing_window("Opening and processing file...")  # Show processing window
        root.update()  # Update the root window to display the processing window

        logging.info(f"Opening file: {filepath}")
        
        # Read the entire workbook
        xl = pd.ExcelFile(filepath)
        
        # Read and filter Device Names sheet
        df_device_names = xl.parse('Device Names')
        df_device_names = df_device_names[df_device_names['Device Type'] == 'GVOP']
        df_device_names = replace_nan_with_empty_string(df_device_names)  # Replace nan values
        display_data(df_device_names)
        
        # Read and filter Source Ports sheet
        df_source_ports = xl.parse('Source Ports')
        df_source_ports = df_source_ports[df_source_ports['Device Type'] == 'GVOP']
        # df_source_ports = replace_nan_with_empty_string(df_source_ports)  # Replace nan values
        
        # Read and filter Destination Ports sheet
        df_destination_ports = xl.parse('Destination Ports')
        df_destination_ports = df_destination_ports[df_destination_ports['Device Type'] == 'GVOP']
        # df_destination_ports = replace_nan_with_empty_string(df_destination_ports)  # Replace nan values
        
        # Read and filter Source Flows sheet
        df_source_flows = xl.parse('Source Flows')
        df_source_flows = df_source_flows[df_source_flows['Device Type'] == 'GVOP']
        # df_source_flows = replace_nan_with_empty_string(df_source_flows)  # Replace nan values

        logging.info("Data loaded successfully.")
        
        # Enable create_xml_button after data is loaded
        create_xml_button.config(state=tk.NORMAL)
        
    except Exception as e:
        logging.error(f"Failed to read file: {str(e)}")
        messagebox.showerror("Error", f"Failed to read file\n{str(e)}")
        # Disable create_xml_button if data loading fails
        create_xml_button.config(state=tk.DISABLED)
    finally:
        hide_processing_window()  # Hide processing window

def display_data(dataframe):
    # Clear previous content in the treeview
    for i in tree.get_children():
        tree.delete(i)
    # Display the dataframe in the treeview
    tree["columns"] = list(dataframe.columns)
    tree["show"] = "headings"
    for column in tree["columns"]:
        tree.heading(column, text=column)
    df_rows = dataframe.to_numpy().tolist()
    for row in df_rows:
        tree.insert("", "end", values=row)
    tree.pack(expand=True, fill=tk.BOTH)

# Function to count flows for each GUID, Interface type, and Spigot Index in Source Flows
def count_flows(guid, interface_type, spigot_index):
    if df_source_flows is not None:
        count = len(df_source_flows[(df_source_flows['GUID'] == guid) &
                                    (df_source_flows['Interface'] == interface_type) &
                                    (df_source_flows['Spigot Index'] == spigot_index) &
                                    (df_source_flows['Flow Enabled'] == True)])
        return count
    else:
        return 0

# Function to create Flow_A and Flow_B elements for each flow
def create_flow_elements(parent, guid, spigot_index, is_source_spigot):
    global flow_a_index, flow_b_index
    
    # Initialize counters for Flow_A and Flow_B
    flow_a_index = 0
    flow_b_index = 0
    
    caps_counts_a = {}  # Dictionary to count caps occurrences for Flow_A
    caps_counts_b = {}  # Dictionary to count caps occurrences for Flow_B

    if df_source_flows is not None:
        for _, flow_row in df_source_flows[(df_source_flows['GUID'] == guid) &
                                           (df_source_flows['Spigot Index'] == spigot_index) &
                                           (df_source_flows['Flow Enabled'] == True)].iterrows():
            interface = flow_row['Interface']
            flow_type = flow_row['Flow Type']
            flow_type = flow_type_replacements.get(flow_type, flow_type)  # Replace flow type if applicable

            # Params values from the Source Flows worksheet
            mcast_address = flow_row['Multicast Address']
            src_address = flow_row['Source Address']
            dst_port = int(flow_row['Dst RTP Port'])  # Convert to integer

            # Attempt to use "Src RTP Port", fall back to "Dst RTP Port" if not available
            try:
                src_port = int(flow_row['Src RTP Port']) if pd.notnull(flow_row['Src RTP Port']) else dst_port
            except KeyError:
                src_port = dst_port  # Fallback in case the column doesn't exist

            params_type = flow_type

            if params_type == "meta":
                params_type = "metadata"

            if interface == "A":
                logging.debug(f"Creating Flow_A: GUID={guid}, Spigot Index={spigot_index}, Flow Type={flow_type}")
                flow_element = ET.SubElement(parent, "Flow_A")
                flow_element.set("idx", str(flow_a_index))
                caps_element = ET.SubElement(flow_element, "Caps")
                caps_element.set(flow_type, "1")
                
                if is_source_spigot:
                    params_element = ET.SubElement(flow_element, "Params")
                    params_element.set("mcastAddress", str(mcast_address))
                    params_element.set("srcAddress", str(src_address))
                    params_element.set("dstPort", str(dst_port))
                    params_element.set("srcPort", str(src_port))
                    params_element.set("type", str(params_type))

                # Update caps count for first spigot
                if flow_a_index == 0 and is_source_spigot:
                    if flow_type in caps_counts_a:
                        caps_counts_a[flow_type] += 1
                    else:
                        caps_counts_a[flow_type] = 1

                flow_a_index += 1
                
            elif interface == "B":
                logging.debug(f"Creating Flow_B: GUID={guid}, Spigot Index={spigot_index}, Flow Type={flow_type}")
                flow_element = ET.SubElement(parent, "Flow_B")
                flow_element.set("idx", str(flow_b_index))
                caps_element = ET.SubElement(flow_element, "Caps")
                caps_element.set(flow_type, "1")
                
                if is_source_spigot:
                    params_element = ET.SubElement(flow_element, "Params")
                    params_element.set("mcastAddress", str(mcast_address))
                    params_element.set("srcAddress", str(src_address))
                    params_element.set("dstPort", str(dst_port))
                    params_element.set("srcPort", str(src_port))
                    params_element.set("type", str(params_type))

                # Update caps count for first spigot
                if flow_b_index == 0 and is_source_spigot:
                    if flow_type in caps_counts_b:
                        caps_counts_b[flow_type] += 1
                    else:
                        caps_counts_b[flow_type] = 1

                flow_b_index += 1
        
        # Store caps counts for the first spigot if it's the first spigot
        if is_source_spigot and spigot_index == 1:
            df_first_spigot_flows_a = [{"idx": idx, "type": key, "count": value} for idx, (key, value) in enumerate(caps_counts_a.items())]
            df_first_spigot_flows_b = [{"idx": idx, "type": key, "count": value} for idx, (key, value) in enumerate(caps_counts_b.items())]

    # Calculate numFlows_A and numFlows_B for destination spigots
    if not is_source_spigot:
        num_flows_a = 3  # Assuming 3 Flow_A elements per destination spigot
        num_flows_b = 3  # Assuming 3 Flow_B elements per destination spigot
        
        parent.set("numFlows_A", str(num_flows_a))
        parent.set("numFlows_B", str(num_flows_b))

def process_and_create_xml(filepath):
    global df_device_names, df_source_ports, df_destination_ports, df_source_flows
    
    if df_device_names is not None and df_source_ports is not None and df_destination_ports is not None and df_source_flows is not None:
        try:
            logging.info(f"Creating XML file: {filepath}")
            
            devices = []  # List to store individual device XML strings
            
            for index, device_row in df_device_names.iterrows():
                guid = device_row['GUID']
                device_name = device_row['Device Name']
                ip_address_a = device_row['IP Address']  # Assuming 'IP Address' column exists in df_device_names
                
                # Find any source address for Interface B matching the GUID and conditions
                source_address_b = ""  # Default value if no valid address found
                
                # Check if there are rows matching the conditions
                rows_matching_conditions = df_source_flows[(df_source_flows['GUID'] == guid) & 
                                                           (df_source_flows['Interface'] == 'B') & 
                                                           (df_source_flows['Flow Enabled'] == True)]
                
                if not rows_matching_conditions.empty:
                    source_address_b = rows_matching_conditions['Source Address'].iloc[0]
                else:
                    logging.warning(f"No valid source address found for GUID {guid} and Interface B.")
                
                # Proceed with creating XML using source_address_b
                device = ET.Element("Device")
                device.set("guid", str(guid))
                device.set("typeName", str(device_name))
                device.set("softVer", "DummyDDS")
                device.set("firmVer", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                device.set("ipAddressA", str(ip_address_a))
                device.set("ipAddressB", str(source_address_b))
                device.set("linkSpeedA", "25000")  # Set linkSpeedA to "25000"
                device.set("linkSpeedB", "25000")  # Set linkSpeedB to "25000"
                
                current_spigot_idx = 0  # Initialize spigot index for the device
                
                # Count number of source spigots
                num_sources = len(df_source_ports[df_source_ports['GUID'] == guid])
                device.set("numSources", str(num_sources))
                
                # Count number of destination spigots
                num_dests = len(df_destination_ports[df_destination_ports['GUID'] == guid])
                device.set("numDests", str(num_dests))
                
                # Process Source Spigots for the current device
                for s_index, src_row in df_source_ports[df_source_ports['GUID'] == guid].iterrows():
                    src_spigot = ET.SubElement(device, "Spigot")
                    
                    # Read the spigot index from the Source Flows worksheet and adjust
                    spigot_index = int(src_row['Spigot Index']) - 1  # Convert 1-based to 0-based index
                    
                    src_spigot.set("idx", str(spigot_index))
                    src_spigot.set("mode", "src")
                    src_spigot.set("format", "3G")
                    
                    numFlows_A = count_flows(guid, "A", spigot_index + 1)
                    numFlows_B = count_flows(guid, "B", spigot_index + 1)
                    
                    src_spigot.set("numFlows_A", str(numFlows_A))
                    src_spigot.set("numFlows_B", str(numFlows_B))
                    
                    # Create Flow_A and Flow_B elements for the current source spigot
                    create_flow_elements(src_spigot, guid, spigot_index + 1, is_source_spigot=True)
                
                # Process Destination Spigots for the current device
                for d_index, dst_row in df_destination_ports[df_destination_ports['GUID'] == guid].iterrows():
                    dst_spigot = ET.SubElement(device, "Spigot")
                    
                    # Adjusted spigot index from Destination Ports worksheet
                    dst_spigot_idx = int(dst_row['Spigot Index']) - 1  # Convert 1-based to 0-based index
                    dst_spigot.set("idx", str(dst_spigot_idx))
                    dst_spigot.set("mode", "dst")
                    dst_spigot.set("format", "3G")
                    
                    # Retrieve all Flow_A and Flow_B instances from the first source spigot dynamically
                    first_spigot_flows_a = df_source_flows[(df_source_flows['GUID'] == guid) & 
                                                           (df_source_flows['Spigot Index'] == 1) & 
                                                           (df_source_flows['Interface'] == 'A') & 
                                                           (df_source_flows['Flow Enabled'] == True)]
                    
                    first_spigot_flows_b = df_source_flows[(df_source_flows['GUID'] == guid) & 
                                                           (df_source_flows['Spigot Index'] == 1) & 
                                                           (df_source_flows['Interface'] == 'B') & 
                                                           (df_source_flows['Flow Enabled'] == True)]
                    
                    numFlows_A = len(first_spigot_flows_a)
                    numFlows_B = len(first_spigot_flows_b)
                    
                    dst_spigot.set("numFlows_A", str(numFlows_A))
                    dst_spigot.set("numFlows_B", str(numFlows_B))
                    
                    if numFlows_A == 0 and numFlows_B == 0:
                        # No source spigots found, add default Flow_A and Flow_B elements
                        add_default_flows(dst_spigot)
                    
                    # Copy all Flow_A and Flow_B instances from the first source spigot to the destination spigot
                    copy_caps_to_destination_spigots(dst_spigot, first_spigot_flows_a, first_spigot_flows_b)
                    
                    # Create Flow_A and Flow_B elements for the current destination spigot
                    create_flow_elements(dst_spigot, guid, dst_spigot_idx + 1, is_source_spigot=False)
                
                devices.append(ET.tostring(device, encoding='unicode'))
            
            # Write XML to file
            with open(filepath, "w") as xml_file:
                xml_file.write('\n'.join(devices))
                
            logging.info(f"XML file created successfully: {filepath}")
        
        except Exception as e:
            logging.error(f"Error creating XML file: {e}")

def add_default_flows(dst_spigot):
    # Add default Flow_A elements
    flow_a_types = ["rfc_4175", "audio_pcm", "metadata"]
    for idx, flow_type in enumerate(flow_a_types):
        flow_a = ET.SubElement(dst_spigot, "Flow_A")
        flow_a.set("idx", str(idx))
        caps_element_a = ET.SubElement(flow_a, "Caps")
        caps_element_a.set(flow_type, "1")
    
    # Add default Flow_B elements
    flow_b_types = ["rfc_4175", "audio_pcm", "metadata"]
    for idx, flow_type in enumerate(flow_b_types):
        flow_b = ET.SubElement(dst_spigot, "Flow_B")
        flow_b.set("idx", str(idx))
        caps_element_b = ET.SubElement(flow_b, "Caps")
        caps_element_b.set(flow_type, "1")

# Function to copy Caps to destination spigots for all flow types
def copy_caps_to_destination_spigots(dst_spigot, flows_a, flows_b):
    global flow_type_replacements
    
    # Reset counters for Flow_A and Flow_B
    flow_a_index = 0
    flow_b_index = 0

    # Iterate over all Flow_A from the first source spigot and copy to destination spigot
    for _, row in flows_a.iterrows():
        flow_type = row['Flow Type']
        flow_type = flow_type_replacements.get(flow_type, flow_type)  # Replace flow type if applicable
        
        flow_a = ET.SubElement(dst_spigot, "Flow_A")
        flow_a.set("idx", str(flow_a_index))  # Reset idx per spigot
        caps_element_a = ET.SubElement(flow_a, "Caps")
        caps_element_a.set(flow_type, "1")  # Assuming count of 1 for each cap type

        flow_a_index += 1
    
    # Iterate over all Flow_B from the first source spigot and copy to destination spigot
    for _, row in flows_b.iterrows():
        flow_type = row['Flow Type']
        flow_type = flow_type_replacements.get(flow_type, flow_type)  # Replace flow type if applicable
        
        flow_b = ET.SubElement(dst_spigot, "Flow_B")
        flow_b.set("idx", str(flow_b_index))  # Reset idx per spigot
        caps_element_b = ET.SubElement(flow_b, "Caps")
        caps_element_b.set(flow_type, "1")  # Assuming count of 1 for each cap type

        flow_b_index += 1

# Function to handle XML creation process after data is loaded
def create_xml_process():
    if df_device_names is not None and df_source_ports is not None and df_destination_ports is not None and df_source_flows is not None:
        try:
            # Prompt user to select directory and enter filename
            filepath = filedialog.asksaveasfilename(
                initialfile="DummyDevices.xml",
                defaultextension=".xml",
                filetypes=[("XML files", "*.xml"), ("All files", "*.*")],
                title="Save XML file as"
            )
            if not filepath:
                return  # User canceled

            show_processing_window("Creating XML file...")  # Show processing window
            root.update()  # Update the root window to display the processing window

            # Process and create XML file
            process_and_create_xml(filepath)
        
        except Exception as e:
            logging.error(f"Failed to create XML file: {str(e)}")
            messagebox.showerror("Error", f"Failed to create XML file:\n{str(e)}")
        finally:
            hide_processing_window()  # Hide processing window
    
    else:
        logging.warning("No data loaded.")

# Create the main application window
root = tk.Tk()
root.title("Create DDS Dummy Devices file")

## Create a frame for the treeview
data_frame = tk.Frame(root)
data_frame.pack(expand=True, fill=tk.BOTH)

# Create a treeview widget to display the dataframe
tree = ttk.Treeview(data_frame)
tree.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)  # Adjust packing

# Create a vertical scrollbar for the tree widget
tree_scrollbar = ttk.Scrollbar(data_frame, orient=tk.VERTICAL, command=tree.yview)
tree_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)  # Adjust packing
tree.config(yscrollcommand=tree_scrollbar.set)

# Create a frame for the log text and buttons
log_frame = tk.Frame(root)
log_frame.pack(expand=True, fill=tk.BOTH)

# Create a text widget to display the tail of the logfile
log_text = tk.Text(log_frame, height=10, wrap=tk.WORD)
log_text.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)

# Create a vertical scrollbar for the log text widget
log_scrollbar = tk.Scrollbar(log_frame, orient=tk.VERTICAL, command=log_text.yview)
log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
log_text.config(yscrollcommand=log_scrollbar.set)

def show_processing_window(message="Processing..."):
    global processing_window
    processing_window = tk.Toplevel(root)
    processing_window.title("Please Wait")
    processing_window.geometry("300x100")
    processing_window.resizable(False, False)
    tk.Label(processing_window, text=message).pack(pady=20)
    processing_window.grab_set()  # Make the processing window modal

def hide_processing_window():
    global processing_window
    if processing_window:
        processing_window.destroy()


# Function to update the log text widget with the tail of the logfile
def update_log_text():
    global log_text
    try:
        # Get current vertical scrollbar position
        scrollbar_pos = log_text.yview()[1]
        
        # Open the log file
        with open('DummyDeviceBuilder.log', 'r') as log_file:
            # Read the last 20 lines (you can adjust the number of lines as needed)
            log_lines = log_file.readlines()[-20:]
            # Clear previous content in the text widget
            log_text.delete('1.0', tk.END)
            # Display the log lines, with color inversion
            for line in log_lines:
                if "ERROR" in line:
                    log_text.insert(tk.END, line, 'error')
                elif "WARNING" in line:
                    log_text.insert(tk.END, line, 'warning')
                else:
                    log_text.insert(tk.END, line)
            
            # Configure tag colors with inversion
            log_text.tag_config('error', foreground='white', background='red')
            log_text.tag_config('warning', foreground='white', background='orange')
            log_text.tag_config('info', foreground='black', background='light grey')  # Adjust as needed
        
        # Scroll to the bottom only if the scrollbar was already at the bottom
        if scrollbar_pos == 1.0:
            log_text.see(tk.END)  # Scroll to the bottom if scrollbar was at the bottom
        
    except FileNotFoundError:
        # Handle case where the log file doesn't exist yet
        pass
    except Exception as e:
        # Display error message if unable to read log file
        log_text.insert(tk.END, f"Error reading log file: {str(e)}", 'error')
    
    # Schedule the update_log_text function to be called again after 1000 milliseconds (1 second)
    log_text.after(1000, update_log_text)

# Call the function initially to populate the log text widget
update_log_text()

# Initial configuration for log_text
log_text.configure(bg='black', fg='white')  # Set background to black and text to white

# Create a frame for the buttons
button_frame = tk.Frame(root)
button_frame.pack(pady=10)

# Create a button to open the file dialog
open_button = tk.Button(button_frame, text="Open IP Configurator Export file", command=open_file)
open_button.pack(side=tk.LEFT, padx=10)

# Create a button to create the XML file
create_xml_button = tk.Button(button_frame, text="Create Dummy DDS file", command=create_xml_process)
create_xml_button.pack(side=tk.LEFT, padx=10)

# Start the main event loop
root.mainloop()