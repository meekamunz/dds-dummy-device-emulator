import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import xml.etree.ElementTree as ET
import logging
from logger_config import configure_logging  # Import configure_logging function

# Global variables to store the DataFrames and extracted data
df_device_names = None
df_source_ports = None
df_destination_ports = None
df_source_flows = None
extracted_data = []

# Configure logging
configure_logging()

# Dictionary for Flow Type replacements
flow_type_replacements = {
    "ST 2110-20": "rfc_4175",
    "ST 2110-30": "audio_pcm",
    "ST 2110-40": "meta",
    "ST 2022-6": "smpte2022_6"
}

def open_file():
    global df_device_names, df_source_ports, df_destination_ports, df_source_flows  # Declare global variables
    filepath = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )
    if not filepath:
        return
    
    try:
        logging.info(f"Opening file: {filepath}")
        
        # Read the entire workbook
        xl = pd.ExcelFile(filepath)
        
        # Read Device Names sheet
        df_device_names = xl.parse('Device Names')
        display_data(df_device_names)
        
        # Read Source Ports sheet
        df_source_ports = xl.parse('Source Ports')
        
        # Read Destination Ports sheet
        df_destination_ports = xl.parse('Destination Ports')
        
        # Read Source Flows sheet
        df_source_flows = xl.parse('Source Flows')
        
        logging.info("Data loaded successfully.")
        
        # Automatically process data and create XML file
        process_and_create_xml()
        
    except Exception as e:
        logging.error(f"Failed to read file: {str(e)}")
        messagebox.showerror("Error", f"Failed to read file\n{str(e)}")

def display_data(dataframe):
    # Clear previous content in the treeview
    for i in tree.get_children():
        tree.delete(i)
    # Display the dataframe in the treeview
    tree["column"] = list(dataframe.columns)
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
def create_flow_elements(parent, guid, spigot_index):
    if df_source_flows is not None:
        flow_a_idx = 0
        flow_b_idx = 0
        caps_counts_a = {}  # Dictionary to count caps occurrences for Flow_A
        caps_counts_b = {}  # Dictionary to count caps occurrences for Flow_B

        for _, flow_row in df_source_flows[(df_source_flows['GUID'] == guid) &
                                           (df_source_flows['Spigot Index'] == spigot_index) &
                                           (df_source_flows['Flow Enabled'] == True)].iterrows():
            interface = flow_row['Interface']
            flow_type = flow_row['Flow Type']
            flow_type = flow_type_replacements.get(flow_type, flow_type)  # Replace flow type if applicable

            if interface == "A":
                if flow_type in caps_counts_a:
                    caps_counts_a[flow_type] += 1
                else:
                    caps_counts_a[flow_type] = 1
                
                flow_element = ET.SubElement(parent, "Flow_A")
                flow_element.set("idx", str(flow_a_idx))
                caps_element = ET.SubElement(flow_element, "Caps")
                caps_element.set(flow_type, str(caps_counts_a[flow_type]))
                flow_a_idx += 1
            elif interface == "B":
                if flow_type in caps_counts_b:
                    caps_counts_b[flow_type] += 1
                else:
                    caps_counts_b[flow_type] = 1
                
                flow_element = ET.SubElement(parent, "Flow_B")
                flow_element.set("idx", str(flow_b_idx))
                caps_element = ET.SubElement(flow_element, "Caps")
                caps_element.set(flow_type, str(caps_counts_b[flow_type]))
                flow_b_idx += 1

# Function to process data and create an XML file
def process_and_create_xml():
    global df_device_names, df_source_ports, df_destination_ports, df_source_flows
    if df_device_names is not None and df_source_ports is not None and df_destination_ports is not None and df_source_flows is not None:
        try:
            logging.info("Creating XML file...")
            
            # Create the root element
            root = ET.Element("Devices")
            
            # Loop through each device
            for index, device_row in df_device_names.iterrows():
                guid = device_row['GUID']
                device_name = device_row['Device Name']
                
                device = ET.SubElement(root, "Device")
                device.set("guid", str(guid))
                device.set("typeName", str(device_name))
                
                # Initialize flow counts for destination spigots
                numFlows_A_dst = 0
                numFlows_B_dst = 0
                
                # Process Source Spigots for the current device
                current_spigot_idx = 0  # Initialize spigot index for the device
                caps_counts_a = {}  # Dictionary to count caps occurrences for Flow_A
                caps_counts_b = {}  # Dictionary to count caps occurrences for Flow_B
                
                for s_index, src_row in df_source_ports[df_source_ports['GUID'] == guid].iterrows():
                    src_spigot = ET.SubElement(device, "Spigot")
                    src_spigot.set("idx", str(current_spigot_idx))
                    src_spigot.set("mode", "src")
                    src_spigot.set("format", "3G")
                    
                    # Count flows for Interface A and B
                    spigot_index = current_spigot_idx + 1  # Convert zero-based index to one-based index
                    numFlows_A = count_flows(guid, "A", spigot_index)
                    numFlows_B = count_flows(guid, "B", spigot_index)
                    
                    src_spigot.set("numFlows_A", str(numFlows_A))
                    src_spigot.set("numFlows_B", str(numFlows_B))
                    
                    # Create flow elements for Interface A and B
                    create_flow_elements(src_spigot, guid, spigot_index)
                    
                    # Set flow counts for destination spigots to match spigot idx="0"
                    if current_spigot_idx == 0:
                        numFlows_A_dst = numFlows_A
                        numFlows_B_dst = numFlows_B
                    
                    # Capture caps data for Flow_A and Flow_B
                    caps_counts_a[current_spigot_idx] = caps_counts_a.get(current_spigot_idx, {})
                    caps_counts_b[current_spigot_idx] = caps_counts_b.get(current_spigot_idx, {})
                    
                    for _, flow_row in df_source_flows[(df_source_flows['GUID'] == guid) &
                                                       (df_source_flows['Spigot Index'] == spigot_index) &
                                                       (df_source_flows['Flow Enabled'] == True)].iterrows():
                        interface = flow_row['Interface']
                        flow_type = flow_row['Flow Type']
                        flow_type = flow_type_replacements.get(flow_type, flow_type)  # Replace flow type if applicable
                        
                        if interface == "A":
                            if flow_type in caps_counts_a[current_spigot_idx]:
                                caps_counts_a[current_spigot_idx][flow_type] += 1
                            else:
                                caps_counts_a[current_spigot_idx][flow_type] = 1
                        elif interface == "B":
                            if flow_type in caps_counts_b[current_spigot_idx]:
                                caps_counts_b[current_spigot_idx][flow_type] += 1
                            else:
                                caps_counts_b[current_spigot_idx][flow_type] = 1
                    
                    current_spigot_idx += 1
                
                # Process Destination Spigots for the current device
                current_spigot_idx = 0  # Reset spigot index for destination spigots
                for d_index, dst_row in df_destination_ports[df_destination_ports['GUID'] == guid].iterrows():
                    dst_spigot = ET.SubElement(device, "Spigot")
                    dst_spigot.set("idx", str(current_spigot_idx + 8))  # Destination spigot index starts from 8
                    dst_spigot.set("mode", "dst")
                    dst_spigot.set("format", "3G")
                    
                    # Use the flow counts from the first source spigot
                    dst_spigot.set("numFlows_A", str(numFlows_A_dst))
                    dst_spigot.set("numFlows_B", str(numFlows_B_dst))
                    
                    # Copy caps data from corresponding source spigot
                    src_spigot_index = current_spigot_idx % len(df_source_ports)  # Calculate source spigot index
                    caps_a_data = caps_counts_a[src_spigot_index]
                    caps_b_data = caps_counts_b[src_spigot_index]
                    
                    if "Flow_A" in caps_a_data:
                        for flow_type, count in caps_a_data.items():
                            caps_element = ET.SubElement(dst_spigot, "Caps")
                            caps_element.set(flow_type, str(count))
                    
                    if "Flow_B" in caps_b_data:
                        for flow_type, count in caps_b_data.items():
                            caps_element = ET.SubElement(dst_spigot, "Caps")
                            caps_element.set(flow_type, str(count))
                    
                    current_spigot_idx += 1
            
            # Create a tree structure and write to an XML file
            tree = ET.ElementTree(root)
            tree.write("DummyDevices.xml", encoding="utf-8", xml_declaration=True)
            
            logging.info("XML file created successfully.")
            messagebox.showinfo("Success", "XML file created successfully!")
            
        except Exception as e:
            logging.error(f"Failed to create XML file: {str(e)}")
            messagebox.showerror("Error", f"Failed to create XML file\n{str(e)}")
    else:
        logging.warning("No data loaded.")
        messagebox.showwarning("Warning", "No data loaded. Please open an Excel file first.")


# Create the main application window
root = tk.Tk()
root.title("Create DDS DummyDevices.xml file")

# Create a frame for the buttons
frame = tk.Frame(root)
frame.pack(pady=10)

# Create a button to open the file dialog
open_button = tk.Button(frame, text="Open Excel File", command=open_file)
open_button.pack(side=tk.LEFT, padx=10)

# Create a treeview widget to display the dataframe
tree = ttk.Treeview(root)
tree.pack(expand=True, fill=tk.BOTH)

# Start the main event loop
root.mainloop()
