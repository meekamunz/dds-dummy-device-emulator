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
df_first_spigot_flows_a = []  # Initialize as empty list
df_first_spigot_flows_b = []  # Initialize as empty list
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
        
        # Read and filter Device Names sheet
        df_device_names = xl.parse('Device Names')
        df_device_names = df_device_names[df_device_names['Device Type'] == 'GVOP']
        display_data(df_device_names)
        
        # Read and filter Source Ports sheet
        df_source_ports = xl.parse('Source Ports')
        df_source_ports = df_source_ports[df_source_ports['Device Type'] == 'GVOP']
        
        # Read and filter Destination Ports sheet
        df_destination_ports = xl.parse('Destination Ports')
        df_destination_ports = df_destination_ports[df_destination_ports['Device Type'] == 'GVOP']
        
        # Read and filter Source Flows sheet
        df_source_flows = xl.parse('Source Flows')
        df_source_flows = df_source_flows[df_source_flows['Device Type'] == 'GVOP']
        
        logging.info("Data loaded successfully.")
        
    except Exception as e:
        logging.error(f"Failed to read file: {str(e)}")
        messagebox.showerror("Error", f"Failed to read file\n{str(e)}")

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
    global df_first_spigot_flows_a, df_first_spigot_flows_b  # Declare global variables
    
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
            src_port = int(flow_row['Src RTP Port'])  # Convert to integer
            params_type = flow_type

            if params_type == "meta":
                params_type = "metadata"

            if interface == "A":
                if flow_type in caps_counts_a:
                    caps_counts_a[flow_type] += 1
                else:
                    caps_counts_a[flow_type] = 1
                
                if is_source_spigot:
                    flow_element = ET.SubElement(parent, "Flow_A")
                    flow_element.set("idx", str(len(caps_counts_a) - 1))
                    caps_element = ET.SubElement(flow_element, "Caps")
                    caps_element.set(flow_type, str(caps_counts_a[flow_type]))
                    
                    # Create Params element only for source spigots
                    params_element = ET.SubElement(flow_element, "Params")
                    params_element.set("mcastAddress", str(mcast_address))
                    params_element.set("srcAddress", str(src_address))
                    params_element.set("dstPort", str(dst_port))
                    params_element.set("srcPort", str(src_port))
                    params_element.set("type", str(params_type))

                if spigot_index == 1:
                    df_first_spigot_flows_a.append({"idx": len(caps_counts_a) - 1, "type": flow_type, "count": caps_counts_a[flow_type]})
                    
            elif interface == "B":
                if flow_type in caps_counts_b:
                    caps_counts_b[flow_type] += 1
                else:
                    caps_counts_b[flow_type] = 1
                
                if is_source_spigot:
                    flow_element = ET.SubElement(parent, "Flow_B")
                    flow_element.set("idx", str(len(caps_counts_b) - 1))
                    caps_element = ET.SubElement(flow_element, "Caps")
                    caps_element.set(flow_type, str(caps_counts_b[flow_type]))
                    
                    # Create Params element only for source spigots
                    params_element = ET.SubElement(flow_element, "Params")
                    params_element.set("mcastAddress", str(mcast_address))
                    params_element.set("srcAddress", str(src_address))
                    params_element.set("dstPort", str(dst_port))
                    params_element.set("srcPort", str(src_port))
                    params_element.set("type", str(params_type))

                if spigot_index == 1:
                    df_first_spigot_flows_b.append({"idx": len(caps_counts_b) - 1, "type": flow_type, "count": caps_counts_b[flow_type]})
        
        # Copy caps from first spigot to all destination spigots
        if not is_source_spigot and spigot_index == 1:
            df_first_spigot_flows_a = [{"idx": flow["idx"], "type": flow["type"], "count": flow["count"]} for flow in df_first_spigot_flows_a]
            df_first_spigot_flows_b = [{"idx": flow["idx"], "type": flow["type"], "count": flow["count"]} for flow in df_first_spigot_flows_b]

# Function to process data and create an XML file
def process_and_create_xml(filepath):
    global df_device_names, df_source_ports, df_destination_ports, df_source_flows
    global df_first_spigot_flows_a, df_first_spigot_flows_b
    
    if df_device_names is not None and df_source_ports is not None and df_destination_ports is not None and df_source_flows is not None:
        try:
            logging.info(f"Creating XML file: {filepath}")
            
            devices = []  # List to store individual device XML strings
            
            for index, device_row in df_device_names.iterrows():
                guid = device_row['GUID']
                device_name = device_row['Device Name']
                ip_address_a = device_row['IP Address']  # Assuming 'IP Address' column exists in df_device_names
                
                # Find any source address for Interface B matching the GUID and conditions
                source_address_b = df_source_flows.loc[(df_source_flows['GUID'] == guid) & 
                                                       (df_source_flows['Interface'] == 'B') & 
                                                       (df_source_flows['Flow Enabled'] == True), 
                                                       'Source Address'].iloc[0]

                device = ET.Element("Device")
                device.set("guid", str(guid))
                device.set("typeName", str(device_name))
                device.set("softVer", "DummyDDS")
                device.set("firmVer", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                device.set("ipAddressA", str(ip_address_a))
                device.set("ipAddressB", str(source_address_b))
                device.set("linkSpeedA", "25000")  # Set linkSpeedA to "25000"
                device.set("linkSpeedB", "25000")  # Set linkSpeedB to "25000"
                
                device_first_spigot_caps_a = {}
                device_first_spigot_caps_b = {}

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
                    src_spigot.set("idx", str(current_spigot_idx))
                    src_spigot.set("mode", "src")
                    src_spigot.set("format", "3G")
                    
                    spigot_index = current_spigot_idx + 1  # Convert zero-based index to one-based index
                    numFlows_A = count_flows(guid, "A", spigot_index)
                    numFlows_B = count_flows(guid, "B", spigot_index)
                    
                    src_spigot.set("numFlows_A", str(numFlows_A))
                    src_spigot.set("numFlows_B", str(numFlows_B))
                    
                    create_flow_elements(src_spigot, guid, spigot_index, is_source_spigot=True)
                    
                    if current_spigot_idx == 0:
                        device_first_spigot_caps_a = df_first_spigot_flows_a.copy()
                        device_first_spigot_caps_b = df_first_spigot_flows_b.copy()
                    
                    current_spigot_idx += 1
                
                # Process Destination Spigots for the current device
                for d_index, dst_row in df_destination_ports[df_destination_ports['GUID'] == guid].iterrows():
                    dst_spigot = ET.SubElement(device, "Spigot")
                    
                    dst_spigot_idx = int(dst_row['Spigot Index']) - 1  # Zero-based index
                    dst_spigot.set("idx", str(dst_spigot_idx))
                    dst_spigot.set("mode", "dst")
                    dst_spigot.set("format", "3G")
                    
                    numFlows_A = count_flows(guid, "A", dst_spigot_idx + 1)
                    numFlows_B = count_flows(guid, "B", dst_spigot_idx + 1)
                    
                    dst_spigot.set("numFlows_A", str(numFlows_A))
                    dst_spigot.set("numFlows_B", str(numFlows_B))

                    copy_caps_to_destination_spigot(dst_spigot, device_first_spigot_caps_a, device_first_spigot_caps_b)
                
                # Convert the device element to a string and add to the devices list
                devices.append(ET.tostring(device, encoding="unicode"))
                
                # Reset first spigot flow counters
                df_first_spigot_flows_a = []
                df_first_spigot_flows_b = []
            
            # Concatenate all device XML strings into a single XML string
            xml_content = "\n".join(devices)
            
            # Write the XML content to the file
            with open(filepath, "w", encoding="utf-8") as f:
                f.write(xml_content)
            
            logging.info(f"XML file created successfully: {filepath}")
            messagebox.showinfo("Success", f"XML file created successfully:\n{filepath}")
        
        except Exception as e:
            logging.error(f"Failed to create XML file: {str(e)}")
            messagebox.showerror("Error", f"Failed to create XML file:\n{str(e)}")
    
    else:
        logging.warning("No data loaded.")

def copy_caps_to_destination_spigot(dst_spigot, first_spigot_caps_a, first_spigot_caps_b):
    # Function to copy Caps entries from the first source spigot to destination spigots
    for flow_a in first_spigot_caps_a:
        flow_element = ET.SubElement(dst_spigot, "Flow_A")
        flow_element.set("idx", str(flow_a["idx"]))
        caps_element = ET.SubElement(flow_element, "Caps")
        caps_element.set(flow_a["type"], str(flow_a["count"]))
        
    for flow_b in first_spigot_caps_b:
        flow_element = ET.SubElement(dst_spigot, "Flow_B")
        flow_element.set("idx", str(flow_b["idx"]))
        caps_element = ET.SubElement(flow_element, "Caps")
        caps_element.set(flow_b["type"], str(flow_b["count"]))

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
            
            # Process and create XML file
            process_and_create_xml(filepath)
        
        except Exception as e:
            logging.error(f"Failed to create XML file: {str(e)}")
            messagebox.showerror("Error", f"Failed to create XML file:\n{str(e)}")
    
    else:
        logging.warning("No data loaded.")

# Create the main application window
root = tk.Tk()
root.title("Create DDS Dummy Devices file")

# Create a frame for the buttons
frame = tk.Frame(root)
frame.pack(pady=10)

# Create a button to open the file dialog
open_button = tk.Button(frame, text="Open IP Configurator Export file", command=open_file)
open_button.pack(side=tk.LEFT, padx=10)

# Create a button to create the XML file
create_xml_button = tk.Button(frame, text="Create Dummy DDS file", command=create_xml_process)
create_xml_button.pack(side=tk.LEFT, padx=10)

# Create a treeview widget to display the dataframe
tree = ttk.Treeview(root)
tree.pack(expand=True, fill=tk.BOTH)

# Start the main event loop
root.mainloop()
