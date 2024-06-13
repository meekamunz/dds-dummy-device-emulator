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
extracted_data = []

# Configure logging
configure_logging()

def open_file():
    global df_device_names, df_source_ports, df_destination_ports  # Declare global variables
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

# Function to count flows for each GUID and Interface type in Source Ports
def count_flows(guid, interface_type):
    if df_source_ports is not None:
        count = len(df_source_ports[(df_source_ports['GUID'] == guid) & (df_source_ports['Interface'] == interface_type)])
        return count
    else:
        return 0

# Function to process data and create an XML file
def process_and_create_xml():
    global df_device_names, df_source_ports, df_destination_ports
    if df_device_names is not None and df_source_ports is not None and df_destination_ports is not None:
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
                
                # Process Source Spigots for the current device
                current_src_idx = 0
                for s_index, src_row in df_source_ports[df_source_ports['GUID'] == guid].iterrows():
                    src_spigot = ET.SubElement(device, "Spigot")
                    src_spigot.set("idx", str(current_src_idx))
                    src_spigot.set("mode", "src")
                    src_spigot.set("format", "3G")
                    
                    # Count flows for Interface A and B
                    numFlows_A = count_flows(guid, "A")
                    numFlows_B = count_flows(guid, "B")
                    
                    src_spigot.set("numFlows_A", str(numFlows_A))
                    src_spigot.set("numFlows_B", str(numFlows_B))
                    
                    current_src_idx += 1
                
                # Process Destination Spigots for the current device
                current_dst_idx = 0  # Reset index for destination spigots
                for d_index, dst_row in df_destination_ports[df_destination_ports['GUID'] == guid].iterrows():
                    dst_spigot = ET.SubElement(device, "Spigot")
                    dst_spigot.set("idx", str(current_dst_idx))
                    dst_spigot.set("mode", "dst")
                    dst_spigot.set("format", "3G")
                    current_dst_idx += 1
            
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
