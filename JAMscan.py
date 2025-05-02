import subprocess
import sys
import threading
import json
import os
import socket
import argparse
import time
from scapy.all import *
import tkinter as tk
from tkinter import ttk
from functools import partial

# File to store the discovered network data
DATA_FILE = 'network_data.json'
MAX_IP_THRESHOLD = 10
VERBOSE = False  # Global variable for verbose output
REFRESH_INTERVAL = 3600  # Time in seconds to refresh network data

# Function for verbose logging
def log_verbose(message):
    if VERBOSE:
        print(message)

# Function to check for administrative privileges
def check_privileges():
    if os.name == 'nt':
        # Windows
        try:
            import ctypes
            is_admin = ctypes.windll.shell32.IsUserAnAdmin()
        except:
            is_admin = False
        if not is_admin:
            print("This script must be run as administrator.")
            sys.exit(1)
    else:
        # Unix/Linux
        if os.geteuid() != 0:
            print("This script must be run as root.")
            sys.exit(1)

# Function to detect the local network's subnet
def detect_local_subnet():
    try:
        # Get the local IP address
        hostname = socket.gethostname()
        local_ip = socket.gethostbyname(hostname)
        # Assume a /24 subnet mask
        network_prefix = '.'.join(local_ip.split('.')[:3]) + '.0/24'
        log_verbose(f"Detected local subnet: {network_prefix}")
        return network_prefix
    except Exception as e:
        log_verbose(f"Error detecting local subnet: {e}")
        return None

# Preliminary network scan to quickly map the network
def preliminary_scan(network_data):
    log_verbose("Performing preliminary network scan...")
    subnet = detect_local_subnet()
    if not subnet:
        print("Could not detect local subnet. Exiting preliminary scan.")
        return

    try:
        # Use arping to perform a fast ARP scan
        ans, unans = arping(subnet, timeout=0.5, verbose=0)
        for snd, rcv in ans:
            ip = rcv.psrc
            mac = rcv.hwsrc
            try:
                hostname = socket.gethostbyaddr(ip)[0]
            except (socket.herror, socket.gaierror):
                hostname = None

            # Update device info
            existing_data = network_data['devices'].get(ip, {})
            existing_data.update({
                'mac': mac,
                'hostname': hostname,
                'last_checked': time.time()
            })
            network_data['devices'][ip] = existing_data

            # Add the device to its respective network
            network_prefix = '.'.join(ip.split('.')[:3])
            network = f"{network_prefix}.0/24"
            if network not in network_data['networks']:
                network_data['networks'][network] = []
            if ip not in network_data['networks'][network]:
                network_data['networks'][network].append(ip)

            log_verbose(f"Discovered device: IP={ip}, MAC={mac}, Hostname={hostname or 'Unknown'}")
    except Exception as e:
        log_verbose(f"Error during preliminary scan: {e}")

    # Save network data after preliminary scan
    save_network_data(network_data)
    print("Preliminary network scan completed and data saved.")

# Function to perform traceroute and get the list of hops
def perform_traceroute(destination_ip, max_hops=5):
    log_verbose(f"Performing traceroute to {destination_ip} with max {max_hops} hops...")
    hops = []
    try:
        for ttl in range(1, max_hops + 1):
            log_verbose(f"Sending packet with TTL={ttl}")
            pkt = IP(dst=destination_ip, ttl=ttl) / ICMP()
            reply = sr1(pkt, verbose=0, timeout=2)
            if reply is None:
                log_verbose(f"No reply for TTL={ttl}")
                break
            else:
                hops.append(reply.src)
                log_verbose(f"Received reply from {reply.src}")
                if reply.src == destination_ip:
                    break
    except Exception as e:
        log_verbose(f"Error during traceroute to {destination_ip}: {e}")
    return hops

# Function to get ARP table entries from a device
def get_arp_table(ip_address):
    log_verbose(f"Retrieving ARP table from {ip_address}...")
    arp_entries = []
    try:
        # Extract the network prefix from the IP address
        network_prefix = '.'.join(ip_address.split('.')[:3]) + '.0/24'
        # Build ARP request packet
        arp_req = Ether(dst='ff:ff:ff:ff:ff:ff') / ARP(pdst=network_prefix)
        # Send the packet and receive responses
        ans, unans = srp(arp_req, timeout=0.5, verbose=0)
        for snd, rcv in ans:
            try:
                hostname = socket.gethostbyaddr(rcv.psrc)[0]  # Get hostname
            except (socket.herror, socket.gaierror):
                hostname = None  # Fallback if hostname is unavailable

            # Determine if the device is a ghost device
            is_ghost = False
            if hostname is None and rcv.hwsrc == "00:00:0c:9f:f1:7e":
                is_ghost = True

            log_verbose(f"ARP entry: IP={rcv.psrc}, MAC={rcv.hwsrc}, Hostname={hostname or 'Unknown'}, Ghost Device={is_ghost}")
            arp_entries.append({
                'ip': rcv.psrc,
                'mac': rcv.hwsrc,
                'hostname': hostname,
                'is_ghost': is_ghost
            })
    except Exception as e:
        log_verbose(f"Error retrieving ARP table from {ip_address}: {e}")
    return arp_entries

# Load or initialize the network data file
def load_network_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, 'r') as f:
            data = json.load(f)
        log_verbose("Network data loaded from file.")
        return data
    else:
        log_verbose("No existing network data found. Initializing new data.")
        return {'devices': {}, 'networks': {}, 'connections': []}

# Save network data to a file
def save_network_data(data):
    with open(DATA_FILE, 'w') as f:
        json.dump(data, f, indent=4)
    log_verbose("Network data saved to file.")

# Function to build or update the network map
def update_network_map(hops, network_data):
    new_data_found = False
    current_time = time.time()
    threads = []

    def process_hop(hop):
        nonlocal new_data_found
        log_verbose(f"Processing hop {hop}")
        device_data = network_data['devices'].get(hop, {})
        last_checked = device_data.get('last_checked', 0)
        if current_time - last_checked > REFRESH_INTERVAL:
            arp_entries = get_arp_table(hop)
            # Filter out ghost devices
            arp_entries = [entry for entry in arp_entries if not entry['is_ghost']]
            log_verbose(f"Found {len(arp_entries)} ARP entries for {hop}")

            # Update hop device info
            existing_data = device_data
            existing_data.update({
                'arp_entries': arp_entries,
                'last_checked': current_time,
                'hostname': get_hostname(hop)
            })
            network_data['devices'][hop] = existing_data
            new_data_found = True

            # Add the device to its respective network
            network_prefix = '.'.join(hop.split('.')[:3])
            network = f"{network_prefix}.0/24"
            if network not in network_data['networks']:
                network_data['networks'][network] = []
            if hop not in network_data['networks'][network]:
                network_data['networks'][network].append(hop)

            # Add connections between the device and its ARP entries
            for entry in arp_entries:
                # Update device info unconditionally
                existing_entry = network_data['devices'].get(entry['ip'], {})
                existing_entry.update({
                    'mac': entry['mac'],
                    'hostname': entry['hostname'],
                    'last_checked': current_time
                })
                network_data['devices'][entry['ip']] = existing_entry

                if {'from': hop, 'to': entry['ip']} not in network_data['connections']:
                    network_data['connections'].append({'from': hop, 'to': entry['ip']})
                    log_verbose(f"Added connection from {hop} to {entry['ip']}")

            # Save network data after processing this hop
            save_network_data(network_data)
        else:
            log_verbose(f"Skipping hop {hop}, last checked {current_time - last_checked:.0f} seconds ago.")

    # Use threading to process hops in parallel
    for hop in hops:
        thread = threading.Thread(target=process_hop, args=(hop,))
        threads.append(thread)
        thread.start()

    for thread in threads:
        thread.join()

    return new_data_found

# Function to get hostname
def get_hostname(ip_address):
    try:
        hostname = socket.gethostbyaddr(ip_address)[0]
    except Exception as e:
        log_verbose(f"Error getting hostname for {ip_address}: {e}")
        hostname = None
    return hostname
    
def display_network_map(network_data, target_network):
    root = tk.Tk()
    root.title("Network Map")

    paned_window = tk.PanedWindow(root, orient=tk.HORIZONTAL)
    paned_window.pack(fill=tk.BOTH, expand=1)

    # Create the initial networks pane with a scrollable canvas
    networks_canvas = tk.Canvas(paned_window)
    networks_scrollbar = tk.Scrollbar(paned_window, orient="vertical", command=networks_canvas.yview)
    networks_canvas.configure(yscrollcommand=networks_scrollbar.set)
    networks_scrollbar.pack(side="right", fill="y")
    networks_canvas.pack(side="left", fill="both", expand=True)

    networks_frame = tk.Frame(networks_canvas)
    networks_canvas.create_window((0, 0), window=networks_frame, anchor='nw')

    # Function to update the scroll region
    def on_frame_configure(event):
        networks_canvas.configure(scrollregion=networks_canvas.bbox("all"))

    networks_frame.bind("<Configure>", on_frame_configure)

    paned_window.add(networks_canvas, minsize=200)

    # Function to update the networks frame
    def update_networks_frame():
        # Clear the networks_frame
        for widget in networks_frame.winfo_children():
            widget.destroy()

        # Sort networks for consistent display
        sorted_networks = sorted(network_data['networks'].items())

        # Display networks as sections (boxes/containers)
        for network, devices in sorted_networks:
            net_frame = tk.LabelFrame(networks_frame, text=f"Network: {network}", padx=10, pady=10)
            net_frame.pack(fill="both", expand=True, padx=10, pady=10)

            # Add a rescan button
            rescan_button = tk.Button(net_frame, text="Rescan Network", relief="raised", anchor="w")
            rescan_button.grid(row=0, column=0, columnspan=2, sticky="we", padx=5, pady=2)
            rescan_button.bind("<Button-1>", partial(expand_item_event, item_ip_or_network=network, level=0))

            # Arrange devices in two columns
            row = 1
            col = 0
            for device_ip in devices:
                device_data = network_data['devices'].get(device_ip, {})
                device_name = device_data.get('hostname') or device_ip
                device_button = tk.Button(net_frame, text=device_name, relief="raised", anchor="w")
                device_button.grid(row=row, column=col, sticky="we", padx=5, pady=2)
                device_button.bind("<Button-1>", partial(expand_item_event, item_ip_or_network=device_ip, level=1))

                # Alternate between column 0 and 1
                if col == 0:
                    col = 1
                else:
                    col = 0
                    row += 1  # Move to the next row after filling both columns

            # Adjust column weights to make buttons expand evenly
            net_frame.columnconfigure(0, weight=1)
            net_frame.columnconfigure(1, weight=1)

    # Function to handle item expansion events
    def expand_item_event(event, item_ip_or_network, level):
        expand_item(item_ip_or_network, level)

    # Function to expand a device or network
    def expand_item(item_ip_or_network, level):
        # Clear panes to the right of the current level
        panes = paned_window.panes()
        for pane in panes[level+1:]:
            paned_window.forget(pane)

        # Check if item is a device or network
        if item_ip_or_network in network_data['devices']:
            # It's a device
            device_ip = item_ip_or_network
            device_data = network_data['devices'].get(device_ip, {})

            # Perform scan on the device
            current_time = time.time()
            last_checked = device_data.get('last_checked', 0)
            if current_time - last_checked > REFRESH_INTERVAL:
                log_verbose(f"Scanning device {device_ip}...")
                arp_entries = get_arp_table(device_ip)
                arp_entries = [entry for entry in arp_entries if not entry['is_ghost']]

                # Update device info
                existing_data = device_data
                existing_data.update({
                    'arp_entries': arp_entries,
                    'last_checked': current_time,
                    'hostname': get_hostname(device_ip)
                })
                network_data['devices'][device_ip] = existing_data

                # Save network data after updating the device
                save_network_data(network_data)

                if update_network_map([device_ip], network_data):
                    # Network data is saved inside update_network_map()
                    pass

            connections = [conn for conn in network_data['connections'] if conn['from'] == device_ip]
            if not connections:
                log_verbose(f"No connections found for device {device_ip}")
                return

            # Create a new frame for connections
            conn_frame = tk.Frame(paned_window)
            paned_window.add(conn_frame, minsize=200)

            # Make the connections frame scrollable
            conn_canvas = tk.Canvas(conn_frame)
            conn_scrollbar = tk.Scrollbar(conn_frame, orient="vertical", command=conn_canvas.yview)
            conn_canvas.configure(yscrollcommand=conn_scrollbar.set)
            conn_scrollbar.pack(side="right", fill="y")
            conn_canvas.pack(side="left", fill="both", expand=True)

            conn_inner_frame = tk.Frame(conn_canvas)
            conn_canvas.create_window((0, 0), window=conn_inner_frame, anchor='nw')

            # Update scroll region
            def on_conn_frame_configure(event):
                conn_canvas.configure(scrollregion=conn_canvas.bbox("all"))

            conn_inner_frame.bind("<Configure>", on_conn_frame_configure)

            tk.Label(conn_inner_frame, text=f"Connections for {device_data.get('hostname') or device_ip}", font=("Arial", 12, "bold")).pack()

            # Arrange connections in a vertical list
            for conn in connections:
                connected_ip = conn['to']
                connected_device = network_data['devices'].get(connected_ip, {})
                connected_name = connected_device.get('hostname') or connected_ip

                conn_button = tk.Button(conn_inner_frame, text=connected_name, relief="raised", anchor="w")
                conn_button.pack(padx=5, pady=2, fill="x")
                conn_button.bind("<Button-1>", partial(expand_item_event, item_ip_or_network=connected_ip, level=level+1))

            # Adjust column weights
            conn_inner_frame.columnconfigure(0, weight=1)
            conn_inner_frame.columnconfigure(1, weight=1)

        elif item_ip_or_network in network_data['networks']:
            # It's a network
            network = item_ip_or_network
            # Rescan the network
            rescan_network(network, network_data)
            # Update the display
            update_networks_frame()
        else:
            log_verbose(f"Item {item_ip_or_network} not found in network data.")

    # Function to rescan the network
    def rescan_network(network, network_data):
        log_verbose(f"Rescanning network {network}...")
        try:
            # Use arping to perform a fast ARP scan
            ans, unans = arping(network, timeout=0.5, verbose=0)
            arp_entries = []

            for snd, rcv in ans:
                ip = rcv.psrc
                mac = rcv.hwsrc.lower()  # Ensure MAC address is in lowercase for comparison
                try:
                    hostname = socket.gethostbyaddr(ip)[0]
                except (socket.herror, socket.gaierror):
                    hostname = None

                # Determine if the device is a ghost device
                is_ghost = False
                if hostname is None and mac == "00:00:0c:9f:f1:7e":
                    is_ghost = True

                log_verbose(f"Discovered device: IP={ip}, MAC={mac}, Hostname={hostname or 'Unknown'}, Ghost Device={is_ghost}")

                arp_entries.append({
                    'ip': ip,
                    'mac': mac,
                    'hostname': hostname,
                    'is_ghost': is_ghost,
                    'last_checked': time.time()
                })

            # Filter out ghost devices
            arp_entries = [entry for entry in arp_entries if not entry['is_ghost']]

            for entry in arp_entries:
                ip = entry['ip']
                mac = entry['mac']
                hostname = entry['hostname']
                last_checked = entry['last_checked']

                # Update device info
                existing_data = network_data['devices'].get(ip, {})
                existing_data.update({
                    'mac': mac,
                    'hostname': hostname,
                    'last_checked': last_checked
                })
                network_data['devices'][ip] = existing_data

                # Add the device to its respective network
                if network not in network_data['networks']:
                    network_data['networks'][network] = []
                if ip not in network_data['networks'][network]:
                    network_data['networks'][network].append(ip)

            # Save network data after rescan
            save_network_data(network_data)
        except Exception as e:
            log_verbose(f"Error during network rescan: {e}")


    # Initialize the networks frame
    update_networks_frame()

    root.mainloop()

def main():
    global VERBOSE
    parser = argparse.ArgumentParser(description="Network mapping script")
    parser.add_argument('-v', '--verbose', action='store_true', help='Enable verbose output')
    args = parser.parse_args()

    VERBOSE = args.verbose

    # Uncomment the following line to enforce administrative privileges
    # check_privileges()

    network_data = load_network_data()

    # If network_data.json is empty, perform a preliminary scan
    if not network_data['devices']:
        preliminary_scan(network_data)

    destination_ip = input("Enter the destination IP address to trace: ")
    max_hops = int(input("Enter the maximum number of hops: "))

    hops = perform_traceroute(destination_ip, max_hops)
    print(f"Discovered hops: {hops}")

    if update_network_map(hops, network_data):
        # Network data is saved inside update_network_map()
        pass

    network_prefix = '.'.join(destination_ip.split('.')[:3])
    target_network = f"{network_prefix}.0/24"

    print("Network map updated.")
    display_network_map(network_data, target_network)

if __name__ == "__main__":
    main()
