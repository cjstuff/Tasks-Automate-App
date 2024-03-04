import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar, Checkbutton
import requests
import csv
import os

def check_ip_reputation(ip, api_key, proxy_settings):
    url = f'https://www.virustotal.com/api/v3/ip_addresses/{ip}'
    headers = {
        'x-apikey': api_key,
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36 Edg/91.0.864.64'
    }
    try:
        response = requests.get(url, headers=headers, proxies=proxy_settings, verify=False)
        if response.status_code == 200:
            data = response.json()
            if 'data' in data:
                return data['data']['attributes']['last_analysis_stats']['malicious']
    except Exception as e:
        print(f"Failed to fetch reputation for {ip}: {e}")
    return None


def classify_reputation(count):
    if count is not None and count > 1:
        return "Malicious"
    else:
        return "Neutral"


def process_csv_file(csv_file, api_key, proxy_settings, progress_bar):
    results = []
    with open(csv_file, 'r') as file:
        reader = csv.reader(file)
        total_lines = sum(1 for line in file)
        file.seek(0)
        for i, row in enumerate(reader):
            ip = row[0]
            malicious_count = check_ip_reputation(ip, api_key, proxy_settings)
            if malicious_count is not None:
                reputation = classify_reputation(malicious_count)
                results.append((ip, malicious_count, reputation))
            progress = (i + 1) / total_lines * 100
            progress_bar['value'] = progress
            root.update_idletasks()
    return results


def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if file_path:
        csv_entry.delete(0, tk.END)
        csv_entry.insert(tk.END, file_path)


def toggle_proxy_entry():
    if proxy_checkbox_var.get():
        proxy_url_entry.config(state='normal')
        proxy_port_entry.config(state='normal')
    else:
        proxy_url_entry.config(state='disabled')
        proxy_port_entry.config(state='disabled')


def run_check():
    csv_file = csv_entry.get()
    api_key = api_key_entry.get()
    use_proxy = proxy_checkbox_var.get()
    proxy_url = proxy_url_entry.get()
    proxy_port = proxy_port_entry.get()

    proxy_settings = None
    if use_proxy:
        if not proxy_url or not proxy_port:
            messagebox.showerror("Error", "Please enter both proxy URL and port.")
            return
        proxy_settings = {'http': f'http://{proxy_url}:{proxy_port}', 'https': f'https://{proxy_url}:{proxy_port}'}

    try:
        progress_bar['value'] = 0
        results = process_csv_file(csv_file, api_key, proxy_settings, progress_bar)
        if results:
            output_file_path = os.path.join(os.path.dirname(csv_file), 'output.csv')
            with open(output_file_path, 'w', newline='') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(['IP', 'Malicious Count', 'Reputation'])
                for ip, malicious_count, reputation in results:
                    writer.writerow([ip, malicious_count, reputation])
            messagebox.showinfo("Information", f"CSV file saved successfully at {output_file_path}.")
        else:
            messagebox.showinfo("Information", "No results found.")
    finally:
        progress_bar['value'] = 100


# Create main window
root = tk.Tk()
root.title("IP Reputation Checker")

# Create CSV file input
csv_label = tk.Label(root, text="CSV File:")
csv_label.pack(pady=(10, 0))
csv_entry = tk.Entry(root, width=50)
csv_entry.pack(padx=10)

browse_button = tk.Button(root, text="Browse", command=browse_file)
browse_button.pack(pady=5)

# Create API key input
api_key_label = tk.Label(root, text="API Key:")
api_key_label.pack(pady=(10, 0))
api_key_entry = tk.Entry(root, width=50)
api_key_entry.pack(padx=10)

# Create proxy checkbox
proxy_checkbox_var = tk.BooleanVar()
proxy_checkbox = tk.Checkbutton(root, text="Use Proxy", variable=proxy_checkbox_var, command=toggle_proxy_entry)
proxy_checkbox.pack(pady=(10, 0))

# Create proxy URL entry
proxy_url_label = tk.Label(root, text="Proxy URL:")
proxy_url_label.pack(pady=(10, 0))
proxy_url_entry = tk.Entry(root, width=50)
proxy_url_entry.pack(padx=10)
proxy_url_entry.config(state='disabled')  # Initially disabled

# Create proxy port entry
proxy_port_label = tk.Label(root, text="Proxy Port:")
proxy_port_label.pack(pady=(10, 0))
proxy_port_entry = tk.Entry(root, width=50)
proxy_port_entry.pack(padx=10)
proxy_port_entry.config(state='disabled')  # Initially disabled

# Create run button
run_button = tk.Button(root, text="Run Check", command=run_check)
run_button.pack(pady=10)

# Create progress bar
progress_bar = Progressbar(root, orient=tk.HORIZONTAL, length=200, mode='determinate')
progress_bar.pack(pady=10)

# Run the main event loop
root.mainloop()
