import os
import random
import re
import string
import time
from random import randint, uniform
import tkinter as tk
from tkinter import filedialog, Text, messagebox, ttk
import pandas as pd
import asyncio
import websockets
import threading
import pickle

def resp_automation(email: str, subject: str, body: str):
    #subject = f"{subject}_{randint(1000000, 99999999999)}"
    body = body.replace("\\", "\\\\").replace("'", "\\'").replace('"', '\\"').replace("\n", "\\n")
    js_code = f"""
    (function() {{
        const desiredData = {{
            email: '{email}',
            subject: '{subject}',
            body: '{body}'
        }};
    
        const waitForElement = (selector, callback) => {{
            const el = document.querySelector(selector);
            if (el) {{
                callback(el);
            }} else {{
                setTimeout(() => waitForElement(selector, callback), 1500);
            }}
        }};
    
        const fillAndSendEmail = () => {{
            waitForElement('.agP.aFw', (toInput) => {{
                toInput.value = desiredData.email;
                setTimeout(() => {{
                    const subjectInput = document.querySelector('.aoT');
                    if (subjectInput) {{
                        subjectInput.value = desiredData.subject;
                        waitForElement('div[aria-label="Message Body"]', (messageBodyDiv) => {{
                            messageBodyDiv.innerHTML = desiredData.body;
                            waitForElement('div[aria-label="Send ‪(Ctrl-Enter)‬"]', (sendButton) => {{
                                if (toInput.value && messageBodyDiv.innerHTML) {{
                                    sendButton.click();
                                }}
                            }});
                        }});
                    }} else {{
                        console.error('Subject input not found.');
                    }}
                }}, 250);
            }});
        }};
    
        const composeButton = document.querySelector('.T-I.T-I-KE.L3');
        if (composeButton) {{
            composeButton.click();
            fillAndSendEmail();
        }} else {{
            console.error("Compose button not found.");
        }}
    }})();"""

    return js_code
class App:
    def __init__(self, root):
        self.saved_values_path = "saved_values.pkl"
        self.root = root
        self.root.title("Mail Automation GUI")
        self.clients = []

        # Top Section - Frame Initialization
        top_frame = tk.Frame(root)
        top_frame.pack(pady=20, fill=tk.X)

        # Top Left Frame for Labels
        top_left_frame = tk.Frame(top_frame)
        top_left_frame.pack(side=tk.LEFT, padx=10)

        # Top Right Frame for Entries/Buttons
        top_right_frame = tk.Frame(top_frame)
        top_right_frame.pack(side=tk.RIGHT, padx=10)

        # UI for CSV Recipients in Top Section
        self.recipients_label = tk.Label(top_left_frame, text="Select Recipients CSV:")
        self.recipients_label.pack(anchor="w")
        self.recipients_entry = tk.Entry(top_right_frame)
        self.recipients_entry.pack(fill=tk.X)
        self.recipients_btn = tk.Button(top_right_frame, text="Browse", command=self.load_recipients)
        self.recipients_btn.pack()

        # UI for Mail Subject in Top Section
        self.subject_label = tk.Label(top_left_frame, text="Subject:")
        self.subject_label.pack(anchor="w", pady=(10,0))
        self.subject_entry = tk.Entry(top_right_frame)
        self.subject_entry.pack(fill=tk.X, pady=(10,0))

        # UI for Number of Tabs in Top Section
        self.tabs_label = tk.Label(top_left_frame, text="Number of Tabs:")
        self.tabs_label.pack(anchor="w", pady=(10,0))
        self.tabs_entry = tk.Entry(top_right_frame)
        self.tabs_entry.pack(fill=tk.X, pady=(10,0))

        # Middle Section - For HTML Content
        self.content_label = tk.Label(root, text="HTML Content:")
        self.content_label.pack(pady=(20,5), anchor="w", padx=20)
        self.content_text = Text(root, height=10)
        self.content_text.pack(fill=tk.BOTH, padx=20, pady=(0,10))

        # Bottom Section - UI for Sending mails
        self.send_btn = tk.Button(root, text="Send Mails", command=self.send_mails)
        self.send_btn.pack(pady=10)
        #Start Server Button
        button_frame = tk.Frame(root)
        button_frame.pack(pady=20)

        self.start_server_btn = tk.Button(button_frame, text="Start Server", command=self.start_server)
        self.start_server_btn.pack(side=tk.LEFT, padx=10)

        self.update_table_btn = tk.Button(button_frame, text="Refresh", command=self.update_table)
        self.update_table_btn.pack(side=tk.LEFT, padx=10)

        # Bottom Section - UI for Table showing status
        bottom_frame = tk.Frame(root)
        bottom_frame.pack(pady=20, fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(bottom_frame, columns=("Tab", "Mails Sent"), show="headings")
        self.tree.heading("Tab", text="Tab")
        self.tree.heading("Mails Sent", text="Mails Sent")
        self.tree.pack(side=tk.TOP, pady=(0,10), fill=tk.BOTH, expand=True)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # UI for Cool Down Delay in Top Section
        self.cool_down_delay_label = tk.Label(top_left_frame, text="Cool Down Delay (seconds):")
        self.cool_down_delay_label.pack(anchor="w", pady=(10,0))
        self.cool_down_delay_entry = tk.Entry(top_right_frame)
        self.cool_down_delay_entry.insert(0, '3')  # default value set to 3
        self.cool_down_delay_entry.pack(fill=tk.X, pady=(10,0))

        # UI for Per Mail Speed in Top Section
        self.per_mail_speed_label = tk.Label(top_left_frame, text="Per Mail Speed (seconds):")
        self.per_mail_speed_label.pack(anchor="w", pady=(10,0))
        self.per_mail_speed_entry = tk.Entry(top_right_frame)
        self.per_mail_speed_entry.insert(0, '5')  # default value set to 2
        self.per_mail_speed_entry.pack(fill=tk.X, pady=(10,0))

        self.load_saved_values()

    def load_saved_values(self):
     if os.path.exists(self.saved_values_path):
        try:
            with open(self.saved_values_path, "rb") as f:
                saved_values = pickle.load(f)
                self.recipients_entry.insert(0, saved_values.get("recipients", ""))
                self.subject_entry.insert(0, saved_values.get("subject", ""))
                self.content_text.insert("1.0", saved_values.get("content", ""))
        except Exception as e:
            print(f"Error loading saved values: {e}")

    def on_closing(self):
    # Save current input values
     saved_values = {
        "recipients": self.recipients_entry.get(),
        "subject": self.subject_entry.get(),
        "content": self.content_text.get("1.0", "end-1c")  # 'end-1c' removes the trailing newline
     }
     with open(self.saved_values_path, "wb") as f:
        pickle.dump(saved_values, f)
    
     self.root.destroy()

    def load_recipients(self):
        file_path = filedialog.askopenfilename()
        if file_path:
            self.recipients_entry.delete(0, tk.END)
            self.recipients_entry.insert(0, file_path)

    def start_server(self):
     try:
         thread = threading.Thread(target=self.run_server)
         thread.start() 
         print(f"Number of connected clients: {len(self.clients)}")
         self.start_server_btn.config(state=tk.DISABLED)
     except Exception as e:
         messagebox.showerror("Error", str(e))


    def handle_placeholders(self, content: str, email: str, name: str, index: int, address: str = "") -> str:
        if content is None:
            raise ValueError("The content argument is None. Please provide valid content.")
        invoice = str(randint(1000000000, 9999999999))
        # Placeholder options
        amt_id_options = ["Final Amount", "Total Amount", "Amount", "Final Price", "Total Price"]
        qty_id_options = ["Qty", "Quantity"]
        tax_id_options = ["VAT", "TAX", "Duty"]
        invoice_id_options = ["INV0ICE", "Inv", "Ref Id","Memo ID","Recp Id"]
        salutation_options = ["Hi", "Hello", "Dear", "Greetings"]
        time_options = ["24","48","12","36","72"]
        email_optionns = ["Email","Customer_Email","User_Mail","Official_Mail","Registered_User","Client_mail","Contact"]
        product_options =["Bitcoin ","BTC","Crypto","E-Currency","E-Coin"]
        prefix_options =["Purchase","Order","Transaction","Buyout","Procurement"]
        productID_options =["Description","Summary","Item","Product"]
        rate = 28888
        qty_no_value = round(uniform(0.010000, 0.014999), 6)
        final_price = qty_no_value * rate
        amount_before_tax = final_price
        tax_no_value = round(uniform(0.99, 3.99), 2)
        final_price += tax_no_value
        content = content.replace("$amt_id", random.choice(amt_id_options))
        content = content.replace("$qty_id", random.choice(qty_id_options))
        content = content.replace("$tax_id", random.choice(tax_id_options))
        content = content.replace("$invoice_id", random.choice(invoice_id_options))
        content = content.replace("$salutation", random.choice(salutation_options))
        content = content.replace("$time", random.choice(time_options))
        content = content.replace("$email_id", random.choice(email_optionns))
        content = content.replace("$product_id", random.choice(productID_options))
        content = content.replace("$product", random.choice(product_options) + " " + random.choice(prefix_options))
        content = content.replace("$qty_no", str(qty_no_value))
        content = content.replace("$tax_no", "$" + str(tax_no_value))
        content = content.replace("$price", "$" + str(amount_before_tax))
        content = content.replace("$amt_no", "${:.2f}".format(final_price))
        content = content.replace("$invoice_no", invoice)
        content = content.replace("$addr", address)
        content = content.replace("$cust_email", email)
        content = content.replace("$name", name)
        alignments = ['left', 'center', 'right']
        cheader_files = sorted([x for x in os.listdir('CHEADERS') if x.endswith(".txt")])
        cheader_path = cheader_files[index % len(cheader_files)]
        alignment = alignments[index % len(alignments)]
        try:
            with open(os.path.join('CHEADERS', cheader_path), 'r') as file:
                cheader_content = file.read().strip()
            order_no = ''.join(random.choice(string.ascii_uppercase + string.digits) for _ in range(12))
            cheader_content = cheader_content.replace("$orderno", order_no)
            styled_cheader_content = f'<p style="text-align: {alignment};"><strong>{cheader_content}</strong></p>'
            content = content.replace("<p>$chead</p>", styled_cheader_content)
        except Exception as e:
            print(f"Error reading from CHEADERS folder: {str(e)}")
        align_change_lines = re.findall('<.*? id="\$alignchng_\d+".*?>.*?<.*?>', content)
        if align_change_lines:
            for line in align_change_lines:
                alignment = random.choice(alignments)
                new_line = re.sub('id="\$alignchng_\d+"', f'style="text-align: {alignment};"', line, 1)
                content = content.replace(line, new_line)
        if "$jaadu" in content:
            charset = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()></*-+"
            replacement = ''.join(random.choice(charset) for _ in range(15))
            content = content.replace("$jaadu", replacement)
        return content


    def generate_alphanumeric(self, length: int) -> str:
        """Generate a random alphanumeric string of the specified length."""
        chars = string.ascii_letters + string.digits
        return ''.join(random.choice(chars) for _ in range(length))

    def send_mails_to_client(self, client, data, subject, content_template):
     try:
        email_count = 0  # Initialize the counter
        for index, (email, address, name) in enumerate(data): 
            content = self.handle_placeholders(content_template, email, name, index, address)
            message = resp_automation(email, subject, content)
            asyncio.run_coroutine_threadsafe(self.serve_dynamic_mail(client, message), client['loop'])
            client['counter'] += 1
            email_count += 1  # Increment the counter after sending an email
            cool_down_delay = float(self.cool_down_delay_entry.get() or "4")  # Default is 4 seconds if left empty
            per_mail_speed = float(self.per_mail_speed_entry.get() or "2")  # Default is 2 seconds if left empty

            if email_count % 15 == 0:  # Check if 10 emails have been sent
             time.sleep(cool_down_delay)
            time.sleep(per_mail_speed)
     except Exception as e:
        print("Error in send_mails_to_client:", str(e))


    def send_mails(self):
        recipients_file = self.recipients_entry.get()
        subject = self.subject_entry.get()
        content = self.content_text.get("1.0", tk.END)
        try:
            df_contacts = pd.read_csv(recipients_file)  
            if 'email' not in df_contacts.columns or 'address' not in df_contacts.columns or 'name' not in df_contacts.columns:
                messagebox.showerror("Error", "The CSV file must have 'email', 'address', and 'name' columns.")
                return

            num_tabs = int(self.tabs_entry.get())
            distributed_emails_list = self.distribute_emails(df_contacts, num_tabs)
            for i, subset_emails_with_addresses in enumerate(distributed_emails_list):
                print(f"Client {i} assigned emails: {subset_emails_with_addresses}")
                if i < len(self.clients):
                    client = self.clients[i]
                    threading.Thread(target=self.send_mails_to_client, args=(client, subset_emails_with_addresses, subject, content)).start()
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def distribute_emails(self, df_contacts, num_clients):
        df_contacts = df_contacts.dropna(subset=['email'])
        emails_per_client = len(df_contacts) // num_clients
        extra_emails = len(df_contacts) % num_clients
        start_idx = 0
        distributed_data = []
        for i in range(num_clients):
            end_idx = start_idx + emails_per_client
            if extra_emails > 0:
                end_idx += 1
                extra_emails -= 1
            subset_df = df_contacts.iloc[start_idx:end_idx]
            distributed_data.append(subset_df[['email', 'address', 'name']].values.tolist())
            start_idx = end_idx
        return distributed_data
 
    def update_table(self):
     for row in self.tree.get_children():
        self.tree.delete(row)
     self.clients = [client for client in self.clients if client['ws'].open]
     for i, client in enumerate(self.clients):
        self.tree.insert('', 'end', values=(f"Tab {i+1}", client.get('counter', 0)))

    def run_server(self):
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        start_server = websockets.serve(self.register_client, "localhost", 5678)
        loop.run_until_complete(start_server)
        loop.run_forever()

    async def register_client(self, websocket, path):
        client = {'ws': websocket, 'loop': asyncio.get_event_loop(), 'data': None, 'counter': 0}
        self.clients.append(client)
        try:
            async for _ in websocket:
                pass
        except:
            pass
        finally:
            self.clients = [client for client in self.clients if client['ws'].open]  # Cleanup disconnected clients

    async def serve_dynamic_mail(self, client, message):
     await client['ws'].send(message)

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
    app_instance = App()
    thread = threading.Thread(target=app_instance.run_server)
    thread.start()

