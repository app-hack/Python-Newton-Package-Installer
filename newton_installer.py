import serial
import serial.tools.list_ports  # Added for auto-enumeration
import time
import threading
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

# --- Constants & Protocol Definitions ---
VERSION = "0.1A-Py"
MAX_INFO_LEN = 256
TIMEOUT = 30

FRAME_START = b'\x16\x10\x02'
FRAME_END = b'\x10\x03'

LR_FRAME = bytes([
    0x17, 0x01, 0x02, 0x01, 0x06, 0x01, 0x00, 0x00, 
    0x00, 0x00, 0xff, 0x02, 0x01, 0x02, 0x03, 0x01, 
    0x01, 0x04, 0x02, 0x40, 0x00, 0x08, 0x01, 0x03
])

class NewtonInstallerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title(f"Python Newton Installer - v{VERSION}")
        self.port = None
        self.lt_seq_no = 0
        self.file_list = []
        
        self.setup_ui()
        self.refresh_ports() # Auto-populate on startup
        
    def setup_ui(self):
        # Configuration Frame
        cfg_frame = ttk.LabelFrame(self.root, text="Connection Settings", padding="10")
        cfg_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(cfg_frame, text="Serial Port:").grid(row=0, column=0, sticky=tk.W)
        
        # Changed from Entry to Combobox for Enumeration
        self.port_combo = ttk.Combobox(cfg_frame, width=25)
        self.port_combo.grid(row=0, column=1, padx=5)
        
        ttk.Button(cfg_frame, text="↻", width=3, command=self.refresh_ports).grid(row=0, column=2)
        
        ttk.Label(cfg_frame, text="Speed:").grid(row=0, column=3, sticky=tk.W, padx=(10,0))
        self.speed_combo = ttk.Combobox(cfg_frame, values=["9600", "19200", "38400", "57600", "115200"], width=10)
        self.speed_combo.set("38400")
        self.speed_combo.grid(row=0, column=4, padx=5)

        # File Frame
        file_frame = ttk.LabelFrame(self.root, text="Package Queue", padding="10")
        file_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.file_label = ttk.Label(file_frame, text="No files selected.")
        self.file_label.pack(side=tk.LEFT, padx=5)
        ttk.Button(file_frame, text="Add Files", command=self.add_files).pack(side=tk.RIGHT)

        # Debug Output
        self.debug_text = tk.Text(self.root, height=12, state='disabled', background="#1e1e1e", foreground="#00ff00")
        self.debug_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill=tk.X, padx=10)
        ttk.Button(btn_frame, text="Copy Logs", command=self.copy_debug).pack(side=tk.LEFT, pady=5)
        self.start_btn = ttk.Button(btn_frame, text="Start Installation", command=self.start_thread)
        self.start_btn.pack(side=tk.RIGHT, pady=5)

        self.progress = ttk.Progressbar(self.root, orient=tk.HORIZONTAL, mode='determinate')
        self.progress.pack(fill=tk.X, padx=10, pady=10)

    def refresh_ports(self):
        """Automatically find all available serial ports."""
        ports = serial.tools.list_ports.comports()
        port_list = [p.device for p in ports]
        self.port_combo['values'] = port_list
        if port_list:
            self.port_combo.set(port_list[0])
        else:
            self.port_combo.set("No ports found")
        self.log(f"Enumerated {len(port_list)} ports.")

    # --- Communication Logic (Mirroring your C code exactly) ---
    def fcs_calc(self, fcs_word, octet):
        for _ in range(8):
            if ((fcs_word % 256) & 0x01 == 0x01) ^ ((octet & (1 << _)) != 0):
                fcs_word = (fcs_word // 2) ^ 0xa001
            else:
                fcs_word //= 2
        return fcs_word

    def send_frame(self, head, info=None):
        fcs_word = 0
        self.port.write(FRAME_START)
        for b in head:
            fcs_word = self.fcs_calc(fcs_word, b)
            self.port.write(bytes([b]))
            if b == 0x10: self.port.write(bytes([0x10]))
        if info:
            for b in info:
                fcs_word = self.fcs_calc(fcs_word, b)
                self.port.write(bytes([b]))
                if b == 0x10: self.port.write(bytes([0x10]))
        self.port.write(FRAME_END)
        fcs_word = self.fcs_calc(fcs_word, FRAME_END[1])
        self.port.write(bytes([fcs_word % 256, fcs_word // 256]))

    def recv_frame(self):
        fcs_word = 0
        frame_data = bytearray()
        state = 0
        while state < 3:
            b = self.port.read(1)
            if not b: return None
            if state == 0 and b[0] == 0x16: state = 1
            elif state == 1 and b[0] == 0x10: state = 2
            elif state == 2 and b[0] == 0x02: state = 3
            else: state = 0

        state = 0
        while state < 2:
            b = self.port.read(1)
            if not b: return None
            val = b[0]
            if state == 0:
                if val == 0x10: state = 1
                else:
                    fcs_word = self.fcs_calc(fcs_word, val)
                    frame_data.append(val)
            elif state == 1:
                if val == 0x10: 
                    fcs_word = self.fcs_calc(fcs_word, val)
                    frame_data.append(val)
                    state = 0
                elif val == 0x03:
                    fcs_word = self.fcs_calc(fcs_word, val)
                    state = 2
        
        fcs_in = self.port.read(2)
        if len(fcs_in) < 2: return None
        if (fcs_word % 256 != fcs_in[0]) or (fcs_word // 256 != fcs_in[1]):
            return None
        return frame_data

    def wait_la_frame(self, seq):
        while True:
            recv = self.recv_frame()
            if not recv: continue
            if recv[1] == 0x04: self.send_la_frame(recv[2]) 
            if recv[1] == 0x05 and recv[2] == seq: return True

    def send_la_frame(self, seq):
        self.send_frame(bytes([0x03, 0x05, seq, 0x01]))

    def send_lt_frame(self, info):
        self.send_frame(bytes([0x02, 0x04, self.lt_seq_no]), info)

    # --- Installation Engine ---
    def run_installer(self):
        try:
            selected_port = self.port_combo.get()
            if not selected_port or selected_port == "No ports found":
                raise Exception("Please select a valid serial port.")

            self.port = serial.Serial(selected_port, int(self.speed_combo.get()), timeout=2)
            self.log(f"Port {selected_port} opened. Waiting for Newton...")
            
            while True:
                recv = self.recv_frame()
                if recv and recv[1] == 0x01: break
            
            self.log("Connected! Sending LR Frame...")
            self.send_frame(LR_FRAME)
            self.wait_la_frame(self.lt_seq_no)
            self.lt_seq_no += 1

            while True:
                recv = self.recv_frame()
                if recv and recv[1] == 0x04:
                    self.send_la_frame(recv[2])
                    break
            
            self.log("Sending dockdock handshake...")
            self.send_lt_frame(b"newtdockdock\x00\x00\x00\x04\x00\x00\x00\x04")
            self.wait_la_frame(self.lt_seq_no)
            self.lt_seq_no += 1

            while True:
                recv = self.recv_frame()
                if recv and recv[1] == 0x04:
                    self.send_la_frame(recv[2])
                    try:
                        name_bytes = recv[24:]
                        name = ""
                        for i in range(0, len(name_bytes), 2):
                            if name_bytes[i] == 0: break
                            name += chr(name_bytes[i])
                        self.log(f"Newton Owner: {name}")
                    except: pass
                    break

            self.send_lt_frame(b"newtdockstim\x00\x00\x00\x04\x00\x00\x00\x1e")
            self.wait_la_frame(self.lt_seq_no)
            self.lt_seq_no += 1

            while True:
                recv = self.recv_frame()
                if recv and recv[1] == 0x04:
                    self.send_la_frame(recv[2])
                    break

            for path in self.file_list:
                with open(path, "rb") as f:
                    data = f.read()
                    f_len = len(data)
                    self.log(f"Uploading {path}...")
                    lpkg = b"newtdocklpkg" + f_len.to_bytes(4, byteorder='big')
                    self.send_lt_frame(lpkg)
                    self.wait_la_frame(self.lt_seq_no)
                    self.lt_seq_no += 1

                    for i in range(0, f_len, MAX_INFO_LEN):
                        chunk = data[i:i+MAX_INFO_LEN]
                        while len(chunk) % 4 != 0: chunk += b'\x00'
                        self.send_lt_frame(chunk)
                        self.wait_la_frame(self.lt_seq_no)
                        self.lt_seq_no = (self.lt_seq_no + 1) % 256
                        self.progress['value'] = (i / f_len) * 100
                        self.root.update_idletasks()
                    
                    self.log(f"Finished {path}")

            self.send_lt_frame(b"newtdockdisc\x00\x00\x00\x00")
            self.log("Session Finished.")
            self.port.close()

        except Exception as e:
            self.log(f"ERROR: {str(e)}")
        finally:
            self.start_btn.config(state='normal')

    def log(self, msg):
        self.debug_text.config(state='normal')
        self.debug_text.insert(tk.END, f"{msg}\n")
        self.debug_text.see(tk.END)
        self.debug_text.config(state='disabled')

    def add_files(self):
        files = filedialog.askopenfilenames(filetypes=[("Package", "*.pkg")])
        if files:
            self.file_list = list(files)
            self.file_label.config(text=f"{len(files)} files queued.")

    def copy_debug(self):
        self.root.clipboard_clear()
        self.root.clipboard_append(self.debug_text.get("1.0", tk.END))

    def start_thread(self):
        if not self.file_list: return
        self.start_btn.config(state='disabled')
        threading.Thread(target=self.run_installer, daemon=True).start()

if __name__ == "__main__":
    root = tk.Tk()
    app = NewtonInstallerGUI(root)
    root.mainloop()
