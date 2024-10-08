import os
import platform
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from threading import Thread
import socket
import speedtest
import csv
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
import threading
import json
import openpyxl
from plyer import notification
import configparser

PING_HOST = '8.8.8.8'
SPEED_TEST_SERVERS = {
    'Default': None,
    'New York': 10556,
    'London': 6032,
    'Tokyo': 6087
}

class NetworkPerformanceTool:
    def __init__(self, master):
        self.master = master
        master.title("Advanced Network Performance Tool")
        master.geometry("900x700")

        self.notebook = ttk.Notebook(master)
        self.notebook.pack(expand=True, fill='both', padx=10, pady=10)

        self.results_frame = ttk.Frame(self.notebook)
        self.graph_frame = ttk.Frame(self.notebook)

        self.notebook.add(self.results_frame, text='Results')
        self.notebook.add(self.graph_frame, text='Graphs')

        self.setup_results_frame()
        self.setup_graph_frame()

        self.history = []

        self.auto_test_running = False
        self.auto_test_timer = None

        # إضافة عناصر واجهة المستخدم الجديدة
        self.auto_test_frame = ttk.Frame(self.results_frame)
        self.auto_test_frame.pack(pady=5)

        self.interval_label = ttk.Label(self.auto_test_frame, text="Test Interval (seconds):")
        self.interval_label.pack(side=tk.LEFT)

        self.interval_entry = ttk.Entry(self.auto_test_frame, width=5)
        self.interval_entry.insert(0, "60")  # القيمة الافتراضية هي 60 ثانية (دقيقة واحدة)
        self.interval_entry.pack(side=tk.LEFT, padx=5)

        self.warning_label = ttk.Label(self.auto_test_frame, text="", foreground="red")
        self.warning_label.pack(side=tk.LEFT, padx=5)

        self.auto_test_button = ttk.Button(self.auto_test_frame, text="Start Auto Test", command=self.toggle_auto_test)
        self.auto_test_button.pack(side=tk.LEFT, padx=5)

        self.setup_menu()
        self.load_settings()

    def setup_results_frame(self):
        self.output_text = tk.Text(self.results_frame, height=25, width=80)
        self.output_text.pack(padx=10, pady=10)

        self.server_var = tk.StringVar(value='Default')
        server_menu = ttk.OptionMenu(self.results_frame, self.server_var, 'Default', *SPEED_TEST_SERVERS.keys())
        server_menu.pack(pady=5)

        self.run_button = ttk.Button(self.results_frame, text="Run Tests", command=self.run_tests)
        self.run_button.pack(pady=5)

        self.save_button = ttk.Button(self.results_frame, text="Save Results", command=self.save_results)
        self.save_button.pack(pady=5)

        self.save_pdf_button = ttk.Button(self.results_frame, text="Save PDF Report", command=self.save_pdf_report)
        self.save_pdf_button.pack(pady=5)

        self.progress_bar = ttk.Progressbar(self.results_frame, length=300, mode='determinate')
        self.progress_bar.pack(pady=10)

    def setup_graph_frame(self):
        self.figure = Figure(figsize=(8, 8))
        self.ax1 = self.figure.add_subplot(211)
        self.ax2 = self.figure.add_subplot(212)
        
        self.canvas = FigureCanvasTkAgg(self.figure, master=self.graph_frame)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)

    def toggle_auto_test(self):
        if self.auto_test_running:
            self.stop_auto_test()
        else:
            self.start_auto_test()

    def start_auto_test(self):
        try:
            interval = float(self.interval_entry.get())
            if interval < 1:
                raise ValueError("Interval must be at least 1 second.")
            if interval < 10:
                self.warning_label.config(text="Warning: Short intervals may affect network performance.")
            else:
                self.warning_label.config(text="")
        except ValueError as e:
            tk.messagebox.showerror("Error", str(e))
            return

        self.auto_test_running = True
        self.auto_test_button.config(text="Stop Auto Test")
        self.run_tests()
        self.schedule_next_test(interval)

    def stop_auto_test(self):
        self.auto_test_running = False
        self.auto_test_button.config(text="Start Auto Test")
        if self.auto_test_timer:
            self.auto_test_timer.cancel()

    def schedule_next_test(self, interval):
        if self.auto_test_running:
            self.auto_test_timer = threading.Timer(interval, self.run_auto_test)
            self.auto_test_timer.start()

    def run_auto_test(self):
        self.run_tests()
        try:
            interval = float(self.interval_entry.get())
            if interval < 1:
                raise ValueError("Interval must be at least 1 second.")
            if interval < 10:
                self.warning_label.config(text="Warning: Short intervals may affect network performance.")
            else:
                self.warning_label.config(text="")
            self.schedule_next_test(interval)
        except ValueError as e:
            self.stop_auto_test()
            tk.messagebox.showerror("Error", str(e) + " Auto test stopped.")

    def run_tests(self):
        self.run_button.config(state='disabled')
        self.output_text.delete(1.0, tk.END)
        self.progress_bar['value'] = 0
        
        def run():
            results = self.perform_tests()
            self.history.append(results)
            self.master.after(0, self.update_output, results)
            self.master.after(0, self.update_graphs)
            self.check_for_notifications(results)
        
        Thread(target=run).start()

    def update_output(self, results):
        self.output_text.insert(tk.END, f"Test Time: {results['timestamp']}\n\n")
        self.output_text.insert(tk.END, "Ping Test Results:\n")
        self.output_text.insert(tk.END, results['ping']['output'] + "\n")
        self.output_text.insert(tk.END, self.interpret_ping(results['ping']) + "\n\n")

        self.output_text.insert(tk.END, "DNS Lookup Results:\n")
        self.output_text.insert(tk.END, results['dns']['output'] + "\n")
        self.output_text.insert(tk.END, self.interpret_dns(results['dns']) + "\n\n")

        self.output_text.insert(tk.END, "Speed Test Results:\n")
        self.output_text.insert(tk.END, f"Download Speed: {results['speed']['download']:.2f} Mbps\n")
        self.output_text.insert(tk.END, f"Upload Speed: {results['speed']['upload']:.2f} Mbps\n")
        self.output_text.insert(tk.END, f"Ping: {results['speed']['ping']:.2f} ms\n")
        self.output_text.insert(tk.END, self.interpret_speed(results['speed']) + "\n")

        self.run_button.config(state='normal')
        self.update_graphs()

    def update_progress(self, value):
        self.progress_bar['value'] = value

    def ping(self, host):
        param = '-n' if platform.system().lower() == 'windows' else '-c'
        command = f"ping {param} 4 {host}"
        try:
            result = subprocess.run(command, stdout=subprocess.PIPE, shell=True, text=True)
            output = result.stdout
            avg_time = float(output.split('Average = ')[-1].split('ms')[0])
            return {'output': output, 'avg_time': avg_time}
        except Exception as e:
            return {'output': f"Error during ping test: {str(e)}", 'avg_time': None}

    def dns_lookup(self, domain):
        try:
            ip = socket.gethostbyname(domain)
            return {'output': f"IP address of {domain}: {ip}", 'success': True}
        except Exception as e:
            return {'output': f"Error during DNS lookup: {str(e)}", 'success': False}

    def speed_test(self):
        try:
            st = speedtest.Speedtest()
            server_id = SPEED_TEST_SERVERS[self.server_var.get()]
            if server_id:
                st.get_servers([server_id])
            else:
                st.get_best_server()
            download_speed = st.download() / 1_000_000
            upload_speed = st.upload() / 1_000_000
            ping = st.results.ping
            return {'download': download_speed, 'upload': upload_speed, 'ping': ping}
        except Exception as e:
            return {'download': 0, 'upload': 0, 'ping': 0, 'error': str(e)}

    def interpret_ping(self, result):
        if result['avg_time'] is not None:
            if result['avg_time'] < 50:
                return "Interpretation: Excellent ping time."
            elif result['avg_time'] < 100:
                return "Interpretation: Good ping time for most applications."
            else:
                return "Interpretation: High ping time might affect real-time applications."
        else:
            return "Interpretation: Unable to determine ping time."

    def interpret_dns(self, result):
        if result['success']:
            return "Interpretation: DNS lookup successful. Your DNS is working correctly."
        else:
            return "Interpretation: DNS lookup failed. There might be an issue with your DNS server or internet connection."

    def interpret_speed(self, result):
        if 'error' in result:
            return f"Interpretation: Speed test failed. Error: {result['error']}"
        else:
            if result['download'] > 100 and result['upload'] > 50:
                return "Interpretation: Excellent internet speed."
            elif result['download'] > 25 and result['upload'] > 10:
                return "Interpretation: Good internet speed for most applications."
            else:
                return "Interpretation: Your internet speed might be insufficient for some applications."

    def save_results(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                                 filetypes=[("CSV files", "*.csv")])
        if file_path:
            with open(file_path, 'w', newline='') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(['Timestamp', 'Ping (ms)', 'DNS Lookup', 'Download Speed (Mbps)', 'Upload Speed (Mbps)', 'Speedtest Ping (ms)'])
                for result in self.history:
                    writer.writerow([
                        result['timestamp'],
                        result['ping']['avg_time'],
                        'Success' if result['dns']['success'] else 'Failure',
                        result['speed']['download'],
                        result['speed']['upload'],
                        result['speed']['ping']
                    ])
            tk.messagebox.showinfo("Save Successful", f"Results saved to {file_path}")

    def update_graphs(self):
        if len(self.history) < 2:  # نحتاج على الأقل نقطتين لرسم خط
            return

        self.ax1.clear()
        self.ax2.clear()

        timestamps = [result['timestamp'] for result in self.history]
        ping_times = [result['ping']['avg_time'] for result in self.history if result['ping']['avg_time'] is not None]
        download_speeds = [result['speed']['download'] for result in self.history]
        upload_speeds = [result['speed']['upload'] for result in self.history]

        if ping_times:
            self.ax1.plot(timestamps, ping_times, 'b-', label='Ping Time')
            self.ax1.set_ylabel('Ping Time (ms)')
            self.ax1.set_title('Ping Time History')
            self.ax1.legend()
            self.ax1.tick_params(axis='x', rotation=45)

        self.ax2.plot(timestamps, download_speeds, 'g-', label='Download')
        self.ax2.plot(timestamps, upload_speeds, 'r-', label='Upload')
        self.ax2.set_ylabel('Speed (Mbps)')
        self.ax2.set_title('Internet Speed History')
        self.ax2.legend()
        self.ax2.tick_params(axis='x', rotation=45)

        self.figure.tight_layout()
        self.canvas.draw()

    def save_pdf_report(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".pdf",
                                                 filetypes=[("PDF files", "*.pdf")])
        if file_path:
            doc = SimpleDocTemplate(file_path, pagesize=letter)
            styles = getSampleStyleSheet()
            elements = []

            # Add title
            elements.append(Paragraph("Network Performance Report", styles['Title']))
            elements.append(Spacer(1, 12))

            # Add latest test results
            if self.history:
                latest_result = self.history[-1]
                elements.append(Paragraph(f"Latest Test Results ({latest_result['timestamp']})", styles['Heading2']))
                elements.append(Spacer(1, 6))

                # Ping results
                elements.append(Paragraph("Ping Test Results:", styles['Heading3']))
                elements.append(Paragraph(f"Average Ping Time: {latest_result['ping']['avg_time']} ms", styles['Normal']))
                elements.append(Paragraph(self.interpret_ping(latest_result['ping']), styles['Normal']))
                elements.append(Spacer(1, 6))

                # DNS results
                elements.append(Paragraph("DNS Lookup Results:", styles['Heading3']))
                elements.append(Paragraph(latest_result['dns']['output'], styles['Normal']))
                elements.append(Paragraph(self.interpret_dns(latest_result['dns']), styles['Normal']))
                elements.append(Spacer(1, 6))

                # Speed test results
                elements.append(Paragraph("Speed Test Results:", styles['Heading3']))
                elements.append(Paragraph(f"Download Speed: {latest_result['speed']['download']:.2f} Mbps", styles['Normal']))
                elements.append(Paragraph(f"Upload Speed: {latest_result['speed']['upload']:.2f} Mbps", styles['Normal']))
                elements.append(Paragraph(f"Ping: {latest_result['speed']['ping']:.2f} ms", styles['Normal']))
                elements.append(Paragraph(self.interpret_speed(latest_result['speed']), styles['Normal']))
                elements.append(Spacer(1, 12))

            # Add history table
            elements.append(Paragraph("Test History", styles['Heading2']))
            elements.append(Spacer(1, 6))

            data = [['Timestamp', 'Ping (ms)', 'DNS Lookup', 'Download (Mbps)', 'Upload (Mbps)', 'Speedtest Ping (ms)']]
            for result in self.history:
                data.append([
                    result['timestamp'],
                    f"{result['ping']['avg_time']:.2f}" if result['ping']['avg_time'] is not None else 'N/A',
                    'Success' if result['dns']['success'] else 'Failure',
                    f"{result['speed']['download']:.2f}",
                    f"{result['speed']['upload']:.2f}",
                    f"{result['speed']['ping']:.2f}"
                ])

            table = Table(data)
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 8),
                ('TOPPADDING', (0, 1), (-1, -1), 6),
                ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ]))
            elements.append(table)

            # Build the PDF
            doc.build(elements)
            tk.messagebox.showinfo("Save Successful", f"PDF report saved to {file_path}")

    def setup_menu(self):
        menubar = tk.Menu(self.master)
        self.master.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Export to Excel", command=self.export_to_excel)
        file_menu.add_command(label="Export to JSON", command=self.export_to_json)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.master.quit)

        settings_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Settings", menu=settings_menu)
        settings_menu.add_command(label="Save Settings", command=self.save_settings)

    def load_settings(self):
        config = configparser.ConfigParser()
        try:
            config.read('settings.ini')
            interval = config.getint('Settings', 'interval', fallback=60)
            self.interval_entry.delete(0, tk.END)
            self.interval_entry.insert(0, str(interval))
        except:
            pass

    def save_settings(self):
        config = configparser.ConfigParser()
        config['Settings'] = {'interval': self.interval_entry.get()}
        with open('settings.ini', 'w') as configfile:
            config.write(configfile)
        tk.messagebox.showinfo("Settings Saved", "Your settings have been saved.")

    def export_to_excel(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(['Timestamp', 'Ping (ms)', 'DNS Lookup', 'Download Speed (Mbps)', 'Upload Speed (Mbps)', 'Speedtest Ping (ms)'])
            for result in self.history:
                ws.append([
                    result['timestamp'],
                    result['ping']['avg_time'],
                    'Success' if result['dns']['success'] else 'Failure',
                    result['speed']['download'],
                    result['speed']['upload'],
                    result['speed']['ping']
                ])
            wb.save(file_path)
            tk.messagebox.showinfo("Export Successful", f"Data exported to {file_path}")

    def export_to_json(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".json",
                                                 filetypes=[("JSON files", "*.json")])
        if file_path:
            with open(file_path, 'w') as f:
                json.dump(self.history, f, indent=4)
            tk.messagebox.showinfo("Export Successful", f"Data exported to {file_path}")

    def perform_tests(self):
        self.update_progress(20)
        ping_result = self.ping(PING_HOST)
        self.update_progress(40)
        dns_result = self.dns_lookup('google.com')
        self.update_progress(60)
        speed_result = self.speed_test()
        self.update_progress(100)
        
        return {
            'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            'ping': ping_result,
            'dns': dns_result,
            'speed': speed_result
        }

    def check_for_notifications(self, results):
        # يمكنك تخصيص الشروط حسب احتياجاتك
        if results['speed']['download'] < 5:  # إذا كانت سرعة التنزيل أقل من 5 Mbps
            notification.notify(
                title='Low Download Speed',
                message=f"Download speed is {results['speed']['download']:.2f} Mbps",
                app_name='Network Performance Tool'
            )
        if results['speed']['upload'] < 1:  # إذا كانت سرعة الرفع أقل من 1 Mbps
            notification.notify(
                title='Low Upload Speed',
                message=f"Upload speed is {results['speed']['upload']:.2f} Mbps",
                app_name='Network Performance Tool'
            )

if __name__ == "__main__":
    root = tk.Tk()
    app = NetworkPerformanceTool(root)
    root.mainloop()