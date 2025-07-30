import logging
import tkinter as tk
import queue

class QueueHandler(logging.Handler):
    def __init__(self, log_queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record):
        self.log_queue.put(self.format(record))

def configure_logging(log_queue):
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.DEBUG)
    
    file_handler = logging.FileHandler('dynamic_transaction_compiler.log')
    file_handler.setLevel(logging.DEBUG)
    file_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(file_formatter)
    root_logger.addHandler(file_handler)
    
    queue_handler = QueueHandler(log_queue)
    queue_handler.setLevel(logging.INFO)
    queue_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    queue_handler.setFormatter(queue_formatter)
    root_logger.addHandler(queue_handler)

def check_queues(log_queue, progress_queue, log_text, progress_bar, master):
    # Check log queue
    while not log_queue.empty():
        msg = log_queue.get()
        log_text.insert(tk.END, msg + '\n')
        log_text.see(tk.END)

    # Check progress queue
    while not progress_queue.empty():
        progress = progress_queue.get()
        progress_bar['value'] = progress

    master.after(100, check_queues, log_queue, progress_queue, log_text, progress_bar, master)

def update_progress(progress_queue, value):
    progress_queue.put(value)

def generate_fallback_id(row):
    return (
        str(row.get('Donor First Name', '')) +
        str(row.get('Donor Last Name', '')) +
        str(row.get('Donor Address Line 1', '')) +
        str(row.get('Donor City', '')) +
        str(row.get('Donor State', '')) +
        str(row.get('Donor ZIP', '')) +
        str(row.get('Donor Country', '')) +
        str(row.get('Donor Employer', ''))
    )

def add_unique_id(existing_ids, new_ids):
    unique_ids = set(existing_ids.split(' + ') + new_ids.split(' + '))
    return ' + '.join(unique_ids)