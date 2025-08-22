import tkinter as tk
import sys, base64, aiohttp, asyncio, time, requests, threading, json
from tkinter import ttk, filedialog
from pathlib import Path

TS_LAST_ACCESS_TOKEN = None
ACCESS_TOKEN = None
REFRESH_TOKEN = None
TIMER_STARTED = False

APP_TITLE = "OneShare (by: 公众号“正月十九”)"

def format_bytes(size):
    """Format bytes into a human-readable string."""
    if size <= 1:
        return f"{size} byte"
    if size < 1024:
        return f"{size} bytes"
    for unit in ['bytes', 'KB', 'MB', 'GB', 'TB']:
        if size < 1024:
            return f"{size:.2f} {unit}"
        size /= 1024.0
    return f"{size:.2f} PB"

class AtomicInteger:
    def __init__(self, initial=0):
        self._value = initial
        self._lock = threading.Lock()

    def increment(self, amout = 1):
        with self._lock:
            self._value += amout
            print(f"已下载 {self._value} 字节 ({format_bytes(self._value)})        ", end='\r')
        return self._value
    
    def get(self):
        with self._lock:
            return self._value
        
    def reset(self):
        with self._lock:
            self._value = 0
        return self._value
    
bytes_downloaded = AtomicInteger()

def update_access_token_timer():
    while True:
        global TS_LAST_ACCESS_TOKEN, ACCESS_TOKEN, REFRESH_TOKEN
        if REFRESH_TOKEN:
            resp = requests.post('https://login.microsoftonline.com/common/oauth2/v2.0/token', data={
                'client_id': '2540f447-2f34-4999-b9b1-0bc79a09cd46',
                'grant_type': 'refresh_token',
                'refresh_token': REFRESH_TOKEN,
                'scope': 'Files.ReadWrite.All offline_access'
            })
            if resp.status_code == 200:
                root = resp.json()
                ACCESS_TOKEN = root['access_token']
                time.sleep(60 * 20)  # Refresh every 20 minutes
            else:
                time.sleep(5)

def get_access_token():
    global ACCESS_TOKEN
    return ACCESS_TOKEN

async def fetch_drive_item_next_link(next_link):
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {get_access_token()}"
    }
    async with aiohttp.ClientSession() as session:
        async with session.get(next_link, headers=headers, ssl=False) as response:
            if response.status == 200:
                data = await response.json()
                value = data.get('value', [])
                if data.get('@odata.nextLink', None):
                    value.extend(await fetch_drive_item_next_link(data['@odata.nextLink']))
                return value
            else:
                error_body = await response.text()
                print(f"HTTP Error ({next_link}): {response.status}\n{error_body}")
                return []

async def download_file_item(name, url, save_path):
    save_path = save_path.rstrip('/')
    async with aiohttp.ClientSession() as session:
        async with session.get(url, ssl=False) as response:
            if response.status == 200:
                #data = await response.read()
                with open(save_path + '/' + name, 'w+b') as f:
                    async for chunk in response.content.iter_any():
                        f.write(chunk)
                        bytes_downloaded.increment(len(chunk))
            else:
                error_body = await response.text()
                print(f"HTTP Error ({save_path}/{name}): {response.status}\n{error_body}")

async def download_driveitem(drive_id, item_id, save_path):  
    url = "https://graph.microsoft.com/v1.0/drives/{}/items/{}/children".format(drive_id, item_id)
    data = await fetch_drive_item_next_link(url)

    for item in data:
        if 'folder' in item or 'bundle' in item:
            item_save_path = save_path + '/' + item['name']
            Path.mkdir(Path(item_save_path), parents=True, exist_ok=True)

    file_list = [driveitem for driveitem in data if 'folder' not in driveitem and 'bundle' not in driveitem]
    file_tasks = [asyncio.create_task(download_file_item(f['name'], f['@microsoft.graph.downloadUrl'], save_path)) for f in file_list]
    folder_list = [driveitem for driveitem in data if 'folder' in driveitem or 'bundle' in driveitem]
    folder_tasks = [asyncio.create_task(download_driveitem(drive_id, f['id'], save_path + '/' + f['name'])) for f in folder_list]
    await asyncio.gather(*folder_tasks)
    await asyncio.gather(*file_tasks)
    return

async def download_1drv_ms(url, save_path):
    headers = {
        "Prefer": "redeemSharingLinkIfNecessary",
        "Authorization": f"Bearer {get_access_token()}",
        "Content-Type": "application/json"
    }

    url = "https://graph.microsoft.com/v1.0/shares/u!{}/driveItem".format(
        base64.urlsafe_b64encode(url).rstrip(b'=').decode())
    sys.stdout.write("\033[F")
    async with aiohttp.ClientSession() as session:
        async with session.get(url, ssl=False, headers=headers) as response:
            if response.status == 200:
                driveItem = await response.json()           
                if 'folder' in driveItem or 'bundle' in driveItem:
                    save_path = save_path + '/' + driveItem['name']
                    Path.mkdir(Path(save_path), parents=True, exist_ok=True)
                    await download_driveitem(driveItem['parentReference']['driveId'], driveItem['id'], save_path)
                else:
                    download_url = driveItem.get('@microsoft.graph.downloadUrl')
                    await download_file_item(driveItem['name'], download_url, save_path)
                print(f"下载完成, 共 {bytes_downloaded.get()} 字节        ")
            else:
                error_body = await response.text()
                print(f"HTTP Error (1drv_ms): {response.status}\n{error_body}")
    bytes_downloaded.reset()
    return

def DoOAuth2():
    import webbrowser
    from http.server import HTTPServer, SimpleHTTPRequestHandler
    class OAuth2Handler(SimpleHTTPRequestHandler):
        def do_GET(self):
            self.send_response(200)
            self.send_header('Content-type', 'text/html')
            self.end_headers()
            if self.path.startswith('/?'):
                query_params = self.path[2:].split('&')
                params = {param.split('=')[0]: param.split('=')[1] for param in query_params if '=' in param}
                if 'code' in params:
                    code = params['code']
                    resp = requests.post('https://login.microsoftonline.com/common/oauth2/v2.0/token', data={
                        'client_id': '2540f447-2f34-4999-b9b1-0bc79a09cd46',
                        'grant_type': 'authorization_code',
                        'code': code,
                        'redirect_uri': 'http://localhost:7856/',
                        'scope': 'Files.ReadWrite.All offline_access'
                    })
                    root = resp.json()
                    if not root.get('error'):
                        global ACCESS_TOKEN, REFRESH_TOKEN, TS_LAST_ACCESS_TOKEN
                        ACCESS_TOKEN = root['access_token']
                        REFRESH_TOKEN = root['refresh_token'] 
                        TS_LAST_ACCESS_TOKEN = time.time()
                        self.wfile.write('<!DOCTYPE html><html><head><meta charset="utf-8"/></head><body><div style="font-family:consolas;">Ok</div></body></html>'.encode('utf-8'))
                    else:
                        self.wfile.write(b'Error\n')

    server_address = ('', 7856)
    httpd = HTTPServer(server_address, OAuth2Handler)
    webbrowser.open('https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=2540f447-2f34-4999-b9b1-0bc79a09cd46&response_type=code&redirect_uri=http://localhost:7856/&scope=Files.ReadWrite.All offline_access&response_mode=query')
    httpd.serve_forever()
    httpd.server_close()

me_drive_parents = ['root']
def show_me_drive(tv):
    for item in tv.get_children():
        tv.delete(item)

    global me_drive_parents
    if len(me_drive_parents) >= 2:
        tv.insert('', 'end', iid=me_drive_parents[-2], values=('..（返回上一层）', '', '{}'))
    link = f'https://graph.microsoft.com/v1.0/me/drive/items/{me_drive_parents[-1]}/children?$select=id,name,size,file'
    driveitems = asyncio.run(fetch_drive_item_next_link(link))
    #resp = requests.get(f'https://graph.microsoft.com/v1.0/me/drive/items/{me_drive_parents[-1]}/children?$select=id,name,size,file', headers={
    #        "Authorization": f"Bearer {get_access_token()}",
    #        'Accept-Encoding':'gzip, deflate, br'
    #    })
    #root = resp.json()
    #for driveitem in root['value']:
    for driveitem in driveitems:
        name = f"{'📄' if 'file' in driveitem else '📁'} {driveitem['name']}"
        size = format_bytes(driveitem['size']) if 'file' in driveitem else ''
        resource = json.dumps(driveitem,indent=2)
        tv.insert('', 'end', iid=driveitem['id'], values=(name, size, resource))

def dbclick_me_drive(tv, event):    
    global me_drive_parents
    iid = tv.identify_row(event.y)
    if iid:
        item_details = tv.item(iid)
        root = json.loads(item_details.get('values', [None, None, '{}'])[2])
        if 'file' in root:
            pass
        else:
            if root.get('id'):
                me_drive_parents.append(root['id'])
            else:
                me_drive_parents.pop()
            show_me_drive(tv)
        return

async def download_me_driveitem_file(fileid_to_download, folder_path_var):
    url = f'https://graph.microsoft.com/v1.0/me/drive/items/{fileid_to_download}?select=id,name,@microsoft.graph.downloadUrl'
    resp = requests.get(url, headers={
            "Authorization": f"Bearer {get_access_token()}",
            'Accept-Encoding':'gzip, deflate, br'
        })
    root = resp.json()
    if '@microsoft.graph.downloadUrl' in root:
        await download_file_item(root['name'], root['@microsoft.graph.downloadUrl'], folder_path_var)

async def download_me_driveitem_folder(folderid_to_download, foldername_to_download, folder_path_var):
    url = f'https://graph.microsoft.com/v1.0/me/drive/items/{folderid_to_download}/children?$select=id,name,size,file'
    driveitems = await fetch_drive_item_next_link(url)
    files_to_download = [(item['id'],item['name']) for item in driveitems if 'file' in item]
    folders_to_download = [(item['id'],item['name']) for item in driveitems if 'file' not in item]
    folder_path_var.rstrip('/')
    folder_path_var = folder_path_var + '/' + foldername_to_download
    Path.mkdir(Path(folder_path_var), parents=True, exist_ok=True)
    await batch_download(files_to_download, folders_to_download, folder_path_var)
        
async def batch_download(files_to_download, folders_to_download, folder_path_var):
    file_tasks = [asyncio.create_task(download_me_driveitem_file(file_to_download[0], folder_path_var)) for file_to_download in files_to_download]
    folder_tasks = [asyncio.create_task(download_me_driveitem_folder(folder_to_download[0], folder_to_download[1], folder_path_var)) for folder_to_download in folders_to_download]
    await asyncio.gather(*file_tasks)
    await asyncio.gather(*folder_tasks)
    return

if __name__ == '__main__':
    oauth_thread = threading.Thread(target=DoOAuth2, daemon=True)
    oauth_thread.start()
    while ACCESS_TOKEN is None:
        time.sleep(1)

    token_update_thread = threading.Thread(target=update_access_token_timer, daemon=True)
    token_update_thread.start()

    #if len(sys.argv) > 0 and sys.argv[0] == 'me':
    if True:
        root = tk.Tk()
        root.title(APP_TITLE)

        # Create a main frame with padding
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        # Configure grid layout to be responsive
        main_frame.grid_columnconfigure(1, weight=1)
        
        folder_path_var = tk.StringVar()
        ttk.Label(main_frame, text="保存至:").grid(row=0, column=0, columnspan=1,sticky=tk.W, padx=5, pady=5)
        save_to_entry = ttk.Entry(main_frame, width=50, textvariable=folder_path_var)
        save_to_entry.grid(row=0, column=1, columnspan=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        def browse_folder(root):
            folder_selected = filedialog.askdirectory(initialdir=save_to_entry.get(), title="选择保存目录", parent=root)
            if folder_selected:
                folder_path_var.set(folder_selected)
        browse_button = ttk.Button(main_frame, text="浏览...", command=lambda:browse_folder(root))
        browse_button.grid(row=0, column=2, columnspan=1, sticky=tk.W, padx=5, pady=5)

        tv = ttk.Treeview(main_frame,columns=('name', 'size'), show='headings')
        tv.heading('name', text='Name', anchor='w')
        tv.heading('size', text='Size', anchor='w')
        tv.column('name', anchor='w',width=300,minwidth=300,stretch=True)
        tv.column('size', anchor='e',width=100,minwidth=100,stretch=False)
        tv.grid(row=1,column=0,columnspan=3,sticky=(tk.W,tk.E,tk.N,tk.S),padx=5,pady=6)
        tv.bind('<Double-1>', lambda event: dbclick_me_drive(tv,event))
        main_frame.grid_rowconfigure(1, weight=1)

        vsb = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=tv.yview)
        vsb.grid(row=1, column=3, sticky=(tk.N, tk.S))
        tv.configure(yscrollcommand=vsb.set)

        show_me_drive(tv)

        context_menu = tk.Menu(root, tearoff=0)
        def download_selected():
            global bytes_downloaded
            bytes_downloaded.reset()
            save_to_path = folder_path_var.get()
            if not save_to_path:
                tk.messagebox.showerror("错误", "请先选择保存目录")
                return
            
            files_to_download = []
            folders_to_download = []
            for iid in tv.selection():
                item_details = tv.item(iid)
                resource = json.loads(item_details.get('values', [None, None, '{}'])[2])
                if 'id' not in resource:
                    continue
                if 'file' in resource:
                    files_to_download.append((resource['id'], resource['name']))
                else:
                    folders_to_download.append((resource['id'], resource['name']))
                    
            asyncio.run(batch_download(files_to_download, folders_to_download, folder_path_var.get()))

        context_menu.add_command(label="下载所选项", command=download_selected)
        def show_context_menu(event):
            context_menu.post(event.x_root, event.y_root)
        if sys.platform == 'win32':
            tv.bind("<Button-3>", lambda event: show_context_menu(event))
        else:
            tv.bind("<Button-2>", lambda event: show_context_menu(event))

        root.minsize(500, 300)        
        root.lift()
        root.mainloop()
    else:
        root = tk.Tk()
        root.title(APP_TITLE)

        # Create a main frame with padding
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Configure grid layout to be responsive
        main_frame.grid_columnconfigure(1, weight=1)

        

        # --- Field 1: "保存至" (Save to) ---
        folder_path_var = tk.StringVar()
        ttk.Label(main_frame, text="保存至:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        save_to_entry = ttk.Entry(main_frame, width=50, textvariable=folder_path_var)
        save_to_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        def browse_folder(root):
            folder_selected = filedialog.askdirectory(initialdir=save_to_entry.get(), title="选择保存目录", parent=root)
            if folder_selected:
                folder_path_var.set(folder_selected)
        browse_button = ttk.Button(main_frame, text="浏览...", command=lambda:browse_folder(root))
        browse_button.grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)

        # --- Field 3: "链接" (Link) ---
        http_link = tk.StringVar()    

        def open_link():
            bytes_downloaded.reset()
            save_to_path = folder_path_var.get()
            if not save_to_path:
                tk.messagebox.showerror("错误", "请先选择保存目录")
                return
            if not http_link.get():
                tk.messagebox.showerror("错误", "请先输入链接")
                return
            save_to_path.rstrip('/')        
            print(f"正在下载 {http_link.get()} 至 \"{save_to_path}\" ...")
            print("请稍候...")
            asyncio.run(download_1drv_ms(http_link.get(), folder_path_var.get()))
            return

        ttk.Label(main_frame, text="链接:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(main_frame, textvariable=http_link).grid(row=2, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        download_button = ttk.Button(main_frame, text="下载链接", command=open_link)
        download_button.grid(row=2, column=2, sticky=tk.W, padx=5, pady=5)
        root.update_idletasks()  # Update the layout to get the button height
        download_button.update_idletasks()  # Ensure the button is fully rendered
        bottom = download_button.winfo_y() + download_button.winfo_height() + 10
        root.geometry(f"800x{bottom}")
        root.resizable(True, False)
        root.minsize(800, bottom)

        root.mainloop()
