import wmi
import json
import os
import tkinter as tk
from tkinter import ttk
import webbrowser

def list_drivers(save_to_file=False):
    try:
        computer = wmi.WMI()
        drivers = []
        for device in computer.Win32_PnPEntity():
            if device.Name and device.DeviceID:
                status = "Yüklü" if device.Status == "OK" else "Yüklü Değil"
                driver_info = {
                    "Cihaz": device.Name,
                    "Device ID": device.DeviceID,
                    "Üretici": device.Manufacturer if hasattr(device, 'Manufacturer') else "Bilinmiyor",
                    "Durum": status,
                    "Sürücü Linki": "Sürücüyü araştır"  # Google arama linki
                }
                drivers.append(driver_info)
        
        if save_to_file:
            file_path = os.path.join(os.getcwd(), "drivers_info.json")
            with open(file_path, "w", encoding="utf-8") as file:
                json.dump(drivers, file, ensure_ascii=False, indent=4)
            print(f"Sürücü bilgileri {file_path} dosyasına kaydedildi.")
        
        return drivers
        
    except Exception as e:
        print(f"Hata: {str(e)}")
        return []

def display_drivers_in_table():
    drivers = list_drivers()

    # Tkinter GUI oluşturma
    root = tk.Tk()
    root.title("Bilgisayar Sürücüleri")
    root.geometry("1000x500")

    # Arama çubuğu oluşturma
    search_frame = tk.Frame(root)
    search_frame.pack(fill="x", padx=10, pady=5)

    search_label = tk.Label(search_frame, text="Ara:")
    search_label.pack(side="left", padx=5)

    search_entry = tk.Entry(search_frame)
    search_entry.pack(side="left", fill="x", expand=True, padx=5)

    # Treeview tablo oluşturma
    columns = ("Cihaz", "Device ID", "Üretici", "Durum", "Sürücü Linki")
    tree = ttk.Treeview(root, columns=columns, show="headings")

    # Sütun başlıkları ekleme
    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, width=200)

    # Sürücü bilgilerini tabloya ekleme
    def populate_treeview(data):
        for item in tree.get_children():
            tree.delete(item)
        for driver in data:
            item_id = tree.insert("", tk.END, values=(
                driver.get("Cihaz", "Bilinmiyor"),
                driver.get("Device ID", "Bilinmiyor"),
                driver.get("Üretici", "Bilinmiyor"),
                driver.get("Durum", "Bilinmiyor"),
                "Sürücüyü araştır" if driver.get("Sürücü Linki") != "Bilinmiyor" else "Bilinmiyor"
            ))
            # Durum rengine göre arka planı ayarla
            if driver.get("Durum") == "Yüklü":
                tree.item(item_id, tags=("yuklu",))
            else:
                tree.item(item_id, tags=("yuklu_degil",))

    # Tag stilleri ekleme
    tree.tag_configure("yuklu", background="lightgreen")
    tree.tag_configure("yuklu_degil", background="lightcoral")

    populate_treeview(drivers)
    tree.pack(expand=True, fill="both")

    # Arama işlevi
    def search_drivers(event=None):
        query = search_entry.get().lower()
        filtered_drivers = [
            driver for driver in drivers
            if (driver.get("Cihaz") and query in driver["Cihaz"].lower()) or
               (driver.get("Device ID") and query in driver["Device ID"].lower()) or
               (driver.get("Üretici") and query in driver["Üretici"].lower()) or
               (driver.get("Durum") and query in driver["Durum"].lower())
        ]
        populate_treeview(filtered_drivers)

    search_entry.bind("<KeyRelease>", search_drivers)

    search_button = tk.Button(search_frame, text="Ara", command=search_drivers)
    search_button.pack(side="left", padx=5)

    # Sürücü Linki hücresine çift tıklandığında Google'da arama yapma
    def on_treeview_double_click(event):
        item = tree.selection()[0]
        device_name = tree.item(item, "values")[0]
        driver_link = f"https://www.google.com/search?q={device_name} driver"
        if driver_link:
            webbrowser.open(driver_link)

    tree.bind("<Double-1>", on_treeview_double_click)

    root.mainloop()

if __name__ == "__main__":
    display_drivers_in_table()
