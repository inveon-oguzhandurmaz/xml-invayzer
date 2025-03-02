import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import requests
import xml.etree.ElementTree as ET
import pandas as pd
import threading
import urllib.parse
import os
from urllib.parse import urlparse
import time
from datetime import datetime

class XMLAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("XML Analiz Uygulaması")
        self.root.geometry("1000x700")
        self.root.configure(padx=20, pady=20)
        
        # Analiz kontrol değişkenleri
        self.analysis_paused = False
        self.analysis_stopped = False
        self.current_analysis_thread = None
        
        # Kaynak dosya adını tutmak için değişken ekle
        self.source_file_name = None
        
        # Ana container frame oluştur
        main_container = ttk.Frame(root)
        main_container.pack(fill=tk.BOTH, expand=True)
        
        self.xml_data = None
        self.tags_hierarchy = {}  # Tag hiyerarşisini tutacak sözlük
        self.results = []
        self.namespaces = {}  # XML namespace'leri tutmak için
        
        # Stil tanımlamaları
        self.style = ttk.Style()
        self.style.configure("TButton", font=('Helvetica', 10, 'bold'))
        self.style.configure("TLabel", font=('Helvetica', 11))
        
        # URL Giriş Kısmı
        url_frame = ttk.LabelFrame(main_container, text="XML URL Giriş", padding=(10, 5))
        url_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(url_frame, text="XML URL:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.url_entry = ttk.Entry(url_frame, width=60)
        self.url_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        self.download_btn = ttk.Button(url_frame, text="XML İndir", command=self.download_xml)
        self.download_btn.grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)
        
        # Veya Dosya Seçme
        file_frame = ttk.LabelFrame(main_container, text="Veya Yerel XML Seç", padding=(10, 5))
        file_frame.pack(fill=tk.X, pady=10)
        
        self.file_path_var = tk.StringVar()
        self.file_path_entry = ttk.Entry(file_frame, textvariable=self.file_path_var, width=60, state="readonly")
        self.file_path_entry.grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        
        self.browse_btn = ttk.Button(file_frame, text="Dosya Seç", command=self.browse_file)
        self.browse_btn.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        # İlerleme çubuğu
        self.progress_frame = ttk.LabelFrame(main_container, text="İşlem Durumu", padding=(10, 5))
        self.progress_frame.pack(fill=tk.X, pady=10)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.progress_frame, variable=self.progress_var, length=700, mode="determinate")
        self.progress_bar.pack(fill=tk.X, padx=5, pady=5)
        
        self.status_var = tk.StringVar()
        self.status_var.set("Hazır")
        self.status_label = ttk.Label(self.progress_frame, textvariable=self.status_var)
        self.status_label.pack(anchor=tk.W, padx=5, pady=5)
        
        # Üç adet tag seçme kısmı
        tag_frame = ttk.LabelFrame(main_container, text="XML Tag Seçimi", padding=(10, 5))
        tag_frame.pack(fill=tk.X, pady=10)
        
        # Base tag seçimi
        ttk.Label(tag_frame, text="Base Tag:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.base_tag_combobox = ttk.Combobox(tag_frame, width=20, state="disabled")
        self.base_tag_combobox.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        self.base_tag_combobox.bind("<<ComboboxSelected>>", self.on_base_tag_selected)
        
        # Parent tag seçimi
        ttk.Label(tag_frame, text="Parent Tag:").grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)
        self.parent_tag_combobox = ttk.Combobox(tag_frame, width=20, state="disabled")
        self.parent_tag_combobox.grid(row=0, column=3, sticky=tk.W, padx=5, pady=5)
        self.parent_tag_combobox.bind("<<ComboboxSelected>>", self.on_parent_tag_selected)
        
        # Child tag seçimi
        ttk.Label(tag_frame, text="Child Tag:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.child_tag_combobox = ttk.Combobox(tag_frame, width=20, state="disabled")
        self.child_tag_combobox.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        
        # Butonlar için frame
        button_frame = ttk.Frame(tag_frame)
        button_frame.grid(row=1, column=2, columnspan=2, sticky=tk.E, padx=5, pady=5)
        
        self.analyze_btn = ttk.Button(button_frame, text="Analiz Et", command=self.analyze_xml, state="disabled")
        self.analyze_btn.pack(side=tk.LEFT, padx=5)
        
        # Kontrol butonları
        self.pause_btn = ttk.Button(button_frame, text="Duraklat", command=self.pause_analysis, state="disabled")
        self.pause_btn.pack(side=tk.LEFT, padx=5)
        
        self.resume_btn = ttk.Button(button_frame, text="Devam Et", command=self.resume_analysis, state="disabled")
        self.resume_btn.pack(side=tk.LEFT, padx=5)
        
        self.stop_btn = ttk.Button(button_frame, text="Durdur", command=self.stop_analysis, state="disabled")
        self.stop_btn.pack(side=tk.LEFT, padx=5)
        
        self.reset_btn = ttk.Button(button_frame, text="Sıfırla", command=self.reset_analysis, state="disabled")
        self.reset_btn.pack(side=tk.LEFT, padx=5)
        
        self.export_btn = ttk.Button(button_frame, text="Excel'e Aktar", command=self.export_to_excel, state="disabled")
        self.export_btn.pack(side=tk.LEFT, padx=5)
        
        # Sonuçlar için tablo frame'i
        result_frame = ttk.LabelFrame(main_container, text="Analiz Sonuçları", padding=(10, 5))
        result_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Treeview için scrollbar
        scrollbar_y = ttk.Scrollbar(result_frame, orient=tk.VERTICAL)
        scrollbar_x = ttk.Scrollbar(result_frame, orient=tk.HORIZONTAL)
        
        # Tablo oluşturma
        self.result_tree = ttk.Treeview(
            result_frame, 
            columns=("url", "status"), 
            show="headings",
            yscrollcommand=scrollbar_y.set,
            xscrollcommand=scrollbar_x.set,
            selectmode="extended"  # Çoklu seçime izin ver
        )
        
        # Scrollbar'ları konumlandırma
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        scrollbar_y.config(command=self.result_tree.yview)
        scrollbar_x.config(command=self.result_tree.xview)
        
        # Tablo başlıkları
        self.result_tree.heading("url", text="Request URL")
        self.result_tree.heading("status", text="Response Status Code")
        
        # Sütun genişlikleri
        self.result_tree.column("url", width=500, stretch=True)
        self.result_tree.column("status", width=150, anchor=tk.CENTER)
        
        self.result_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Hücre seçimi için özel bağlama
        self.result_tree.bind("<Button-1>", self.on_tree_click)
        self.result_tree.bind("<Control-c>", self.copy_selected_cell)

    def browse_file(self):
        """Yerel XML dosyası seçme işlevi"""
        file_path = filedialog.askopenfilename(filetypes=[("XML Dosyaları", "*.xml")])
        if file_path:
            self.file_path_var.set(file_path)
            self.load_xml_from_file(file_path)
    
    def download_xml(self):
        """URL'den XML indirme işlevi"""
        url = self.url_entry.get().strip()
        if not url:
            messagebox.showerror("Hata", "Lütfen geçerli bir URL girin")
            return
        
        # URL'nin geçerli olup olmadığını kontrol et
        try:
            result = urlparse(url)
            if not all([result.scheme, result.netloc]):
                messagebox.showerror("Hata", "Geçerli bir URL giriniz")
                return
        except:
            messagebox.showerror("Hata", "URL ayrıştırma hatası")
            return
        
        # İndirme işlemini ayrı bir thread'de başlat
        self.status_var.set("XML indiriliyor...")
        self.progress_var.set(0)
        self.download_btn.config(state="disabled")
        self.browse_btn.config(state="disabled")
        
        threading.Thread(target=self._download_xml_thread, args=(url,), daemon=True).start()
    
    def _download_xml_thread(self, url):
        """XML indirme işlemi için thread"""
        try:
            # URL'den dosya adını çıkar
            self.source_file_name = os.path.splitext(os.path.basename(urlparse(url).path))[0]
            if not self.source_file_name:
                self.source_file_name = "downloaded_xml"
                
            # İndirme işlemi simülasyonu
            for i in range(101):
                time.sleep(0.02)  # İndirme simülasyonu
                self.progress_var.set(i)
                self.root.update_idletasks()
            
            # İndirme işlemi
            response = requests.get(url, stream=True)
            if response.status_code != 200:
                self.root.after(0, lambda: messagebox.showerror("Hata", f"XML indirilemedi. Durum Kodu: {response.status_code}"))
                self.status_var.set("XML indirme hatası")
                self.root.after(0, self._reset_ui)
                return
            
            # Geçici dosyaya kaydet ve yükle
            temp_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "temp.xml")
            with open(temp_file, 'wb') as f:
                f.write(response.content)
            
            self.root.after(0, lambda: self.load_xml_from_file(temp_file))
            
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Hata", f"XML indirme hatası: {str(e)}"))
            self.status_var.set("Hata: " + str(e))
            self.root.after(0, self._reset_ui)
    
    def _reset_ui(self):
        """UI bileşenlerini sıfırla"""
        self.download_btn.config(state="normal")
        self.browse_btn.config(state="normal")
    
    def load_xml_from_file(self, file_path):
        """XML dosyasını yükle ve tag hiyerarşisini çıkar"""
        try:
            # Kaynak dosya adını kaydet
            self.source_file_name = os.path.splitext(os.path.basename(file_path))[0]
            
            tree = ET.parse(file_path)
            self.xml_data = tree
            root = tree.getroot()
            
            # RSS elementini atla ve içindeki ilk elementi root olarak kabul et
            if root.tag.lower().endswith('rss'):
                if len(root) > 0:
                    root = root[0]  # İlk child'ı root olarak kullan
            
            # Namespace'leri kaydet
            self.namespaces = {}
            for key, value in root.items():
                if key.startswith("{") or key.startswith("xmlns:"):
                    ns_parts = key.split("}")
                    if len(ns_parts) > 1:
                        ns = ns_parts[0].strip("{")
                        prefix = key.split(":")[1] if ":" in key else ""
                        self.namespaces[prefix] = ns
            
            # Tag hiyerarşisini oluştur
            self.tags_hierarchy = {}
            self._build_tag_hierarchy(root)
            
            # Root tag'i base tag listesine ekle
            root_tag = self._get_tag_name(root.tag)
            self.base_tag_combobox['values'] = [root_tag]
            self.base_tag_combobox.current(0)
            self.base_tag_combobox.config(state="readonly")
            
            # İlk base tag için parent tag'leri yükle
            self.on_base_tag_selected(None)
            
            self.status_var.set(f"XML yüklendi.")
            self.progress_var.set(100)
            self.reset_btn.config(state="normal")
            
        except Exception as e:
            messagebox.showerror("Hata", f"XML yükleme hatası: {str(e)}")
            self.status_var.set("XML yükleme hatası: " + str(e))
        finally:
            self.download_btn.config(state="normal")
            self.browse_btn.config(state="normal")
    
    def _get_tag_name(self, tag):
        """Namespace ile tag ismini ayırma"""
        if "}" in tag:
            return tag.split("}")[1]
        return tag
    
    def _build_tag_hierarchy(self, element, path=[]):
        """XML ağacındaki tag hiyerarşisini oluştur"""
        tag = self._get_tag_name(element.tag)
        current_path = path + [tag]
        
        # Mevcut yolun son tag'i için alt tagları topla
        if len(current_path) > 0:
            parent_path = '.'.join(current_path[:-1]) if len(current_path) > 1 else ""
            if parent_path not in self.tags_hierarchy:
                self.tags_hierarchy[parent_path] = []
            
            if tag not in self.tags_hierarchy[parent_path]:
                self.tags_hierarchy[parent_path].append(tag)
        
        # Alt elementlere devam et
        for child in element:
            self._build_tag_hierarchy(child, current_path)
    
    def on_base_tag_selected(self, event):
        """Base tag seçildiğinde parent tag'leri göster"""
        selected_base = self.base_tag_combobox.get()
        
        if selected_base:
            # Seçili base tag'in alt tag'leri
            parent_tags = self.tags_hierarchy.get("", [])
            parent_tags = [tag for tag in parent_tags if tag != selected_base]
            
            # Base tag'in doğrudan altındaki tag'ler
            direct_children = self.tags_hierarchy.get(selected_base, [])
            parent_tags.extend(direct_children)
            
            # Tekrarlananları kaldır ve sırala
            parent_tags = sorted(list(set(parent_tags)))
            
            if parent_tags:
                self.parent_tag_combobox['values'] = parent_tags
                self.parent_tag_combobox.current(0)
                self.parent_tag_combobox.config(state="readonly")
                self.on_parent_tag_selected(None)
            else:
                self.parent_tag_combobox.set("")
                self.parent_tag_combobox.config(state="disabled")
                self.child_tag_combobox.set("")
                self.child_tag_combobox.config(state="disabled")
                self.analyze_btn.config(state="disabled")
        else:
            self.parent_tag_combobox.set("")
            self.parent_tag_combobox.config(state="disabled")
    
    def on_parent_tag_selected(self, event):
        """Parent tag seçildiğinde child tag'leri göster"""
        selected_base = self.base_tag_combobox.get()
        selected_parent = self.parent_tag_combobox.get()
        
        if selected_parent:
            # Parent tag'in altındaki tag'ler
            parent_path = selected_base if selected_parent in self.tags_hierarchy.get("", []) else selected_parent
            child_path = selected_parent if parent_path == selected_base else f"{selected_base}.{selected_parent}"
            
            child_tags = self.tags_hierarchy.get(parent_path, []) + self.tags_hierarchy.get(child_path, [])
            
            # Aynı tag'i kaldır ve sırala
            child_tags = [tag for tag in child_tags if tag != selected_parent]
            child_tags = sorted(list(set(child_tags)))
            
            if child_tags:
                self.child_tag_combobox['values'] = child_tags
                self.child_tag_combobox.current(0)
                self.child_tag_combobox.config(state="readonly")
                self.analyze_btn.config(state="normal")
            else:
                self.child_tag_combobox.set("")
                self.child_tag_combobox.config(state="disabled")
                # Child tag olmasa bile seçili tag'leri analiz edebilmek için analiz düğmesini aktif bırak
                self.analyze_btn.config(state="normal")
        else:
            self.child_tag_combobox.set("")
            self.child_tag_combobox.config(state="disabled")
            self.analyze_btn.config(state="disabled")
    
    def reset_analysis(self):
        """Analiz sonuçlarını ve seçimleri sıfırla"""
        # Tabloyu temizle
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        
        # Sonuçları temizle
        self.results = []
        
        # Dosya yolu alanını temizle
        self.file_path_var.set("")
        
        # Export butonunu devre dışı bırak
        self.export_btn.config(state="disabled")
        
        # Tag seçimlerini yeniden aktif et
        self.base_tag_combobox.config(state="readonly")
        if self.parent_tag_combobox.get():
            self.parent_tag_combobox.config(state="readonly")
        if self.child_tag_combobox.get():
            self.child_tag_combobox.config(state="readonly")
        
        # Analiz butonunu aktif et
        self.analyze_btn.config(state="normal")
        
        self.status_var.set("Analiz sıfırlandı.")
        self.progress_var.set(0)
    
    def analyze_xml(self):
        """Seçilen tag'leri analiz et ve içindeki URL'lere istek at"""
        if not self.xml_data:
            messagebox.showerror("Hata", "Önce bir XML dosyası yükleyin")
            return
        
        base_tag = self.base_tag_combobox.get()
        parent_tag = self.parent_tag_combobox.get()
        child_tag = self.child_tag_combobox.get()
        
        if not base_tag:
            messagebox.showerror("Hata", "Lütfen en az bir base tag seçin")
            return
        
        # Analiz durumunu sıfırla
        self.analysis_paused = False
        self.analysis_stopped = False
        
        # UI'ı hazırla
        self.analyze_btn.config(state="disabled")
        self.base_tag_combobox.config(state="disabled")
        self.parent_tag_combobox.config(state="disabled")
        self.child_tag_combobox.config(state="disabled")
        self.download_btn.config(state="disabled")
        self.browse_btn.config(state="disabled")
        
        # Kontrol butonlarını aktifleştir
        self.pause_btn.config(state="normal")
        self.stop_btn.config(state="normal")
        self.resume_btn.config(state="disabled")
        
        self.progress_var.set(0)
        self.status_var.set("Analiz başlatılıyor...")
        
        # Tabloyu temizle
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
            
        # Analiz işlemi için thread başlat
        self.current_analysis_thread = threading.Thread(
            target=self._analyze_thread, 
            args=(base_tag, parent_tag, child_tag), 
            daemon=True
        )
        self.current_analysis_thread.start()
    
    def _find_elements_by_path(self, base_tag, parent_tag=None, child_tag=None):
        """Belirtilen tag yoluna göre elementleri bul"""
        root = self.xml_data.getroot()
        root_tag = self._get_tag_name(root.tag)
        
        results = []
        
        # Base tag root ise doğrudan root'u kullan
        if base_tag == root_tag:
            base_elements = [root]
        else:
            # Tüm base tag'leri bul
            base_elements = []
            for elem in root.iter():
                if self._get_tag_name(elem.tag) == base_tag:
                    base_elements.append(elem)
        
        # Parent tag belirtilmemişse base elementleri dön
        if not parent_tag:
            return base_elements
        
        # Tüm base elementlerinin altındaki parent tag'leri bul
        parent_elements = []
        for base_elem in base_elements:
            for child in base_elem.iter():
                if self._get_tag_name(child.tag) == parent_tag:
                    parent_elements.append(child)
        
        # Child tag belirtilmemişse parent elementleri dön
        if not child_tag:
            return parent_elements
        
        # Tüm parent elementlerinin altındaki child tag'leri bul
        child_elements = []
        for parent_elem in parent_elements:
            for child in parent_elem.iter():
                if self._get_tag_name(child.tag) == child_tag:
                    child_elements.append(child)
        
        return child_elements
    
    def _analyze_thread(self, base_tag, parent_tag, child_tag):
        """XML analiz işlemi için thread"""
        try:
            # Browser benzeri headers
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
                'Accept-Language': 'tr,en-US;q=0.7,en;q=0.3',
                'Accept-Encoding': 'gzip, deflate, br',
                'Connection': 'keep-alive',
                'Upgrade-Insecure-Requests': '1',
                'Sec-Fetch-Dest': 'document',
                'Sec-Fetch-Mode': 'navigate',
                'Sec-Fetch-Site': 'none',
                'Sec-Fetch-User': '?1',
                'Cache-Control': 'max-age=0'
            }

            elements_to_analyze = []
            
            if child_tag:
                elements_to_analyze = self._find_elements_by_path(base_tag, parent_tag, child_tag)
            elif parent_tag:
                elements_to_analyze = self._find_elements_by_path(base_tag, parent_tag)
            else:
                elements_to_analyze = self._find_elements_by_path(base_tag)
            
            if not elements_to_analyze:
                tag_path = f"{base_tag}"
                if parent_tag:
                    tag_path += f" > {parent_tag}"
                if child_tag:
                    tag_path += f" > {child_tag}"
                self.root.after(0, lambda: messagebox.showinfo("Bilgi", f"'{tag_path}' yolunda element bulunamadı"))
                self.root.after(0, self._reset_analyze_ui)
                return
                
            total_elements = len(elements_to_analyze)
            self.status_var.set(f"{total_elements} adet element bulundu. URL'ler analiz ediliyor...")
            
            self.results = []
            processed_count = 0
            
            for element in elements_to_analyze:
                # Durdurma kontrolü
                if self.analysis_stopped:
                    self.root.after(0, lambda: self.status_var.set("Analiz durduruldu."))
                    self.root.after(0, self._reset_analyze_ui)
                    return
                
                # Duraklatma kontrolü
                while self.analysis_paused and not self.analysis_stopped:
                    time.sleep(0.1)
                    continue
                
                if self.analysis_stopped:
                    self.root.after(0, lambda: self.status_var.set("Analiz durduruldu."))
                    self.root.after(0, self._reset_analyze_ui)
                    return
                
                # Element içinden URL'leri çıkar
                urls = self._extract_urls(element)
                
                # Eğer element metni doğrudan bir URL ise
                if not urls and element.text and (element.text.strip().startswith('http://') or element.text.strip().startswith('https://')):
                    urls = [element.text.strip()]
                
                for url in urls:
                    try:
                        # HTTP isteği at
                        response = requests.head(url, headers=headers, timeout=10, allow_redirects=True)
                        
                        if response.status_code == 403:
                            response = requests.get(url, headers=headers, timeout=10, allow_redirects=True)
                            
                        status_code = response.status_code
                    except Exception as e:
                        status_code = f"Hata: {str(e)}"
                    
                    # Sonucu kaydet
                    self.results.append((url, status_code))
                    
                    # Tablo arayüzünü güncelle
                    self.root.after(0, lambda u=url, s=status_code: self.result_tree.insert("", tk.END, values=(u, s)))
                
                # İlerleme sayacını artır
                processed_count += 1
                
                # İlerleme çubuğunu güncelle
                progress = (processed_count / total_elements) * 100
                self.progress_var.set(progress)
                if not self.analysis_paused:
                    self.status_var.set(f"Analiz ediliyor: {processed_count}/{total_elements} ({int(progress)}%)")
                self.root.update_idletasks()
            
            # Analiz tamamlandı
            if not self.analysis_stopped:
                self.status_var.set(f"Analiz tamamlandı. {len(self.results)} URL kontrol edildi.")
                self.export_btn.config(state="normal")
            if self.analysis_paused:
                 self.status_var.set("Analiz duraklatıldı...");
            
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Hata", f"Analiz hatası: {str(e)}"))
            self.status_var.set("Analiz hatası: " + str(e))
        finally:
            self.root.after(0, self._reset_analyze_ui)
            self.current_analysis_thread = None
    
    def _extract_urls(self, element):
        """Bir element içinden URL'leri çıkarma"""
        urls = []
        
        # Element metnini kontrol et
        if element.text and ('http://' in element.text or 'https://' in element.text):
            # Basit URL çıkarma - gerçek durumda daha karmaşık olabilir
            text = element.text.strip()
            if text.startswith('http://') or text.startswith('https://'):
                urls.append(text)
        
        # Alt elementleri kontrol et
        for child in element:
            # Özyinelemeli olarak alt elementlerdeki URL'leri de kontrol et
            urls.extend(self._extract_urls(child))
            
        # Nitelikler (attributes) içindeki URL'leri kontrol et
        for attr_name, attr_value in element.attrib.items():
            if attr_value and ('http://' in attr_value or 'https://' in attr_value):
                if attr_value.startswith('http://') or attr_value.startswith('https://'):
                    urls.append(attr_value)
        
        return urls
    
    def _reset_analyze_ui(self):
        """Analiz UI'ını sıfırla"""
        self.analyze_btn.config(state="normal")
        self.base_tag_combobox.config(state="readonly")
        self.parent_tag_combobox.config(state="readonly")
        self.child_tag_combobox.config(state="readonly")
        self.download_btn.config(state="normal")
        self.browse_btn.config(state="normal")
        
        # Kontrol butonlarını devre dışı bırak
        self.pause_btn.config(state="disabled")
        self.resume_btn.config(state="disabled")
        self.stop_btn.config(state="disabled")
    
    def on_tree_click(self, event):
        """Treeview'da tıklama olayını yakala"""
        region = self.result_tree.identify("region", event.x, event.y)
        if region == "cell":
            column = self.result_tree.identify_column(event.x)
            column_index = int(column.replace("#", "")) - 1
            item = self.result_tree.identify_row(event.y)
            if item and column_index >= 0:
                # Tek bir hücreyi seç
                self.result_tree.selection_set(item)
                # Tıklanan sütunu hatırla
                self.selected_column = column_index
    
    def copy_selected_cell(self, event):
        """Seçili hücrenin içeriğini kopyala"""
        selection = self.result_tree.selection()
        if selection and hasattr(self, 'selected_column'):
            item = selection[0]
            values = self.result_tree.item(item, 'values')
            if values and len(values) > self.selected_column:
                # Seçili hücrenin değerini kopyala
                self.root.clipboard_clear()
                self.root.clipboard_append(str(values[self.selected_column]))
                self.status_var.set("Değer panoya kopyalandı")
                return "break"  # Varsayılan davranışı engelle
    
    def export_to_excel(self):
        """Sonuçları Excel olarak dışa aktar"""
        if not self.results:
            messagebox.showinfo("Bilgi", "Dışa aktarılacak sonuç yok")
            return
        
        # Varsayılan dosya adını oluştur
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"{self.source_file_name or 'xml'}_analyzed_{timestamp}.xlsx"
        
        # Dosya kaydetme iletişim kutusu
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Dosyaları", "*.xlsx"), ("Tüm Dosyalar", "*.*")],
            initialfile=default_filename
        )
        
        if not file_path:
            return
        
        try:
            # Sonuçları DataFrame'e dönüştür
            df = pd.DataFrame(self.results, columns=["Request Url", "Response Status Code"])
            
            # Excel olarak kaydet
            df.to_excel(file_path, index=False)
            
            self.status_var.set(f"Sonuçlar başarıyla dışa aktarıldı: {file_path}")
            messagebox.showinfo("Başarılı", f"Sonuçlar başarıyla dışa aktarıldı:\n{file_path}")
            
        except Exception as e:
            messagebox.showerror("Hata", f"Excel dışa aktarma hatası: {str(e)}")
            self.status_var.set("Excel dışa aktarma hatası: " + str(e))

    def pause_analysis(self):
        """Analizi duraklat"""
        self.analysis_paused = True
        self.pause_btn.config(state="disabled")
        self.resume_btn.config(state="normal")
        self.status_var.set("Analiz duraklatıldı...")
        self.export_btn.config(state="normal")

    def resume_analysis(self):
        """Analizi devam ettir"""
        self.analysis_paused = False
        self.pause_btn.config(state="normal")
        self.resume_btn.config(state="disabled")
        self.status_var.set("Analiz devam ediyor...")
        self.download_btn.config(state="disabled")

    def stop_analysis(self):
        """Analizi durdur"""
        self.analysis_stopped = True
        self.analysis_paused = False
        self.pause_btn.config(state="disabled")
        self.resume_btn.config(state="disabled")
        self.stop_btn.config(state="disabled")
        self.status_var.set("Analiz durduruluyor...")
        
        # Analiz thread'i bitmesini bekle
        if self.current_analysis_thread and self.current_analysis_thread.is_alive():
            self.current_analysis_thread.join(timeout=0.5)  # En fazla 0.5 saniye bekle
            self.status_var.set("Analiz durduruldu.")
            self._reset_analyze_ui()

        self.export_btn.config(state="normal")    


if __name__ == "__main__":
    root = tk.Tk()
    app = XMLAnalyzerApp(root)
    root.mainloop()