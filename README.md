# XML Invayzer

XML dosyalarını analiz eden ve içindeki URL'leri kontrol eden bir masaüstü uygulaması.

## Özellikler

- XML dosyalarını URL'den indirme veya yerel dosyadan yükleme
- XML içindeki tag'leri hiyerarşik olarak görüntüleme
- Seçilen tag'lerdeki URL'leri analiz etme
- URL'lerin durum kodlarını ve yanıt sürelerini kontrol etme
- Sonuçları Excel dosyasına aktarma
- Analiz işlemini duraklatma, devam ettirme ve durdurma
- Analiz süresini takip etme

## Kurulum

### Windows için Kurulum

1. [Python 3.7 veya daha yüksek bir sürümü](https://www.python.org/downloads/) indirin ve kurun.
2. Kurulum sırasında "Add Python to PATH" seçeneğini işaretleyin.
3. Komut istemini (Command Prompt) yönetici olarak açın.
4. Aşağıdaki komutları sırasıyla çalıştırın:

```bash
pip install --upgrade pip
pip install -r requirements.txt
python setup.py install
```

### Linux için Kurulum

1. Terminal'i açın.
2. Aşağıdaki komutları sırasıyla çalıştırın:

```bash
sudo apt-get update
sudo apt-get install python3 python3-pip
pip3 install --upgrade pip
pip3 install -r requirements.txt
python3 setup.py install
```

## Kullanım

Uygulamayı başlatmak için:

```bash
xml-invayzer
```

veya

```bash
python -m xml_analyzer
```

## Gereksinimler

- Python 3.7 veya daha yüksek
- tkinter (genellikle Python ile birlikte gelir)
- requests
- pandas
- openpyxl
- aiohttp

## Lisans

MIT Lisansı 