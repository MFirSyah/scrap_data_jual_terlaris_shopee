import time
import random
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from bs4 import BeautifulSoup
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

def parse_sales(sales_text):
    if not sales_text:
        return 0
    cleaned_text = sales_text.split(' ')[0].replace('+', '').strip()
    if 'RB' in cleaned_text.upper():
        cleaned_text = cleaned_text.upper().replace(',', '.').replace('RB', '')
        try:
            return int(float(cleaned_text) * 1000)
        except ValueError:
            return 0
    else:
        try:
            return int(cleaned_text.replace('.', ''))
        except ValueError:
            return 0

def load_all_sold_out_products(driver):
    print("\n🔁 Memuat semua produk yang sudah habis stok...")
    while True:
        try:
            see_more = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'div.shop-sold-out-see-more > button.shopee-button-outline'))
            )
            driver.execute_script("arguments[0].scrollIntoView();", see_more)
            time.sleep(random.uniform(1.5, 2.5))
            see_more.click()
            print("➡️ Klik 'Lihat Lainnya' untuk muat produk habis berikutnya...")
            time.sleep(random.uniform(2.5, 3.5))
        except TimeoutException:
            print("✅ Semua produk habis sudah dimuat.")
            break

def extract_products_from_soup(soup, existing_names, product_card_class, section_filter=None):
    if section_filter:
        product_items = section_filter.select(f'div.{product_card_class.replace(" ", ".")}')
    else:
        product_items = soup.select(f'div.{product_card_class.replace(" ", ".")}')
    
    result = []

    for item in product_items:
        name_element = item.select_one('div.line-clamp-2')
        name = name_element.get_text(strip=True) if name_element else "Nama Tidak Ditemukan"

        price_element = item.select_one('span.truncate')
        price_text = price_element.get_text(strip=True) if price_element else "0"
        try:
            price = int(price_text.replace('.', '').replace(',', ''))
        except ValueError:
            price = 0

        sales_element = item.find('div', string=lambda t: t and 'Terjual' in t)
        sales_text = sales_element.get_text(strip=True) if sales_element else ""
        sales = parse_sales(sales_text)

        if name not in existing_names:
            result.append({
                "Nama Produk": name,
                "Harga": price,
                "Terjual per Bulan": sales,
            })
            existing_names.add(name)

    return result

def scrape_shopee_products(url, preview_limit):
    print("🚀 Menghubungkan ke Chrome yang sudah terbuka...")
    options = Options()
    options.debugger_address = "127.0.0.1:9222"
    driver = webdriver.Chrome(service=Service(), options=options)

    print(f"🔗 Mengunjungi URL: {url}")
    driver.get(url)

    try:
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CLASS_NAME, 'shop-search-result-view__item'))
        )
        print("✅ Halaman berhasil dimuat.")
    except TimeoutException:
        print("❌ Gagal memuat halaman atau tidak ada produk ditemukan.")
        driver.quit()
        return

    produk_tersedia = []
    produk_habis = []
    seen_names_tersedia = set()
    seen_names_habis = set()
    page_count = 1

    # === Tahap 1: Scraping produk tersedia
    while True:
        print(f"\n📄 Scraping Produk Tersedia - Halaman {page_count}...")
        last_height = driver.execute_script("return document.body.scrollHeight")
        for _ in range(3):
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(random.uniform(2, 3))
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height

        soup = BeautifulSoup(driver.page_source, 'html.parser')
        new_data = extract_products_from_soup(soup, seen_names_tersedia, "shop-search-result-view__item")
        if not new_data:
            print("🚫 Tidak ada produk baru tersedia.")
            break

        start_idx = len(produk_tersedia) + 1
        produk_tersedia.extend(new_data)

        for i, d in enumerate(new_data, start=start_idx):
            if i <= preview_limit:
                print(f"  [Ready #{i}] {d['Nama Produk']} | Harga: Rp{d['Harga']:,} | Terjual: {d['Terjual per Bulan']}")
            elif i == preview_limit + 1:
                print(f"\n... (Menampilkan preview hingga {preview_limit} data produk tersedia, proses tetap lanjut) ...")

        try:
            next_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.shopee-icon-button--right'))
            )
            print("➡️ Klik Next...")
            next_button.click()
            page_count += 1

            delay = random.uniform(1, 10)
            print(f"⏳ Menunggu {delay:.2f} detik sebelum lanjut halaman berikutnya...")
            time.sleep(delay)

        except (TimeoutException, NoSuchElementException):
            print("🏁 Tidak ada tombol Next. Lanjut ke produk habis.")
            break

    # === Jeda sebelum scraping produk habis ===
    print("😴 Istirahat sebentar sebelum lanjut ke produk habis...")
    time.sleep(random.uniform(8, 12))

    # === Tahap 2: Scraping produk habis
    load_all_sold_out_products(driver)
    print("\n📦 Scraping Produk Habis...")
    soup = BeautifulSoup(driver.page_source, 'html.parser')

    all_sections = soup.select('div.shopee-header-section')
    new_data = []

    for section in all_sections:
        title = section.select_one('div.shopee-header-section__header__title')
        if title and 'Kamu Mungkin Suka' in title.text:
            continue  # Lewati bagian rekomendasi

        section_data = extract_products_from_soup(
            soup, seen_names_habis, "shop-collection-view__item", section_filter=section
        )
        new_data.extend(section_data)

        delay = random.uniform(1, 10)
        print(f"⏳ Menunggu {delay:.2f} detik sebelum lanjut ke section berikutnya...")
        time.sleep(delay)

    produk_habis.extend(new_data)

    for i, d in enumerate(new_data, start=1):
        if i <= preview_limit:
            print(f"  [Habis #{i}] {d['Nama Produk']} | Harga: Rp{d['Harga']:,} | Terjual: {d['Terjual per Bulan']}")
        elif i == preview_limit + 1:
            print(f"\n... (Menampilkan preview hingga {preview_limit} data produk habis, proses tetap lanjut) ...")

    driver.quit()

    # === Simpan dan Preview
    df_tersedia = pd.DataFrame(produk_tersedia)
    df_habis = pd.DataFrame(produk_habis)

    print(f"\n📊 Total Produk Tersedia: {len(df_tersedia)}")
    print(f"📊 Total Produk Habis   : {len(df_habis)}")

    print("\n📋 Preview Akhir Produk Tersedia:")
    print(df_tersedia.head(preview_limit).to_string(index=True) if not df_tersedia.empty else "Tidak ada data.")

    print("\n📋 Preview Akhir Produk Habis:")
    print(df_habis.head(preview_limit).to_string(index=True) if not df_habis.empty else "Tidak ada data.")

    print("\n💾 Menyimpan data ke Excel...")
    with pd.ExcelWriter("hasil_scraping_shopee.xlsx", engine='openpyxl') as writer:
        df_tersedia.to_excel(writer, sheet_name="Produk_Tersedia", index=False)
        df_habis.to_excel(writer, sheet_name="Produk_Habis", index=False)
    print("✅ Data berhasil disimpan ke 'hasil_scraping_shopee.xlsx'")

# === PROGRAM UTAMA ===
if __name__ == "__main__":
    url = input("🔗 Masukkan URL toko Shopee: ").strip()
    while not url.startswith("https://shopee.co.id"):
        url = input("❌ URL tidak valid. Masukkan ulang (harus diawali https://shopee.co.id): ").strip()

    try:
        limit = int(input("📊 Masukkan jumlah produk untuk preview (misal: 20 atau 1000): ").strip())
        if limit <= 0:
            raise ValueError
    except ValueError:
        print("⚠️ Input tidak valid. Default digunakan: 1000")
        limit = 1000

    scrape_shopee_products(url, limit)
