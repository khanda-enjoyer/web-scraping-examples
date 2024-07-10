from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium_stealth import stealth
from bs4 import BeautifulSoup as bs
import time
import re
import openpyxl

options = webdriver.ChromeOptions()
options.add_argument('start-maximixed')

options.add_experimental_option('excludeSwitches', ['enable-automation'])
options.add_experimental_option('useAutomationExtension', False)
driver = webdriver.Chrome(options=options)

stealth(driver,
        languages=['en-US', 'en'],
        vendor='Google Inc.',
        platform='Win32',
        webgl_vendor='Intel Inc.',
        renderer='Intel Iris OpenGL Engine',
        fix_hairline=True,
        )

# Создаем новую книгу Excel
wb = openpyxl.Workbook()
# Получаем активный лист
ws = wb.active

ws.append(['Название', 'Описание', 'Материал', 'Обработка края', 'Дополнительные характеристики', 'Высота ворса', 'Вес_г_м2', 'Особенность', 'Категория', 'Коллекция', 'Стиль', 'Цвет', 'Материал2', 'Форма', 'Высота ворса2', 'Метод производства', 'База', 'Краевая отделка', 'Фото']) # type: ignore

links = []
i = 0
for page in range (1, 2):
    url = f'https://www.ekohalionline.com/tum-urunler?sayfa={page}'
    driver.get(url)
    time.sleep(1)
    print(f'Page = {page}')
    mainbloack = driver.find_element(By.CLASS_NAME, 'ProductListContent')
    offers = mainbloack.find_elements(By.CSS_SELECTOR, '.detailLink.detailUrl')
    for item in offers:
        link = item.get_attribute('href')
        links.append(link)
    for item in offers:
        time.sleep(1)
        driver.get(links[i])
        i += 1 
        print(f'товар №{i}' )
        submainblock = driver.find_element(By.CLASS_NAME, 'ProductDetailMain')
        dirtyname = submainblock.find_element(By.CLASS_NAME, 'ProductName').text.strip()

        caps_words = re.findall(r'\b[A-Z]+\b', dirtyname)
        caps_words = [word for word in caps_words if len(word) > 1]
        cleanname = ' '.join(caps_words)
        print(cleanname)

        all_words = re.findall(r'\b\w+\b', dirtyname)
        non_caps_words = [word for word in all_words if not word.isupper()]
        description = ' '.join(non_caps_words)
        print(description)
        try:
            material_dirty = submainblock.find_element(By.XPATH, "//strong[contains(text(), 'Malzeme')]")
            not_material = material_dirty.text.strip()
            parent_element_material = material_dirty.find_element(By.XPATH, "./..")
            material = parent_element_material.text.replace(not_material, "").strip()
            print(material)
        except:
            material = ' '
            print('no material')
            
        try:
            edge_dirty = submainblock.find_element(By.XPATH, "//strong[contains(text(), 'Kenar Bitimi')]")
            not_edge = edge_dirty.text.strip()
            parent_element_edge = edge_dirty.find_element(By.XPATH, "./..")
            edge = parent_element_edge.text.replace(not_edge, "").strip()
            print(edge)
        except:
            edge = ' '
            print('no edge')
        
        try:
            char_dirty = submainblock.find_element(By.XPATH, "//strong[contains(text(), 'Ek Özellikler')]")
            not_char = char_dirty.text.strip()
            parent_element_char = char_dirty.find_element(By.XPATH, "./..")
            char = parent_element_char.text.replace(not_char, "").strip()
            print(char)
        except:
            char = ' '
            print('no char')

        try:
            height_dirty = submainblock.find_element(By.XPATH, "//strong[contains(text(), 'Hav Yüksekliği')]")
            not_height = height_dirty.text.strip()
            parent_element_height = height_dirty.find_element(By.XPATH, "./..")
            height = parent_element_height.text.replace(not_height, "").strip()
            print(height)
        except:
            height = ' '
            print('no height')
        
        try:
            weight_dirty = submainblock.find_element(By.XPATH, "//strong[contains(text(), 'M2 Ağırlığı (gr)')]")
            not_weight = weight_dirty.text.strip()
            parent_element_weight = weight_dirty.find_element(By.XPATH, "./..")
            weight = parent_element_weight.text.replace(not_weight, "").strip()
            print(weight)
        except:
            weight = ' '
            print('no weight')

        # таблица    
        try:
            feature = submainblock.find_element(By.CLASS_NAME, "ozellik").find_element(By.CLASS_NAME, "t2").text
            print(feature)
        except:
            print('no feature')
            feature = ' '

        try:
            categories = submainblock.find_element(By.CLASS_NAME, "kategori").find_element(By.CLASS_NAME, "t2").text
            print(categories)
        except:
            print('no categories')
            feature = ' '

        try:
            collection = submainblock.find_element(By.CLASS_NAME, "koleksiyon").find_element(By.CLASS_NAME, "t2").text
            print(collection)
        except:
            print('no collection')
            feature = ' '

        try:
            style = submainblock.find_element(By.CLASS_NAME, "tarz").find_element(By.CLASS_NAME, "t2").text
            print(style)
        except:
            print('no style')
            feature = ' '
        
        try:
            color = submainblock.find_element(By.CLASS_NAME, "renk").find_element(By.CLASS_NAME, "t2").text
            print(color)
        except:
            print('no color')
            feature = ' '

        try:
            material2 = submainblock.find_element(By.CLASS_NAME, "malzeme").find_element(By.CLASS_NAME, "t2").text
            print(material2)
        except:
            print('no material2')
            feature = ' '

        try:
            shape = submainblock.find_element(By.CLASS_NAME, "sekil").find_element(By.CLASS_NAME, "t2").text
            print(shape)
        except:
            print('no shape')
            feature = ' '

        try:
            height2 = submainblock.find_element(By.CLASS_NAME, "hav-yuksekligi").find_element(By.CLASS_NAME, "t2").text
            print(height2)
        except:
            print('no height2')
            feature = ' '
        
        try:
            production_method = submainblock.find_element(By.CLASS_NAME, "uretim-sekli").find_element(By.CLASS_NAME, "t2").text
            print(production_method)
        except:
            print('no production_method')
            feature = ' '
 
        try:
            base = submainblock.find_element(By.CLASS_NAME, "taban").find_element(By.CLASS_NAME, "t2").text
            print(base)
        except:
            print('no base')
            feature = ' '

        try:
            edge2 = submainblock.find_element(By.CLASS_NAME, "kenar-bitimi").find_element(By.CLASS_NAME, "t2").text
            print(edge2)
        except:
            print('no edge2')
            feature = ' '

        # картинки
        images_nostr = []
        try:
            # Попробуйте найти основное изображение
            big_image_link = submainblock.find_element(By.CLASS_NAME, "Images").find_element(By.TAG_NAME, 'img').get_attribute('src')
            images_nostr.append(big_image_link)
        except:
            print("Основное изображение не найдено")

        try:
            # Попробуйте найти дополнительные изображения
            small_images = submainblock.find_elements(By.CLASS_NAME, 'AltImgCapSmallImg')
            for image in small_images:
                image_link = image.find_element(By.TAG_NAME, 'img').get_attribute('src')
                images_nostr.append(image_link)
        except:
            print("Дополнительные изображения не найдены")

        # Обработайте случай, когда фото отсутствует
        if images_nostr:
            images = ', '.join(images_nostr)
        else:
            images = 'Фото отсутствует'

        ws.append([cleanname, description, material, edge, char, height, weight, feature, categories, collection, style, color, material2, shape, height2, production_method, base, edge2, images]) # type: ignore
        print(f'Добавлено: {cleanname}')

wb.save('ecohali.xlsx') 
time.sleep(2)
driver.quit()