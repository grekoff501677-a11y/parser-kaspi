import sys
import time
import random
import os
import pandas as pd
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QLineEdit, QTextEdit, QProgressBar,
    QTableWidget, QTableWidgetItem, QHeaderView, QFileDialog, QComboBox,
    QTabWidget, QGroupBox, QCheckBox
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QUrl
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
from PyQt5.QtGui import QDesktopServices, QFont, QIcon


class ComparisonTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        main_layout = QVBoxLayout(self)
        file_selection_layout = QHBoxLayout()

        self.old_file_btn = QPushButton("Загрузить файл прошлого периода")
        self.current_file_btn = QPushButton("Загрузить файл текущего периода")
        self.compare_btn = QPushButton("Сравнить файлы")
        self.export_btn = QPushButton("Экспорт результатов")

        self.old_file_path = None
        self.current_file_path = None
        self.old_data = None
        self.current_data = None
        self.disappeared_items = None
        self.new_items = None

        self.old_file_label = QLabel("Файл прошлого периода: не выбран")
        self.current_file_label = QLabel("Файл текущего периода: не выбран")

        self.results_table = QTableWidget()
        self.results_table.setColumnCount(3)
        self.results_table.setHorizontalHeaderLabels(["URL", "Код товара", "Наименование"])
        self.results_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.results_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.results_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        # Разрешаем переход по ссылке при двойном клике
        self.results_table.cellDoubleClicked.connect(self.open_link)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)

        self.old_file_btn.clicked.connect(self.load_old_file)
        self.current_file_btn.clicked.connect(self.load_current_file)
        self.compare_btn.clicked.connect(self.compare_files)
        self.export_btn.clicked.connect(self.export_results)

        self.compare_btn.setEnabled(False)
        self.export_btn.setEnabled(False)

        file_selection_layout.addWidget(self.old_file_btn)
        file_selection_layout.addWidget(self.current_file_btn)
        file_selection_layout.addWidget(self.compare_btn)
        file_selection_layout.addWidget(self.export_btn)

        main_layout.addLayout(file_selection_layout)
        main_layout.addWidget(self.old_file_label)
        main_layout.addWidget(self.current_file_label)
        main_layout.addWidget(QLabel("Результаты сравнения:"))
        main_layout.addWidget(self.results_table)
        main_layout.addWidget(QLabel("Лог:"))
        main_layout.addWidget(self.log_text)

    def log(self, message):
        self.log_text.append(message)

    def load_old_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Выберите файл прошлого периода", "", "Excel files (*.xlsx *.xls)")
        if path:
            try:
                self.old_data = pd.read_excel(path)
                self.old_file_path = path
                self.old_file_label.setText(f"Файл прошлого периода: {os.path.basename(path)}")
                self.log(f"Загружен файл прошлого периода: {path}")
                cols = ["Код товара", "Наименование", "URL"]
                if not all(c in self.old_data.columns for c in cols):
                    self.log("ВНИМАНИЕ: Структура файла не соответствует ожидаемой")
                self._check_ready()
            except Exception as e:
                self.log(f"Ошибка при загрузке: {e}")

    def load_current_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Выберите файл текущего периода", "", "Excel files (*.xlsx *.xls)")
        if path:
            try:
                self.current_data = pd.read_excel(path)
                self.current_file_path = path
                self.current_file_label.setText(f"Файл текущего периода: {os.path.basename(path)}")
                self.log(f"Загружен файл текущего периода: {path}")
                cols = ["Код товара", "Наименование", "URL"]
                if not all(c in self.current_data.columns for c in cols):
                    self.log("ВНИМАНИЕ: Структура файла не соответствует ожидаемой")
                self._check_ready()
            except Exception as e:
                self.log(f"Ошибка при загрузке: {e}")

    def _check_ready(self):
        if self.old_data is not None and self.current_data is not None:
            self.compare_btn.setEnabled(True)

    def compare_files(self):
        if self.old_data is None or self.current_data is None:
            self.log("Сначала загрузите оба файла")
            return
        self.results_table.setRowCount(0)
        old = set(self.old_data['Код товара'])
        curr = set(self.current_data['Код товара'])
        gone = self.old_data[self.old_data['Код товара'].isin(old - curr)]
        new = self.current_data[self.current_data['Код товара'].isin(curr - old)]
        rows = len(gone) + len(new)
        self.results_table.setRowCount(rows)
        r = 0
        for _, it in gone.iterrows():
            for c, v in enumerate([it['URL'], it['Код товара'], it['Наименование']]):
                cell = QTableWidgetItem(str(v))
                if c == 0:
                    # Заменяем merchant id на Kama в URL
                    url = str(v).replace('m=30090572', 'm=Kama')
                    cell.setData(Qt.UserRole, url)
                cell.setBackground(Qt.red)
                self.results_table.setItem(r, c, cell)
            r += 1
        for _, it in new.iterrows():
            for c, v in enumerate([it['URL'], it['Код товара'], it['Наименование']]):
                cell = QTableWidgetItem(str(v))
                if c == 0:
                    # Заменяем merchant id на Kama в URL
                    url = str(v).replace('m=30090572', 'm=Kama')
                    cell.setData(Qt.UserRole, url)
                cell.setBackground(Qt.green)
                self.results_table.setItem(r, c, cell)
            r += 1
        self.log(f"Исчезнувшие: {len(gone)}, Новые: {len(new)}")
        self.disappeared_items, self.new_items = gone, new
        self.export_btn.setEnabled(True)

    def open_link(self, row, col):
        if col == 0:
            item = self.results_table.item(row, col)
            if item:
                url = item.data(Qt.UserRole)
                QDesktopServices.openUrl(QUrl(url))

    def export_results(self):
        if self.disappeared_items is None or self.new_items is None:
            self.log("Сначала выполните сравнение")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Сохранить результаты",
                                              f"comparison_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx",
                                              "Excel files (*.xlsx)")
        if path:
            try:
                with pd.ExcelWriter(path) as w:
                    # Приводим порядок столбцов к [URL, Код товара, Наименование]
                    def reorder(df):
                        cols = [c for c in ['URL', 'Код товара', 'Наименование'] if c in df.columns]
                        cols += [c for c in df.columns if c not in cols]
                        return df[cols]

                    reorder(self.disappeared_items).to_excel(w, sheet_name='Исчезнувшие', index=False)
                    reorder(self.new_items).to_excel(w, sheet_name='Новые', index=False)
                self.log(f"Сохранено: {path}")
            except Exception as e:
                self.log(f"Ошибка при экспорте: {e}")


class ParserThread(QThread):
    update_progress = pyqtSignal(int, int, int, int, int, str)
    update_log = pyqtSignal(str)
    finished_parsing = pyqtSignal(list)
    page_completed = pyqtSignal(int, int)
    update_total_items = pyqtSignal(int)

    def __init__(self, driver, auto_filter_options=None):
        super().__init__()
        self.driver = driver
        self.running = True
        self.products = []
        self.total_checked = 0
        self.total_with_pickup = 0
        self.processed_urls = set()  # Для отслеживания обработанных URL
        self.all_links = []  # Все ссылки на всех страницах
        self.expected_total_items = 0  # Ожидаемое общее количество товаров
        self.auto_filter_options = auto_filter_options or {}  # Опции автоматических фильтров

    def stop(self):
        self.running = False
        self.update_log.emit("Остановка парсера...")

    # Метод для применения автоматических фильтров
    def apply_filters(self):
        try:
            self.update_log.emit("Начинаю применение автоматических фильтров...")

            # 1. Применяем фильтр "Продавцы"
            if self.auto_filter_options.get('seller'):
                seller_name = self.auto_filter_options['seller']
                self.update_log.emit(f"Устанавливаю фильтр продавца: {seller_name}")

                # Находим блок фильтра продавцов
                seller_filter = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//span[contains(@class, 'filters__filter-title') and text()='Продавцы']/.."))
                )

                # Проверяем, нужно ли развернуть список
                try:
                    show_more = seller_filter.find_element(By.XPATH, ".//span[contains(@class, 'filters__spoiler')]")
                    if show_more.is_displayed():
                        self.driver.execute_script("arguments[0].click();", show_more)
                        time.sleep(0.5)
                except:
                    pass

                # Находим и выбираем нужного продавца
                try:
                    seller_option = seller_filter.find_element(
                        By.XPATH,
                        f".//span[contains(@class, 'filters__filter-row__description-label') and text()='{seller_name}']"
                    )
                    # Находим родительский элемент с чекбоксом и кликаем
                    seller_row = seller_option.find_element(By.XPATH,
                                                            "./ancestor::div[contains(@class, 'filters__filter-row')]")
                    if 'active' not in seller_row.get_attribute('class'):
                        checkbox = seller_row.find_element(By.XPATH,
                                                           ".//label[contains(@class, 'filters__filter-row__checkbox')]")
                        self.driver.execute_script("arguments[0].click();", checkbox)
                        time.sleep(1)  # Ждем применения фильтра
                        self.update_log.emit(f"Фильтр продавца установлен: {seller_name}")
                    else:
                        self.update_log.emit(f"Фильтр продавца уже активен: {seller_name}")
                except Exception as e:
                    self.update_log.emit(f"Ошибка при установке фильтра продавца: {str(e)}")

            # 2. Применяем фильтр "Тип" (литые)
            if self.auto_filter_options.get('type'):
                type_value = self.auto_filter_options['type']
                self.update_log.emit(f"Устанавливаю фильтр типа: {type_value}")

                try:
                    # Находим блок фильтра типа
                    type_filter = WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located(
                            (By.XPATH, "//span[contains(@class, 'filters__filter-title') and text()='Тип']/.."))
                    )

                    # Находим и выбираем нужный тип
                    type_option = type_filter.find_element(
                        By.XPATH,
                        f".//span[contains(@class, 'filters__filter-row__description-label') and text()='{type_value}']"
                    )
                    type_row = type_option.find_element(By.XPATH,
                                                        "./ancestor::div[contains(@class, 'filters__filter-row')]")
                    if 'active' not in type_row.get_attribute('class'):
                        checkbox = type_row.find_element(By.XPATH,
                                                         ".//label[contains(@class, 'filters__filter-row__checkbox')]")
                        self.driver.execute_script("arguments[0].click();", checkbox)
                        time.sleep(1)  # Ждем применения фильтра
                        self.update_log.emit(f"Фильтр типа установлен: {type_value}")
                    else:
                        self.update_log.emit(f"Фильтр типа уже активен: {type_value}")
                except Exception as e:
                    self.update_log.emit(f"Ошибка при установке фильтра типа: {str(e)}")

            # 3. Применяем фильтр "Комплектация" (4 диска)
            if self.auto_filter_options.get('equipment'):
                equipment_value = self.auto_filter_options['equipment']
                self.update_log.emit(f"Устанавливаю фильтр комплектации: {equipment_value}")

                try:
                    # Находим блок фильтра комплектации
                    equipment_filter = WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.XPATH,
                                                        "//span[contains(@class, 'filters__filter-title') and text()='Комплектация']/.."))
                    )

                    # Находим и выбираем нужную комплектацию
                    equipment_option = equipment_filter.find_element(
                        By.XPATH,
                        f".//span[contains(@class, 'filters__filter-row__description-label') and text()='{equipment_value}']"
                    )
                    equipment_row = equipment_option.find_element(By.XPATH,
                                                                  "./ancestor::div[contains(@class, 'filters__filter-row')]")
                    if 'active' not in equipment_row.get_attribute('class'):
                        checkbox = equipment_row.find_element(By.XPATH,
                                                              ".//label[contains(@class, 'filters__filter-row__checkbox')]")
                        self.driver.execute_script("arguments[0].click();", checkbox)
                        time.sleep(1)  # Ждем применения фильтра
                        self.update_log.emit(f"Фильтр комплектации установлен: {equipment_value}")
                    else:
                        self.update_log.emit(f"Фильтр комплектации уже активен: {equipment_value}")
                except Exception as e:
                    self.update_log.emit(f"Ошибка при установке фильтра комплектации: {str(e)}")


            # Ждем некоторое время после применения всех фильтров
            time.sleep(2)
            self.update_log.emit("Автоматические фильтры применены")
            return True

        except Exception as e:
            self.update_log.emit(f"Ошибка при применении фильтров: {str(e)}")
            return False

    def collect_all_links(self):
        """Предварительно собирает все ссылки со всех страниц"""
        page = 1
        all_links = []
        self.update_log.emit("Начинаю сбор всех ссылок на товары...")

        while self.running:
            try:
                # Ждем загрузки страницы
                WebDriverWait(self.driver, 15).until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, '.item-card__name-link'))
                )

                # Получаем ссылки на текущей странице
                links = [a.get_attribute('href') for a in
                         self.driver.find_elements(By.CSS_SELECTOR, '.item-card__name-link')]

                if not links:
                    self.update_log.emit(f"Страница {page}: нет товаров, завершаю сбор")
                    break

                # Добавляем новые ссылки (уникальные)
                new_links = [link for link in links if link not in all_links]
                all_links.extend(new_links)

                self.update_log.emit(f"Страница {page}: добавлено {len(new_links)} ссылок, всего: {len(all_links)}")

                # Проверяем наличие следующей страницы
                try:
                    # Проверка на наличие кнопки "Следующая" не в disabled состоянии
                    next_button = self.driver.find_element(By.XPATH,
                                                           "//li[contains(., 'Следующая') and not(contains(@class,'_disabled'))]")
                    # Сохраняем URL текущей страницы для проверки перехода
                    current_url = self.driver.current_url
                    # Кликаем по кнопке
                    self.driver.execute_script("arguments[0].click();", next_button)
                    # Ждем изменения URL
                    WebDriverWait(self.driver, 10).until(EC.url_changes(current_url))
                    # Увеличиваем счетчик страниц
                    page += 1
                    # Небольшая пауза для загрузки
                    time.sleep(0.5 + random.random())
                except Exception as e:
                    self.update_log.emit(f"Достигнута последняя страница или ошибка навигации: {str(e)}")
                    break

            except Exception as e:
                self.update_log.emit(f"Ошибка при сборе ссылок на странице {page}: {str(e)}")
                break

        # После сбора возвращаемся на первую страницу
        if all_links:
            try:
                self.driver.get(self.driver.current_url.split('?')[0])
                time.sleep(1)  # Даем время на загрузку
            except Exception as e:
                self.update_log.emit(f"Ошибка при возврате на первую страницу: {str(e)}")

        self.update_log.emit(f"Сбор ссылок завершен: всего {len(all_links)} товаров")
        return all_links

    def run(self):
        if not self.driver:
            self.update_log.emit("ОШИБКА: браузер не открыт")
            self.finished_parsing.emit([])
            return

        # Применяем автоматические фильтры, если они указаны
        if self.auto_filter_options:
            success = self.apply_filters()
            if not success:
                self.update_log.emit("Не удалось применить автоматические фильтры")
                # Можно продолжить без фильтров или прервать выполнение
                # self.finished_parsing.emit([])
                # return

        # Собираем все ссылки заранее
        self.all_links = self.collect_all_links()

        # Если не удалось собрать ссылки, завершаем работу
        if not self.all_links:
            self.update_log.emit("Не удалось собрать ссылки на товары")
            self.finished_parsing.emit([])
            return

        # Устанавливаем общее количество товаров по количеству собранных ссылок
        self.expected_total_items = len(self.all_links)
        self.update_total_items.emit(self.expected_total_items)
        self.update_log.emit(f"Всего товаров для обработки: {self.expected_total_items}")

        # Обрабатываем все ссылки
        total_links = len(self.all_links)
        self.update_log.emit(f"Начинаю обработку {total_links} товаров")

        # Обрабатываем по батчам
        batch_size = 6
        processed_count = 0

        for i in range(0, total_links, batch_size):
            if not self.running:
                self.update_log.emit("Остановка обработки товаров")
                break

            # Берем очередной батч ссылок
            batch = self.all_links[i:i + batch_size]

            # Открываем каждую ссылку в новой вкладке
            main_handle = self.driver.current_window_handle
            for url in batch:
                if url in self.processed_urls:
                    continue  # Пропускаем уже обработанные URL
                self.driver.execute_script(f"window.open('{url}','_blank')")
                time.sleep(0.3 + random.random() * 0.3)

            # Обрабатываем открытые вкладки
            for handle in self.driver.window_handles[1:]:
                if not self.running:
                    break

                self.driver.switch_to.window(handle)
                cur_url = self.driver.current_url

                # Проверяем, не обрабатывали ли мы уже этот URL
                if cur_url in self.processed_urls:
                    self.driver.close()
                    continue

                # Добавляем URL в обработанные
                self.processed_urls.add(cur_url)

                try:
                    # Ждем загрузки страницы
                    WebDriverWait(self.driver, 8).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'body')))

                    # Проверяем товар на наличие самовывоза
                    prod, msg = self.check_pickup()
                    if prod:
                        self.products.append({'Код товара': prod['code'], 'Наименование': prod['name'], 'URL': cur_url})
                        self.total_with_pickup += 1
                        if msg:
                            self.update_log.emit(msg)
                except Exception as e:
                    self.update_log.emit(f"[ERROR] {cur_url}: {str(e)}")
                finally:
                    # Увеличиваем счетчик проверенных товаров
                    processed_count += 1
                    self.total_checked += 1

                    # Обновляем прогресс
                    page_num = (i // 24) + 1  # Примерно 24 товара на странице
                    self.update_progress.emit(processed_count, total_links, page_num,
                                              self.total_with_pickup, self.total_with_pickup, "")

                    # Закрываем вкладку
                    self.driver.close()

            # Возвращаемся на главную вкладку
            self.driver.switch_to.window(main_handle)

            # Регулярно обновляем интерфейс с текущими результатами
            self.page_completed.emit(0, self.total_with_pickup)
            time.sleep(0.2)

        self.update_log.emit(
            f"Завершено. Всего проверено: {self.total_checked} из {self.expected_total_items}, с самовывозом: {self.total_with_pickup}")

        # Проверяем, все ли товары были обработаны
        if self.expected_total_items > 0 and self.total_checked < self.expected_total_items:
            missed = self.expected_total_items - self.total_checked
            self.update_log.emit(
                f"ВНИМАНИЕ: Не обработано {missed} товаров ({missed / self.expected_total_items * 100:.1f}%)")

        self.finished_parsing.emit(self.products)

        # Закрываем браузер, если скрипт не был остановлен вручную
        if self.running:
            self.driver.quit()

    def check_pickup(self):
        try:
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'td.sellers-table__cell')))
            soup = BeautifulSoup(self.driver.page_source, 'lxml')

            # Извлекаем информацию о товаре
            prod = {'name': 'Не найдено', 'code': 'Не найдено'}
            try:
                prod['name'] = soup.select_one('h1.item__heading').text.strip()
            except:
                pass
            try:
                prod['code'] = soup.select_one('div.item__sku').text.replace('Код товара:', '').strip()
            except:
                pass

            # Проверяем наличие самовывоза
            for row in soup.select("tr:has(.sellers-table__cell)"):
                for cell in row.select("td.sellers-table__cell"):
                    if 'Самовывоз' in cell.text:
                        return prod, f"[НАЙДЕН] {prod['name']} ({prod['code']})"

            # Если самовывоз не найден
            return None, None
        except Exception as e:
            return None, f"[ERROR Pickup]: {str(e)}"


class MerchantParserThread(ParserThread):
    """Поток парсинга только товаров нашего магазина (ID 30090572) с классификацией по городу.

    Отличия от базового ParserThread:
    - check_pickup переопределён под проверку наличия продавца с merchant id 30090572
    - Классификация города по наличию блока доставки «Доставка, сегодня, бесплатно»
    - В self.products добавляется поле 'Город'
    """

    OUR_MERCHANT_ID = "30090572"

    def _classify_city_by_delivery(self, seller_node):
        try:
            # Ищем блок доставки внутри узла продавца
            delivery = seller_node.select_one(
                "span.seller-item__delivery-option-link")
            if delivery:
                text = delivery.get_text(strip=True).lower()
                if ('доставка' in text) and ('сегодня' in text) and ('бесплатно' in text):
                    return 'Павлодар'
            # Если точного блока нет — считаем как Астана
            return 'Астана'
        except Exception:
            return 'Астана'

    def _find_our_seller_nodes(self, soup):
        candidates = []
        try:
            # 1) По data-seller-id
            by_attr = soup.select(f"[data-seller-id='{self.OUR_MERCHANT_ID}']")
            candidates.extend(by_attr)
        except Exception:
            pass

        try:
            # 2) По ссылке на merchant id
            links = soup.select("a[href*='merchant?id=']")
            for a in links:
                href = a.get('href', '')
                if self.OUR_MERCHANT_ID in href:
                    # Берём контейнер продавца (поднимаемся к строке продавца)
                    parent = a
                    for _ in range(5):
                        if parent and getattr(parent, 'get', None):
                            cls = parent.get('class', [])
                            if any('seller' in c for c in cls):
                                candidates.append(parent)
                                break
                        parent = getattr(parent, 'parent', None)
        except Exception:
            pass

        # Убираем дубликаты
        unique = []
        seen = set()
        for n in candidates:
            key = str(n)
            if key not in seen:
                seen.add(key)
                unique.append(n)
        return unique

    def check_pickup(self):  # Переопределяем: не фильтруем продавца, только классификация по доставке
        try:
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'td.sellers-table__cell')))
            soup = BeautifulSoup(self.driver.page_source, 'lxml')

            prod = {'name': 'Не найдено', 'code': 'Не найдено'}
            try:
                prod['name'] = soup.select_one('h1.item__heading').text.strip()
            except Exception:
                pass
            try:
                prod['code'] = soup.select_one('div.item__sku').text.replace('Код товара:', '').strip()
            except Exception:
                pass

            # Классифицируем город по наличию блока доставки на странице
            try:
                delivery = soup.select_one("span.seller-item__delivery-option-link")
                text = (delivery.get_text(strip=True).lower() if delivery else '')
                if ('доставка' in text) and ('сегодня' in text) and ('бесплатно' in text):
                    city = 'Павлодар'
                else:
                    city = 'Астана'
            except Exception:
                city = 'Астана'

            return prod | {'city': city}, f"[ТОВАР] {prod['name']} ({prod['code']}) — {city}"
        except Exception as e:
            return None, f"[ERROR Merchant]: {str(e)}"

    def run(self):
        if not self.driver:
            self.update_log.emit("ОШИБКА: браузер не открыт")
            self.finished_parsing.emit([])
            return

        # Собираем все ссылки заранее (наследуемся от базовой реализации)
        self.all_links = self.collect_all_links()

        if not self.all_links:
            self.update_log.emit("Не удалось собрать ссылки на товары")
            self.finished_parsing.emit([])
            return

        self.expected_total_items = len(self.all_links)
        self.update_total_items.emit(self.expected_total_items)
        self.update_log.emit(f"Всего товаров для обработки: {self.expected_total_items}")

        total_links = len(self.all_links)
        self.update_log.emit(f"Начинаю обработку {total_links} товаров (наш магазин)")

        batch_size = 6
        processed_count = 0

        for i in range(0, total_links, batch_size):
            if not self.running:
                self.update_log.emit("Остановка обработки товаров")
                break

            batch = self.all_links[i:i + batch_size]

            main_handle = self.driver.current_window_handle
            for url in batch:
                if url in self.processed_urls:
                    continue
                self.driver.execute_script(f"window.open('{url}','_blank')")
                time.sleep(0.3 + random.random() * 0.3)

            for handle in self.driver.window_handles[1:]:
                if not self.running:
                    break

                self.driver.switch_to.window(handle)
                cur_url = self.driver.current_url

                if cur_url in self.processed_urls:
                    self.driver.close()
                    continue

                self.processed_urls.add(cur_url)

                try:
                    WebDriverWait(self.driver, 8).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'body')))

                    prod, msg = self.check_pickup()
                    if prod:
                        # prod содержит name, code, city
                        self.products.append({
                            'code': prod.get('code', ''),
                            'name': prod.get('name', ''),
                            'url': cur_url,
                            'city': prod.get('city', ''),
                        })
                        self.total_with_pickup += 1
                        if msg:
                            self.update_log.emit(msg)
                except Exception as e:
                    self.update_log.emit(f"[ERROR] {cur_url}: {str(e)}")
                finally:
                    processed_count += 1
                    self.total_checked += 1
                    page_num = (i // 24) + 1
                    self.update_progress.emit(processed_count, total_links, page_num,
                                              self.total_with_pickup, self.total_with_pickup, "")
                    self.driver.close()

            self.driver.switch_to.window(main_handle)
            self.page_completed.emit(0, self.total_with_pickup)
            time.sleep(0.2)

        self.update_log.emit(
            f"Завершено. Всего проверено: {self.total_checked} из {self.expected_total_items}, найдено наших: {self.total_with_pickup}")

        if self.expected_total_items > 0 and self.total_checked < self.expected_total_items:
            missed = self.expected_total_items - self.total_checked
            self.update_log.emit(
                f"ВНИМАНИЕ: Не обработано {missed} товаров ({missed / self.expected_total_items * 100:.1f}%)")

        self.finished_parsing.emit(self.products)

        if self.running:
            self.driver.quit()


class OurShopTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.driver = None
        self.parser_thread = None
        self.products = []  # элементы с ключами: Код товара, Наименование, URL, Город

        main_layout = QVBoxLayout(self)

        # Верхняя панель управления
        ctl = QHBoxLayout()
        self.url_combo = QComboBox()
        self.url_combo.setEditable(True)
        self.url_combo.setMinimumWidth(400)
        self.url_combo.addItems([
            "https://kaspi.kz/shop/c/rims/",
            "https://kaspi.kz/shop/c/wheels/",
            "https://kaspi.kz/shop/c/auto%20parts/",
        ])
        ctl.addWidget(QLabel("URL категории:"))
        ctl.addWidget(self.url_combo)

        self.open_btn = QPushButton('Открыть фильтры')
        self.open_btn.clicked.connect(self.open_filters)
        ctl.addWidget(self.open_btn)

        self.start_btn = QPushButton('Парсинг')
        self.start_btn.setEnabled(False)
        self.start_btn.clicked.connect(self.start_parsing)
        ctl.addWidget(self.start_btn)

        self.stop_btn = QPushButton('Стоп')
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self.stop_parsing)
        ctl.addWidget(self.stop_btn)

        self.export_btn = QPushButton('Экспорт (2 файла)')
        self.export_btn.setEnabled(False)
        self.export_btn.clicked.connect(self.export_two_excel)
        ctl.addWidget(self.export_btn)

        main_layout.addLayout(ctl)

        # Прогресс
        pctl = QHBoxLayout()
        self.progress_bar = QProgressBar()
        self.progress_label = QLabel('Готов к работе')
        pctl.addWidget(self.progress_bar)
        pctl.addWidget(self.progress_label)
        main_layout.addLayout(pctl)

        # Таблица
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(['URL', 'Код товара', 'Наименование', 'Город'])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.table.cellDoubleClicked.connect(self.open_product_url)
        main_layout.addWidget(QLabel('Наши товары:'))
        main_layout.addWidget(self.table)

        # Лог
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont('Monospace', 9))
        main_layout.addWidget(QLabel('Лог:'))
        main_layout.addWidget(self.log_text)

    def log(self, msg):
        self.log_text.append(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")
        self.log_text.verticalScrollBar().setValue(self.log_text.verticalScrollBar().maximum())

    def open_product_url(self, r, c):
        url_item = self.table.item(r, 0)
        if url_item:
            url = url_item.data(Qt.UserRole)
            if url:
                QDesktopServices.openUrl(QUrl(url))

    def update_table(self, prods):
        self.table.setRowCount(len(prods))
        for r, p in enumerate(prods):
            it = QTableWidgetItem(p.get('URL', ''))
            it.setData(Qt.UserRole, p.get('URL', ''))
            self.table.setItem(r, 0, it)
            self.table.setItem(r, 1, QTableWidgetItem(p.get('Код товара', '')))
            self.table.setItem(r, 2, QTableWidgetItem(p.get('Наименование', '')))
            self.table.setItem(r, 3, QTableWidgetItem(p.get('Город', '')))

    def update_progress_bar(self, curr, total, page, found, tot, msg):
        if total > 0:
            p = int(curr / total * 100)
            self.progress_bar.setValue(p)
            self.progress_label.setText(f"Прогресс: {curr}/{total} ({p}%)")
        if msg:
            self.log(msg)

    def open_filters(self):
        url = self.url_combo.currentText()
        if not url.startswith('https://kaspi.kz'):
            self.log('Неверный URL')
            return

        self.log(f"Открытие: {url} (Павлодар)")
        self.products = []
        self.table.setRowCount(0)

        opts = webdriver.ChromeOptions()
        opts.add_argument('--disable-blink-features=AutomationControlled')
        opts.add_argument('--start-maximized')
        opts.add_argument('--disable-images')
        opts.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.110 Safari/537.36')

        self.driver = None
        if os.path.exists('chromedriver.exe'):
            try:
                self.log("Попытка использовать локальный chromedriver.exe...")
                self.driver = webdriver.Chrome(service=Service('chromedriver.exe'), options=opts)
                self.log("✅ Браузер запущен через локальный chromedriver.exe")
            except Exception as e:
                self.log(f"❌ Ошибка локального chromedriver: {e}")

        if not self.driver:
            try:
                self.log("Попытка использовать системный ChromeDriver...")
                self.driver = webdriver.Chrome(options=opts)
                self.log("✅ Браузер запущен через системный ChromeDriver")
            except Exception as e:
                self.log(f"❌ Ошибка системного ChromeDriver: {e}")

        if not self.driver:
            try:
                self.log("Попытка использовать webdriver-manager (требуется VPN)...")
                self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)
                self.log("✅ Браузер запущен через webdriver-manager")
            except Exception as e:
                self.log(f"❌ Ошибка webdriver-manager: {e}")
                self.log("💡 Включите VPN или скачайте chromedriver.exe вручную")
                return

        if not self.driver:
            self.log("❌ Не удалось запустить браузер.")
            return

        # Всегда устанавливаем Павлодар
        self.driver.get(url)
        try:
            self.driver.add_cookie({'name': 'kaspi.storefront.cookie.city', 'value': '551010000', 'domain': 'kaspi.kz', 'path': '/'})
            self.driver.refresh()
            self.log("Город установлен: Павлодар")
        except Exception as e:
            self.log(f"Ошибка установки города: {e}")

        self.open_btn.setEnabled(False)
        self.start_btn.setEnabled(True)

    def start_parsing(self):
        if not self.driver:
            self.log('ОШИБКА: откройте браузер')
            return

        self.log('Старт парсинга (наш магазин)')
        self.parser_thread = MerchantParserThread(self.driver, {})
        self.parser_thread.update_log.connect(self.log)
        self.parser_thread.update_progress.connect(self.update_progress_bar)
        self.parser_thread.finished_parsing.connect(self.parsing_finished)

        self.parser_thread.start()
        self.driver = None  # управление драйвером в потоке
        self.start_btn.setEnabled(False)
        self.open_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)

    def stop_parsing(self):
        if self.parser_thread:
            self.parser_thread.stop()
            self.start_btn.setEnabled(True)
            self.stop_btn.setEnabled(False)

    def parsing_finished(self, products):
        # Преобразуем к одному формату для таблицы
        self.products = []
        for pr in products:
            self.products.append({
                'Код товара': pr.get('code', ''),
                'Наименование': pr.get('name', ''),
                'URL': pr.get('url', pr.get('URL', '')) or '',
                'Город': pr.get('city', ''),
            })

        self.update_table(self.products)
        self.open_btn.setEnabled(True)
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(False)
        self.export_btn.setEnabled(bool(self.products))
        self.log(f"Парсинг завершен: найдено {len(self.products)} наших товаров")

    def export_two_excel(self):
        if not self.products:
            self.log('Нет данных для экспорта')
            return

        import pandas as pd
        # Разделяем по городу
        pvl = [p for p in self.products if p.get('Город') == 'Павлодар']
        ast = [p for p in self.products if p.get('Город') == 'Астана']

        # Сначала диалог для Павлодара
        ts = datetime.now().strftime('%Y%m%d')
        default_pvl = f"products_PVL_{ts}.xlsx"
        path_pvl, _ = QFileDialog.getSaveFileName(self, 'Сохранить (Павлодар)', default_pvl, 'Excel files (*.xlsx)')
        if path_pvl:
            try:
                df = pd.DataFrame(pvl)
                # Упорядочиваем столбцы
                cols = [c for c in ['URL', 'Код товара', 'Наименование', 'Город'] if c in df.columns]
                cols += [c for c in df.columns if c not in cols]
                df = df[cols]
                with pd.ExcelWriter(path_pvl) as writer:
                    df.to_excel(writer, sheet_name='Товары', index=False)
                    info = {
                        'Параметр': ['Дата парсинга', 'Город фиксации куки', 'URL категории', 'Применённые фильтры', 'Всего товаров', 'Найдено (Павлодар)'],
                        'Значение': [
                            datetime.now().strftime('%Y-%m-%d %H:%M'),
                            'Павлодар (cookie=551010000)',
                            self.url_combo.currentText(),
                            'Выбраны вручную в UI (вкл. продавца)',
                            len(self.products),
                            len(pvl)
                        ]
                    }
                    pd.DataFrame(info).to_excel(writer, sheet_name='Информация', index=False)
                self.log(f"Экспортировано (Павлодар): {path_pvl}")
            except Exception as e:
                self.log(f"Ошибка сохранения (Павлодар): {e}")

        # Затем диалог для Астаны
        default_ast = f"products_AST_{ts}.xlsx"
        path_ast, _ = QFileDialog.getSaveFileName(self, 'Сохранить (Астана)', default_ast, 'Excel files (*.xlsx)')
        if path_ast:
            try:
                df = pd.DataFrame(ast)
                # Упорядочиваем столбцы
                cols = [c for c in ['URL', 'Код товара', 'Наименование', 'Город'] if c in df.columns]
                cols += [c for c in df.columns if c not in cols]
                df = df[cols]
                with pd.ExcelWriter(path_ast) as writer:
                    df.to_excel(writer, sheet_name='Товары', index=False)
                    info = {
                        'Параметр': ['Дата парсинга', 'Город фиксации куки', 'URL категории', 'Применённые фильтры', 'Всего товаров', 'Найдено (Астана)'],
                        'Значение': [
                            datetime.now().strftime('%Y-%m-%d %H:%M'),
                            'Павлодар (cookie=551010000)',
                            self.url_combo.currentText(),
                            'Выбраны вручную в UI (вкл. продавца)',
                            len(self.products),
                            len(ast)
                        ]
                    }
                    pd.DataFrame(info).to_excel(writer, sheet_name='Информация', index=False)
                self.log(f"Экспортировано (Астана): {path_ast}")
            except Exception as e:
                self.log(f"Ошибка сохранения (Астана): {e}")

class KaspiParserGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.parser_thread = None
        self.products = []
        self.city_short = {'Астана': 'AST', 'Павлодар': 'PVL'}
        self.pvl_file_path = None
        self.ast_file_path = None
        self.pvl_data = None
        self.ast_data = None
        self.init_ui()
        self.expected_total_items = 0

    def init_ui(self):
        self.setWindowTitle('Kaspi Parser')
        self.setWindowIcon(QIcon('icon.ico'))
        self.setMinimumSize(900, 700)
        tabs = QTabWidget()
        self.setCentralWidget(tabs)

        # --- Parser Tab ---
        w1 = QWidget()
        l1 = QVBoxLayout(w1)

        # Основные элементы управления
        ctl = QHBoxLayout()
        self.url_combo = QComboBox()
        self.url_combo.setEditable(True)
        self.url_combo.setMinimumWidth(400)
        self.url_combo.addItems(["https://kaspi.kz/shop/c/rims/", "https://kaspi.kz/shop/c/wheels/",
                                 "https://kaspi.kz/shop/c/auto%20parts/"])
        ctl.addWidget(QLabel("URL категории:"))
        ctl.addWidget(self.url_combo)

        ctl.addWidget(QLabel("Город:"))
        self.city_combo = QComboBox()
        self.city_combo.addItem('Астана', '710000000')
        self.city_combo.addItem('Павлодар', '551010000')
        ctl.addWidget(self.city_combo)

        # Добавляем группу фильтров
        filter_group = QGroupBox("Автоматические фильтры")
        filter_layout = QVBoxLayout()

        # Настройка продавца
        seller_layout = QHBoxLayout()
        self.auto_filter_seller_check = QCheckBox("Продавец:")
        self.auto_filter_seller_check.setChecked(True)
        self.auto_filter_seller_combo = QComboBox()
        self.auto_filter_seller_combo.addItems(["1 ALTRA AUTO", "AutoTec", "TOP_DISKI"])
        seller_layout.addWidget(self.auto_filter_seller_check)
        seller_layout.addWidget(self.auto_filter_seller_combo)
        filter_layout.addLayout(seller_layout)

        # Настройка типа
        type_layout = QHBoxLayout()
        self.auto_filter_type_check = QCheckBox("Тип:")
        self.auto_filter_type_check.setChecked(True)
        self.auto_filter_type_combo = QComboBox()
        self.auto_filter_type_combo.addItems(["литые", "кованые", "сборные"])
        type_layout.addWidget(self.auto_filter_type_check)
        type_layout.addWidget(self.auto_filter_type_combo)
        filter_layout.addLayout(type_layout)

        # Настройка комплектации
        equipment_layout = QHBoxLayout()
        self.auto_filter_equipment_check = QCheckBox("Комплектация:")
        self.auto_filter_equipment_check.setChecked(True)
        self.auto_filter_equipment_combo = QComboBox()
        self.auto_filter_equipment_combo.addItems(["4 диска"])
        equipment_layout.addWidget(self.auto_filter_equipment_check)
        equipment_layout.addWidget(self.auto_filter_equipment_combo)
        filter_layout.addLayout(equipment_layout)

        filter_group.setLayout(filter_layout)
        ctl.addWidget(filter_group)

        self.filters_btn = QPushButton('Фильтры')
        self.filters_btn.clicked.connect(self.open_filters)
        ctl.addWidget(self.filters_btn)

        self.start_btn = QPushButton('Парсинг')
        self.start_btn.setEnabled(False)
        self.start_btn.clicked.connect(self.start_parsing)
        ctl.addWidget(self.start_btn)

        self.stop_btn = QPushButton('Стоп')
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self.stop_parsing)
        ctl.addWidget(self.stop_btn)

        self.export_btn = QPushButton('Экспорт в Excel')
        self.export_btn.setEnabled(False)
        self.export_btn.clicked.connect(self.export_to_excel)
        ctl.addWidget(self.export_btn)

        self.load_pvl_file_btn = QPushButton('Загрузить файл Павлодара')
        self.load_pvl_file_btn.clicked.connect(self.load_pvl_file)
        ctl.addWidget(self.load_pvl_file_btn)

        self.load_ast_file_btn = QPushButton('Загрузить файл Астаны')
        self.load_ast_file_btn.clicked.connect(self.load_ast_file)
        ctl.addWidget(self.load_ast_file_btn)

        self.remove_duplicates_btn = QPushButton('Удалить дубли')
        self.remove_duplicates_btn.setEnabled(False)
        self.remove_duplicates_btn.clicked.connect(self.remove_duplicates)
        ctl.addWidget(self.remove_duplicates_btn)

        l1.addLayout(ctl)

        # Индикатор прогресса
        pctl = QHBoxLayout()
        self.progress_bar = QProgressBar()
        self.progress_label = QLabel('Готов к работе')
        pctl.addWidget(self.progress_bar)
        pctl.addWidget(self.progress_label)
        l1.addLayout(pctl)

        # Таблица с товарами
        self.products_table = QTableWidget(0, 3)
        self.products_table.setHorizontalHeaderLabels(['URL', 'Код товара', 'Наименование'])
        self.products_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.products_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.products_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.products_table.cellDoubleClicked.connect(self.open_product_url)
        l1.addWidget(QLabel('Найденные товары:'))
        l1.addWidget(self.products_table)

        # Лог
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont('Monospace', 9))
        l1.addWidget(QLabel('Лог:'))
        l1.addWidget(self.log_text)

        tabs.addTab(w1, 'Парсер')

        # --- Comparison Tab ---
        tabs.addTab(ComparisonTab(), 'Сравнение')

        # --- Our Shop Tab ---
        self.our_shop_tab = OurShopTab()
        tabs.addTab(self.our_shop_tab, 'Наш магазин')

        self.statusBar().showMessage('Готов к работе')
        self.log('Приложение готово')

    def log(self, msg):
        self.log_text.append(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")
        self.log_text.verticalScrollBar().setValue(self.log_text.verticalScrollBar().maximum())

    def update_total_items(self, count):
        self.expected_total_items = count
        self.log(f"Всего товаров для обработки: {count}")

    def update_progress_bar(self, curr, total, page, found, tot, msg):
        if total > 0:
            p = int(curr / total * 100)
            self.progress_bar.setValue(p)

            # Обновляем текст с информацией о прогрессе
            if self.expected_total_items > 0:
                self.progress_label.setText(
                    f"Прогресс: {curr}/{total} ({p}%) | Проверено: {self.parser_thread.total_checked}/{self.expected_total_items} | С самовывозом: {tot}")
            else:
                self.progress_label.setText(f"Прогресс: {curr}/{total} ({p}%) | Всего: {tot}")

            self.statusBar().showMessage(f"Проверено: {curr}/{total}")
        if msg:
            self.log(msg)

    def update_table(self, prods):
        self.products_table.setRowCount(len(prods))
        for r, p in enumerate(prods):
            it = QTableWidgetItem(p['URL'])
            it.setData(Qt.UserRole, p['URL'])
            self.products_table.setItem(r, 0, it)
            self.products_table.setItem(r, 1, QTableWidgetItem(p['Код товара']))
            self.products_table.setItem(r, 2, QTableWidgetItem(p['Наименование']))

    def open_product_url(self, r, c):
        url = self.products_table.item(r, 0).data(Qt.UserRole)
        if url:
            QDesktopServices.openUrl(QUrl(url))

    def open_filters(self):
        url = self.url_combo.currentText()
        if not url.startswith('https://kaspi.kz'):
            self.log('Неверный URL')
            return

        code = self.city_combo.currentData()
        name = self.city_combo.currentText()

        self.log(f"Открытие: {url} ({name})")
        self.products = []
        self.products_table.setRowCount(0)

        # Настройки браузера
        opts = webdriver.ChromeOptions()
        # Отключаем обнаружение автоматизации
        opts.add_argument('--disable-blink-features=AutomationControlled')
        opts.add_argument('--start-maximized')
        opts.add_argument('--disable-images')  # Отключаем загрузку изображений для ускорения
        opts.add_argument(
            'user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.110 Safari/537.36')

        # Инициализация браузера - ЭКСПЕРИМЕНТ 4
        # Пробуем разные способы инициализации
        self.driver = None
        
        # Способ 1: Используем локальный chromedriver.exe (если есть)
        if os.path.exists('chromedriver.exe'):
            try:
                self.log("Попытка использовать локальный chromedriver.exe...")
                self.driver = webdriver.Chrome(service=Service('chromedriver.exe'), options=opts)
                self.log("✅ Браузер запущен через локальный chromedriver.exe")
            except Exception as e:
                self.log(f"❌ Ошибка локального chromedriver: {e}")
        
        # Способ 2: Используем системный ChromeDriver
        if not self.driver:
            try:
                self.log("Попытка использовать системный ChromeDriver...")
                self.driver = webdriver.Chrome(options=opts)
                self.log("✅ Браузер запущен через системный ChromeDriver")
            except Exception as e:
                self.log(f"❌ Ошибка системного ChromeDriver: {e}")
        
        # Способ 3: Используем webdriver-manager (требует интернет/VPN)
        if not self.driver:
            try:
                self.log("Попытка использовать webdriver-manager (требуется VPN)...")
                self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)
                self.log("✅ Браузер запущен через webdriver-manager")
            except Exception as e:
                self.log(f"❌ Ошибка webdriver-manager: {e}")
                self.log("💡 Включите VPN или скачайте chromedriver.exe вручную")
                return
        
        if not self.driver:
            self.log("❌ Не удалось запустить браузер. Включите VPN или установите chromedriver.exe")
            return

        # Устанавливаем куки для города
        self.driver.get(url)
        self.driver.add_cookie(
            {'name': 'kaspi.storefront.cookie.city', 'value': code, 'domain': 'kaspi.kz', 'path': '/'})
        self.driver.refresh()
        self.log(f"Город установлен: {name}")

        # Применяем автоматические фильтры, если выбрано
        filter_options = self.get_filter_options()

        if filter_options:
            self.log('Применяем автоматические фильтры')
            self.apply_filters_directly()
        else:
            self.log('Автоматические фильтры не выбраны, выберите фильтры вручную в браузере')

        self.filters_btn.setEnabled(False)
        self.start_btn.setEnabled(True)

    def apply_filters_directly(self):
        """Применяет выбранные фильтры непосредственно в браузере"""
        if not self.driver:
            self.log("Ошибка: браузер не запущен")
            return False

        try:
            self.log("Начинаю применение автоматических фильтров...")
            time.sleep(3)  # Увеличиваем время ожидания загрузки страницы с фильтрами

            # 1. Применяем фильтр "Продавцы"
            if self.auto_filter_seller_check.isChecked():
                seller_name = self.auto_filter_seller_combo.currentText()
                self.log(f"Устанавливаю фильтр продавца: {seller_name}")

                try:
                    # Пробуем найти блок фильтра продавцов разными способами
                    try:
                        seller_filter = WebDriverWait(self.driver, 10).until(
                            EC.presence_of_element_located((By.XPATH,
                                                            "//span[contains(@class, 'filters__filter-title') and contains(text(), 'Продавцы')]/.."))
                        )
                    except:
                        self.log("Использую альтернативный способ поиска фильтра продавца")
                        seller_filter = WebDriverWait(self.driver, 10).until(
                            EC.presence_of_element_located((By.XPATH,
                                                            "//div[contains(@class, 'filters__filter')]/span[contains(text(), 'Продавцы')]/.."))
                        )

                    # Скроллим к элементу
                    self.driver.execute_script("arguments[0].scrollIntoView(true);", seller_filter)
                    time.sleep(1)

                    # Проверяем, нужно ли развернуть список
                    try:
                        show_more = seller_filter.find_element(By.XPATH,
                                                               ".//span[contains(@class, 'filters__spoiler')]")
                        if show_more.is_displayed():
                            self.driver.execute_script("arguments[0].click();", show_more)
                            time.sleep(1)
                    except:
                        pass

                    # Находим и выбираем нужного продавца, используя разные стратегии поиска
                    try:
                        # Вариант 1: Поиск по точному тексту
                        seller_option = WebDriverWait(seller_filter, 10).until(
                            EC.presence_of_element_located((By.XPATH,
                                                            f".//span[contains(@class, 'filters__filter-row__description-label') and normalize-space()='{seller_name}']"))
                        )
                    except:
                        try:
                            # Вариант 2: Менее строгий поиск по содержимому
                            self.log("Использую альтернативный способ поиска продавца")
                            seller_option = WebDriverWait(seller_filter, 10).until(
                                EC.presence_of_element_located(
                                    (By.XPATH, f".//span[contains(text(), '{seller_name}')]"))
                            )
                        except:
                            # Крайний случай: ищем по всей странице
                            self.log("Поиск продавца по всей странице")
                            seller_option = WebDriverWait(self.driver, 10).until(
                                EC.presence_of_element_located((By.XPATH, f"//span[contains(text(), '{seller_name}')]"))
                            )

                    # Находим родительский элемент с чекбоксом и кликаем
                    try:
                        seller_row = seller_option.find_element(By.XPATH,
                                                                "./ancestor::div[contains(@class, 'filters__filter-row')]")
                        # Проверка активности
                        is_active = 'active' in seller_row.get_attribute('class')

                        if not is_active:
                            # Пробуем найти и кликнуть на чекбокс разными способами
                            try:
                                checkbox = seller_row.find_element(By.XPATH,
                                                                   ".//label[contains(@class, 'filters__filter-row__checkbox')]")
                                self.driver.execute_script("arguments[0].click();", checkbox)
                            except:
                                # Альтернативный способ: клик по строке целиком
                                self.driver.execute_script("arguments[0].click();", seller_row)

                            time.sleep(2)  # Увеличиваем время ожидания применения фильтра
                            self.log(f"Фильтр продавца установлен: {seller_name}")
                        else:
                            self.log(f"Фильтр продавца уже активен: {seller_name}")
                    except Exception as e:
                        self.log(f"Ошибка при клике на фильтр продавца: {str(e)}")
                        # Пробуем кликнуть напрямую на опцию как крайний вариант
                        self.driver.execute_script("arguments[0].click();", seller_option)
                        time.sleep(2)
                except Exception as e:
                    self.log(f"Ошибка при установке фильтра продавца: {str(e)}")

            # 2. Применяем фильтр "Тип" (литые)
            if self.auto_filter_type_check.isChecked():
                type_value = self.auto_filter_type_combo.currentText()
                self.log(f"Устанавливаю фильтр типа: {type_value}")

                try:
                    # Пробуем найти блок фильтра типа разными способами
                    try:
                        type_filter = WebDriverWait(self.driver, 10).until(
                            EC.presence_of_element_located((By.XPATH,
                                                            "//span[contains(@class, 'filters__filter-title') and contains(text(), 'Тип')]/.."))
                        )
                    except:
                        self.log("Использую альтернативный способ поиска фильтра типа")
                        type_filter = WebDriverWait(self.driver, 10).until(
                            EC.presence_of_element_located((By.XPATH,
                                                            "//div[contains(@class, 'filters__filter')]/span[contains(text(), 'Тип')]/.."))
                        )

                    # Скроллим к элементу
                    self.driver.execute_script("arguments[0].scrollIntoView(true);", type_filter)
                    time.sleep(1)

                    # Находим и выбираем нужный тип с использованием разных стратегий
                    try:
                        type_option = WebDriverWait(type_filter, 10).until(
                            EC.presence_of_element_located((By.XPATH,
                                                            f".//span[contains(@class, 'filters__filter-row__description-label') and normalize-space()='{type_value}']"))
                        )
                    except:
                        try:
                            self.log("Использую альтернативный способ поиска типа")
                            type_option = WebDriverWait(type_filter, 10).until(
                                EC.presence_of_element_located((By.XPATH, f".//span[contains(text(), '{type_value}')]"))
                            )
                        except:
                            self.log("Поиск типа по всей странице")
                            type_option = WebDriverWait(self.driver, 10).until(
                                EC.presence_of_element_located((By.XPATH, f"//span[contains(text(), '{type_value}')]"))
                            )

                    # Находим родительский элемент и кликаем на него
                    try:
                        type_row = type_option.find_element(By.XPATH,
                                                            "./ancestor::div[contains(@class, 'filters__filter-row')]")
                        is_active = 'active' in type_row.get_attribute('class')

                        if not is_active:
                            try:
                                checkbox = type_row.find_element(By.XPATH,
                                                                 ".//label[contains(@class, 'filters__filter-row__checkbox')]")
                                self.driver.execute_script("arguments[0].click();", checkbox)
                            except:
                                self.driver.execute_script("arguments[0].click();", type_row)

                            time.sleep(2)
                            self.log(f"Фильтр типа установлен: {type_value}")
                        else:
                            self.log(f"Фильтр типа уже активен: {type_value}")
                    except Exception as e:
                        self.log(f"Ошибка при клике на фильтр типа: {str(e)}")
                        # Пробуем кликнуть напрямую
                        self.driver.execute_script("arguments[0].click();", type_option)
                        time.sleep(2)
                except Exception as e:
                    self.log(f"Ошибка при установке фильтра типа: {str(e)}")

            # 3. Применяем фильтр "Комплектация" (4 диска)
            if self.auto_filter_equipment_check.isChecked():
                equipment_value = self.auto_filter_equipment_combo.currentText()
                self.log(f"Устанавливаю фильтр комплектации: {equipment_value}")

                try:
                    # Пробуем найти блок фильтра комплектации разными способами
                    try:
                        equipment_filter = WebDriverWait(self.driver, 10).until(
                            EC.presence_of_element_located((By.XPATH,
                                                            "//span[contains(@class, 'filters__filter-title') and contains(text(), 'Комплектация')]/.."))
                        )
                    except:
                        self.log("Использую альтернативный способ поиска фильтра комплектации")
                        equipment_filter = WebDriverWait(self.driver, 10).until(
                            EC.presence_of_element_located((By.XPATH,
                                                            "//div[contains(@class, 'filters__filter')]/span[contains(text(), 'Комплектация')]/.."))
                        )

                    # Скроллим к элементу
                    self.driver.execute_script("arguments[0].scrollIntoView(true);", equipment_filter)
                    time.sleep(1)

                    # Находим и выбираем нужную комплектацию
                    try:
                        equipment_option = WebDriverWait(equipment_filter, 10).until(
                            EC.presence_of_element_located((By.XPATH,
                                                            f".//span[contains(@class, 'filters__filter-row__description-label') and normalize-space()='{equipment_value}']"))
                        )
                    except:
                        try:
                            self.log("Использую альтернативный способ поиска комплектации")
                            equipment_option = WebDriverWait(equipment_filter, 10).until(
                                EC.presence_of_element_located(
                                    (By.XPATH, f".//span[contains(text(), '{equipment_value}')]"))
                            )
                        except:
                            self.log("Поиск комплектации по всей странице")
                            equipment_option = WebDriverWait(self.driver, 10).until(
                                EC.presence_of_element_located(
                                    (By.XPATH, f"//span[contains(text(), '{equipment_value}')]"))
                            )

                    # Находим родительский элемент и кликаем на него
                    try:
                        equipment_row = equipment_option.find_element(By.XPATH,
                                                                      "./ancestor::div[contains(@class, 'filters__filter-row')]")
                        is_active = 'active' in equipment_row.get_attribute('class')

                        if not is_active:
                            try:
                                checkbox = equipment_row.find_element(By.XPATH,
                                                                      ".//label[contains(@class, 'filters__filter-row__checkbox')]")
                                self.driver.execute_script("arguments[0].click();", checkbox)
                            except:
                                self.driver.execute_script("arguments[0].click();", equipment_row)

                            time.sleep(2)
                            self.log(f"Фильтр комплектации установлен: {equipment_value}")
                        else:
                            self.log(f"Фильтр комплектации уже активен: {equipment_value}")
                    except Exception as e:
                        self.log(f"Ошибка при клике на фильтр комплектации: {str(e)}")
                        # Пробуем кликнуть напрямую
                        self.driver.execute_script("arguments[0].click();", equipment_option)
                        time.sleep(2)
                except Exception as e:
                    self.log(f"Ошибка при установке фильтра комплектации: {str(e)}")

            # Ждем некоторое время после применения всех фильтров
            time.sleep(3)
            self.log("Автоматические фильтры успешно применены")
            return True

        except Exception as e:
            self.log(f"Общая ошибка при применении фильтров: {str(e)}")
            return False

    def get_filter_options(self):
        """Собирает выбранные опции фильтров"""
        filter_options = {}

        if self.auto_filter_seller_check.isChecked():
            filter_options['seller'] = self.auto_filter_seller_combo.currentText()

        if self.auto_filter_type_check.isChecked():
            filter_options['type'] = self.auto_filter_type_combo.currentText()

        if self.auto_filter_equipment_check.isChecked():
            filter_options['equipment'] = self.auto_filter_equipment_combo.currentText()

        return filter_options

    def start_parsing(self):
        if not hasattr(self, 'driver') or not self.driver:
            self.log('ОШИБКА: откройте браузер')
            return

        self.log('Старт парсинга')

        # Получаем опции фильтров (для передачи в поток)
        filter_options = self.get_filter_options()

        # Инициализация потока парсера с опциями фильтров
        # Поскольку фильтры уже применены при открытии браузера
        # передаем пустой словарь, чтобы не применять снова
        self.parser_thread = ParserThread(self.driver, {})

        # Подключение сигналов
        self.parser_thread.update_log.connect(self.log)
        self.parser_thread.update_progress.connect(self.update_progress_bar)
        self.parser_thread.finished_parsing.connect(self.parsing_finished)
        self.parser_thread.page_completed.connect(self.page_completed)
        self.parser_thread.update_total_items.connect(self.update_total_items)

        # Для сохранения информации о примененных фильтрах в экспорте
        self.parser_thread.auto_filter_options = filter_options

        # Запуск потока и обновление интерфейса
        self.parser_thread.start()
        self.driver = None  # Передаем управление драйвером в поток
        self.start_btn.setEnabled(False)
        self.filters_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)

    def stop_parsing(self):
        if self.parser_thread:
            self.parser_thread.stop()
            self.start_btn.setEnabled(True)
            self.stop_btn.setEnabled(False)

    def parsing_finished(self, products):
        self.products = products
        self.update_table(products)

        # Добавляем информацию о соотношении найденных и ожидаемых товаров
        if self.expected_total_items > 0:
            coverage = (self.parser_thread.total_checked / self.expected_total_items) * 100
            self.log(
                f"Парсинг завершен: проверено {self.parser_thread.total_checked} из {self.expected_total_items} ({coverage:.1f}%)")
            self.log(f"Найдено товаров с самовывозом: {len(products)}")
        else:
            self.log(f"Парсинг завершен: {len(products)} найдено")

        # Обновляем интерфейс
        self.filters_btn.setEnabled(True)
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(False)
        self.export_btn.setEnabled(bool(products))

        self.statusBar().showMessage(f"Завершено: {len(products)} товаров с самовывозом")

    def page_completed(self, page, found):
        # Обновляем таблицу по мере обработки данных
        if self.parser_thread:
            self.update_table(self.parser_thread.products)

    def export_to_excel(self):
        if not self.products:
            self.log('Нет данных')
            return

        # Формируем имя файла с городом и датой
        ts = datetime.now().strftime('%Y%m%d')
        city = self.city_combo.currentText()
        short = self.city_short.get(city, 'CITY')
        default = f"products_{short}_{ts}.xlsx"

        # Диалог сохранения
        path, _ = QFileDialog.getSaveFileName(self, 'Сохранить', default, 'Excel files (*.xlsx)')
        if path:
            try:
                # Сохраняем данные в Excel
                df = pd.DataFrame(self.products)
                # Упорядочиваем столбцы
                cols = [c for c in ['URL', 'Код товара', 'Наименование'] if c in df.columns]
                cols += [c for c in df.columns if c not in cols]
                df = df[cols]

                # Добавляем дополнительную информацию о парсинге
                with pd.ExcelWriter(path) as writer:
                    df.to_excel(writer, sheet_name='Товары', index=False)

                    # Добавляем лист с информацией о парсинге
                    if hasattr(self, 'parser_thread') and self.parser_thread:
                        # Получаем информацию о примененных фильтрах
                        filter_info = []
                        if self.parser_thread.auto_filter_options:
                            for k, v in self.parser_thread.auto_filter_options.items():
                                filter_info.append(f"{k}: {v}")
                        filter_str = ", ".join(filter_info) if filter_info else "Нет"

                        info = {
                            'Параметр': [
                                'Дата парсинга',
                                'Город',
                                'URL категории',
                                'Применённые фильтры',
                                'Всего товаров в выборке',
                                'Проверено товаров',
                                'Процент охвата',
                                'Найдено с самовывозом'
                            ],
                            'Значение': [
                                datetime.now().strftime('%Y-%m-%d %H:%M'),
                                city,
                                self.url_combo.currentText(),
                                filter_str,
                                self.expected_total_items,
                                self.parser_thread.total_checked,
                                f"{(self.parser_thread.total_checked / self.expected_total_items * 100):.1f}%" if self.expected_total_items > 0 else "Н/Д",
                                len(self.products)
                            ]
                        }
                        pd.DataFrame(info).to_excel(writer, sheet_name='Информация', index=False)

                self.log(f"Экспортировано: {path}")
            except Exception as e:
                self.log(f"Ошибка сохранения: {e}")

    def load_pvl_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Выберите файл Павлодара", "", "Excel files (*.xlsx *.xls)")
        if path:
            try:
                self.pvl_data = pd.read_excel(path)
                self.pvl_file_path = path
                self.log(f"Загружен файл Павлодара: {os.path.basename(path)}")
                self._check_duplicates_ready()
            except Exception as e:
                self.log(f"Ошибка при загрузке файла Павлодара: {e}")

    def load_ast_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Выберите файл Астаны", "", "Excel files (*.xlsx *.xls)")
        if path:
            try:
                self.ast_data = pd.read_excel(path)
                self.ast_file_path = path
                self.log(f"Загружен файл Астаны: {os.path.basename(path)}")
                self._check_duplicates_ready()
            except Exception as e:
                self.log(f"Ошибка при загрузке файла Астаны: {e}")

    def _check_duplicates_ready(self):
        if self.pvl_data is not None and self.ast_data is not None:
            self.remove_duplicates_btn.setEnabled(True)
            self.log("Оба файла загружены. Можно удалять дубли.")

    def remove_duplicates(self):
        if self.pvl_data is None or self.ast_data is None:
            self.log("Сначала загрузите оба файла")
            return
        
        # Находим дубли
        pvl_codes = set(self.pvl_data['Код товара'])
        ast_codes = set(self.ast_data['Код товара'])
        duplicates = pvl_codes & ast_codes
        
        self.log(f"Найдено дублей: {len(duplicates)}")
        
        # Удаляем дубли из Астаны
        ast_cleaned = self.ast_data[~self.ast_data['Код товара'].isin(duplicates)]
        
        # Сохраняем очищенный файл Астаны
        ts = datetime.now().strftime('%Y%m%d_%H%M')
        path, _ = QFileDialog.getSaveFileName(self, "Сохранить очищенный файл Астаны", 
                                            f"products_AST_cleaned_{ts}.xlsx", "Excel files (*.xlsx)")
        if path:
            try:
                with pd.ExcelWriter(path) as writer:
                    ast_cleaned.to_excel(writer, sheet_name='Товары', index=False)
                    
                    # Добавляем информацию о дублях
                    info = {
                        'Параметр': ['Дата очистки', 'Исходный файл Павлодара', 'Исходный файл Астаны', 
                                    'Найдено дублей', 'Товаров в Астане до очистки', 'Товаров в Астане после очистки'],
                        'Значение': [
                            datetime.now().strftime('%Y-%m-%d %H:%M'),
                            os.path.basename(self.pvl_file_path) if self.pvl_file_path else 'Н/Д',
                            os.path.basename(self.ast_file_path) if self.ast_file_path else 'Н/Д',
                            len(duplicates),
                            len(self.ast_data),
                            len(ast_cleaned)
                        ]
                    }
                    pd.DataFrame(info).to_excel(writer, sheet_name='Информация', index=False)
                    
                    # Лист с дублями
                    if duplicates:
                        dup_rows = []
                        for code in sorted(duplicates):
                            pvl_row = self.pvl_data[self.pvl_data['Код товара'] == code].iloc[0]
                            ast_row = self.ast_data[self.ast_data['Код товара'] == code].iloc[0]
                            dup_rows.append({
                                'Код товара': code,
                                'Наименование': pvl_row['Наименование'],
                                'URL (Павлодар)': pvl_row['URL'],
                                'URL (Астана)': ast_row['URL']
                            })
                        pd.DataFrame(dup_rows).to_excel(writer, sheet_name='Дубли', index=False)
                
                self.log(f"Очищенный файл сохранен: {path}")
                self.log(f"Удалено дублей: {len(duplicates)}")
            except Exception as e:
                self.log(f"Ошибка при сохранении: {e}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = KaspiParserGUI()
    win.showMaximized()
    sys.exit(app.exec_())