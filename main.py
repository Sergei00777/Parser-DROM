import sys
import os
import json
import requests
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QTextEdit, QPushButton,
    QVBoxLayout, QHBoxLayout, QFileDialog, QProgressBar, QMessageBox,
    QListWidget, QTabWidget, QLineEdit, QPlainTextEdit, QTableWidget,
    QTableWidgetItem, QDialog, QVBoxLayout
)
from PyQt5.QtCore import Qt
from bs4 import BeautifulSoup
from openpyxl import Workbook


class CarDetailsDialog(QDialog):
    def __init__(self, car_data, parent=None):
        super().__init__(parent)
        self.setWindowTitle("🚗 Детали объявления")
        self.setGeometry(100, 100, 450, 400)
        layout = QVBoxLayout()
        for key, value in car_data.items():
            label = QLabel(f"<b>{key}:</b> {value}")
            layout.addWidget(label)
        close_button = QPushButton("Закрыть")
        close_button.clicked.connect(self.accept)
        layout.addWidget(close_button)
        self.setLayout(layout)


class HelpTab(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout()
        title = QLabel("<h2>❓ Справка по программе</h2>")
        title.setStyleSheet("font-size: 16px; font-weight: bold;")
        layout.addWidget(title)
        instructions = QLabel("""
        <h3>Как использовать программу:</h3>
        <ol>
            <li><b>Вкладка '📥 Загрузка'</b>: Вставьте ссылки на Drom.ru (по одной на строку).</li>
            <li>Нажмите <i>➕ Добавить ссылки</i>, чтобы загрузить их в список.</li>
            <li>Выберите папку для сохранения HTML файлов с помощью кнопки <i>📁 Выбрать папку</i>.</li>
            <li>Нажмите <i>💾 Скачать и сохранить</i>, чтобы начать скачивание объявлений.</li>
            <li><b>Вкладка '🔍 Парсинг'</b>: Выберите один или несколько HTML-файлов, которые вы сохранили.</li>
            <li>Нажмите <i>📊 Показать результаты</i>, чтобы увидеть данные из объявлений.</li>
            <li>Результаты можно экспортировать в TXT или Excel.</li>
        </ol>
        """)
        instructions.setWordWrap(True)
        layout.addWidget(instructions)

        functions = QLabel("""
        <h3>Описание функций:</h3>
        <ul>
            <li><b>Добавить ссылки</b> — добавляет введённые URL в список.</li>
            <li><b>Сохранить HTML</b> — скачивает содержимое страниц и сохраняет в указанную папку.</li>
            <li><b>Парсинг</b> — извлекает информацию из HTML файлов (название, пробег, цена, двигатель, коробка передач).</li>
            <li><b>Показать детали</b> — открывает окно с полной информацией по выбранному автомобилю.</li>
            <li><b>Экспорт</b> — позволяет сохранить результаты в формате TXT или Excel.</li>
        </ul>
        """)
        functions.setWordWrap(True)
        layout.addWidget(functions)

        errors = QLabel("""
        <h3>Частые ошибки и их решение:</h3>
        <ul>
            <li><b>Ошибка: Не выбрана папка</b> — обязательно укажите путь к папке перед загрузкой.</li>
            <li><b>Ошибка: Файл не найден</b> — проверьте, что файл существует и имеет расширение .html.</li>
            <li><b>Ошибка: Неверный URL</b> — убедитесь, что ссылки корректны и ведут на drom.ru.</li>
            <li><b>Ошибка при парсинге</b> — файл может быть повреждён или иметь неправильный формат.</li>
        </ul>
        """)
        errors.setWordWrap(True)
        layout.addWidget(errors)
        layout.addStretch()
        self.setLayout(layout)


class ModernParserApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("🚗 Парсер объявлений Drom.ru")
        self.setGeometry(100, 100, 900, 700)
        self.setStyleSheet("""
            background-color: #2e2e2e;
            color: white;
            font-family: 'Segoe UI', sans-serif;
        """)
        # Переменные
        self.html_files_paths = []  # Несколько файлов для парсинга
        self.results = []
        self.downloaded_files = []  # Список скачанных файлов
        self.urls = []  # Список URL
        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout()
        # === Вкладки ===
        tabs = QTabWidget()
        tabs.setStyleSheet("""
            QTabBar::tab {
                background: #3e3e3e;
                color: white;
                padding: 10px;
                border-top-left-radius: 5px;
                border-top-right-radius: 5px;
            }
            QTabBar::tab:selected {
                background: #4CAF50;
            }
        """)
        download_tab = QWidget()
        parse_tab = QWidget()
        help_tab = HelpTab()  # Новая вкладка "Справка"
        tabs.addTab(download_tab, "📥 Загрузка")
        tabs.addTab(parse_tab, "🔍 Парсинг")
        tabs.addTab(help_tab, "❓ Справка")
        self.create_download_tab(download_tab)
        self.create_parse_tab(parse_tab)
        main_layout.addWidget(tabs)
        self.setLayout(main_layout)

    def create_download_tab(self, tab):
        layout = QVBoxLayout()
        self.url_input = QTextEdit()
        self.url_input.setPlaceholderText("Вставьте ссылки (по одной на строку)")
        self.url_input.setStyleSheet("background-color: #1e1e1e; color: white; border-radius: 5px;")
        self.add_button = QPushButton("➕ Добавить ссылки")
        self.url_list = QListWidget()
        self.url_list.setStyleSheet("background-color: #1e1e1e; color: white; border-radius: 5px;")
        self.folder_button = QPushButton("📁 Выбрать папку")
        self.folder_path = QLineEdit()
        self.folder_path.setPlaceholderText("Папка для сохранения")
        self.folder_path.setStyleSheet("background-color: #1e1e1e; color: white; border-radius: 5px;")
        self.save_all_button = QPushButton("💾 Скачать и сохранить")
        self.save_all_button.setStyleSheet("background-color: #4CAF50; color: white; padding: 10px; border-radius: 5px;")
        self.progress_bar_download = QProgressBar()
        self.progress_bar_download.setValue(0)
        self.progress_bar_download.setTextVisible(False)
        self.progress_bar_download.setStyleSheet("""
            QProgressBar {
                border: 1px solid #444;
                border-radius: 5px;
                background-color: #1e1e1e;
            }
            QProgressBar::chunk {
                background-color: #2196F3;
                width: 20px;
            }
        """)
        layout.addWidget(QLabel("🔗 Введите ссылки:"))
        layout.addWidget(self.url_input)
        layout.addWidget(self.add_button)
        layout.addWidget(QLabel("📋 Список ссылок:"))
        layout.addWidget(self.url_list)
        layout.addWidget(self.folder_button)
        layout.addWidget(self.folder_path)
        layout.addWidget(self.save_all_button)
        layout.addWidget(self.progress_bar_download)
        tab.setLayout(layout)
        # Подключение событий
        self.add_button.clicked.connect(self.add_urls)
        self.folder_button.clicked.connect(self.select_folder)
        self.save_all_button.clicked.connect(self.download_selected)

    def create_parse_tab(self, tab):
        layout = QVBoxLayout()
        self.load_single_file_button = QPushButton("📄 Выбрать один HTML-файл")
        self.load_multi_file_button = QPushButton("📂 Выбрать несколько HTML-файлов")
        self.clear_button = QPushButton("🧹 Очистить")
        self.show_results_button = QPushButton("📊 Показать результаты")
        self.save_txt_button = QPushButton("📄 Сохранить в TXT")
        self.save_excel_button = QPushButton("📘 Сохранить в Excel")
        self.output_table = QTableWidget()
        self.output_table.setColumnCount(5)
        self.output_table.setHorizontalHeaderLabels(["Название", "Пробег", "Цена", "Двигатель", "Коробка"])
        self.output_table.setStyleSheet("background-color: #1e1e1e; color: white; border-radius: 5px;")
        self.output_table.cellClicked.connect(self.show_car_details)
        self.progress_bar_parse = QProgressBar()
        self.progress_bar_parse.setValue(0)
        self.progress_bar_parse.setTextVisible(True)
        self.progress_bar_parse.setStyleSheet("""
            QProgressBar {
                border: 1px solid #444;
                border-radius: 5px;
                background-color: #121212;
                color: white;
            }
            QProgressBar::chunk {
                background-color: #2196F3;
                width: 20px;
            }
        """)
        parse_buttons_layout = QHBoxLayout()
        parse_buttons_layout.addWidget(self.load_single_file_button)
        parse_buttons_layout.addWidget(self.load_multi_file_button)
        parse_buttons_layout.addWidget(self.clear_button)
        parse_action_layout = QHBoxLayout()
        parse_action_layout.addWidget(self.show_results_button)
        parse_action_layout.addWidget(self.save_txt_button)
        parse_action_layout.addWidget(self.save_excel_button)
        layout.addLayout(parse_buttons_layout)
        layout.addLayout(parse_action_layout)
        layout.addWidget(self.output_table)
        layout.addWidget(self.progress_bar_parse)
        tab.setLayout(layout)
        # Стили кнопок
        for btn in [
            self.load_single_file_button, self.load_multi_file_button,
            self.clear_button, self.show_results_button,
            self.save_txt_button, self.save_excel_button
        ]:
            btn.setStyleSheet("""
                QPushButton {
                    background-color: #3e3e3e;
                    color: white;
                    padding: 8px;
                    border-radius: 5px;
                }
                QPushButton:hover {
                    background-color: #555555;
                }
            """)
        # Подключение событий
        self.load_single_file_button.clicked.connect(self.select_html_file)
        self.load_multi_file_button.clicked.connect(self.select_multiple_html_files)
        self.clear_button.clicked.connect(self.clear_parser_data)
        self.show_results_button.clicked.connect(self.parse_and_show)
        self.save_txt_button.clicked.connect(self.save_parsed_results_txt)
        self.save_excel_button.clicked.connect(self.save_parsed_results_excel)

    def add_urls(self):
        text = self.url_input.toPlainText().strip()
        if not text:
            return
        urls = [url.strip() for url in text.splitlines() if url.strip()]
        self.urls.extend(urls)
        self.url_list.clear()
        self.url_list.addItems(urls)

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Выберите папку для сохранения")
        if folder:
            self.folder_path.setText(folder)

    def download_selected(self):
        selected_items = [self.url_list.item(i).text() for i in range(self.url_list.count())]
        folder = self.folder_path.text().strip()
        if not selected_items:
            QMessageBox.warning(self, "Ошибка", "Нет ссылок для загрузки.")
            return
        if not folder:
            QMessageBox.warning(self, "Ошибка", "Выберите папку для сохранения.")
            return
        headers = {"User-Agent": "Mozilla/5.0"}
        self.downloaded_files.clear()
        self.progress_bar_download.setRange(0, len(selected_items))
        self.progress_bar_download.setValue(0)
        for idx, url in enumerate(selected_items, start=1):
            try:
                response = requests.get(url, headers=headers, timeout=10)
                response.raise_for_status()
                filename = f"{idx}_{url.split('/')[-2]}.html"
                full_path = os.path.join(folder, filename)
                with open(full_path, "w", encoding="utf-8") as f:
                    f.write(response.text)
                self.downloaded_files.append(full_path)
            except Exception as e:
                QMessageBox.warning(self, "Ошибка", f"Не удалось загрузить {url}: {e}")
            self.progress_bar_download.setValue(idx)
        QMessageBox.information(self, "Готово", f"Загружено {len(self.downloaded_files)} файлов.")

    def select_html_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите HTML файл", "", "HTML Files (*.html);;All Files (*)"
        )
        if file_path:
            self.html_files_paths = [file_path]
            self.output_table.setRowCount(0)
            self.parse_and_show()

    def select_multiple_html_files(self):
        file_paths, _ = QFileDialog.getOpenFileNames(
            self, "Выберите HTML файлы", "", "HTML Files (*.html);;All Files (*)"
        )
        if file_paths:
            self.html_files_paths = file_paths
            self.output_table.setRowCount(0)
            self.parse_and_show()

    def clear_parser_data(self):
        self.html_files_paths = []
        self.results = []
        self.output_table.setRowCount(0)
        self.progress_bar_parse.setValue(0)

    def parse_and_show(self):
        if not self.html_files_paths:
            QMessageBox.warning(self, "Ошибка", "❌ Не выбран ни один HTML-файл.")
            return
        self.results = []
        total_files = len(self.html_files_paths)
        self.progress_bar_parse.setRange(0, total_files * 2)
        self.progress_bar_parse.setValue(0)
        for idx, file_path in enumerate(self.html_files_paths, start=1):
            try:
                with open(file_path, "r", encoding="utf-8") as f:
                    html_content = f.read()
                soup = BeautifulSoup(html_content, "html.parser")
                script_tag = soup.find("script", {"data-drom-module": "bulls-list-auto"})
                if script_tag and script_tag.string:
                    data = json.loads(script_tag.string)
                    ads = self.find_bulls(data)
                    if ads and isinstance(ads, list):
                        for ad in ads:
                            if not isinstance(ad, dict):
                                continue
                            title = ad.get("title", "").strip()
                            price = ad.get("price")
                            attributes = [item["payload"] for item in ad.get("attributes", [])
                                          if item.get("type") == "plain"]

                            mileage = next((a for a in attributes if "км" in a), "Не указано")
                            engine = next((a for a in attributes if any(k in a.lower() for k in ["л", "л.с.", "литр", "турбодизель"])), "Не указано")
                            transmission = next((a for a in attributes if any(k in a.lower() for k in ["акпп", "автомат", "механика", "вариатор", "робот"])), "Не указано")

                            price_str = f"{int(price):,} ₽" if price else "Цена не указана"
                            self.results.append({
                                "title": title,
                                "mileage": mileage,
                                "price": price_str,
                                "engine": engine,
                                "transmission": transmission
                            })
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"❌ Ошибка при обработке файла {file_path}: {e}")
                return
            self.progress_bar_parse.setValue(idx)
        self.progress_bar_parse.setValue(total_files * 2)
        self.output_table.setRowCount(len(self.results))
        for row, car in enumerate(self.results):
            self.output_table.setItem(row, 0, QTableWidgetItem(car['title']))
            self.output_table.setItem(row, 1, QTableWidgetItem(car['mileage']))
            self.output_table.setItem(row, 2, QTableWidgetItem(car['price']))
            self.output_table.setItem(row, 3, QTableWidgetItem(car['engine']))
            self.output_table.setItem(row, 4, QTableWidgetItem(car['transmission']))

    def show_car_details(self, row, column):
        car_data = self.results[row]
        dialog = CarDetailsDialog(car_data, self)
        dialog.exec_()

    def save_parsed_results_txt(self):
        if not self.results:
            QMessageBox.warning(self, "Ошибка", "❌ Нет данных для сохранения.")
            return
        save_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить результаты как", "parsed_ads.txt", "Text Files (*.txt)"
        )
        if save_path:
            try:
                with open(save_path, "w", encoding="utf-8") as f:
                    for res in self.results:
                        line = f"{res['title']} | Пробег: {res['mileage']} | Цена: {res['price']} | Двигатель: {res['engine']} | Коробка: {res['transmission']}\n"
                        f.write(line)
                QMessageBox.information(self, "Успех", "✅ Результаты успешно сохранены в TXT.")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"❌ Ошибка при сохранении: {e}")

    def save_parsed_results_excel(self):
        if not self.results:
            QMessageBox.warning(self, "Ошибка", "❌ Нет данных для сохранения.")
            return
        save_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить результаты как", "parsed_ads.xlsx", "Excel Files (*.xlsx)"
        )
        if not save_path:
            return
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Объявления"
            ws.append(["Название", "Пробег", "Цена", "Двигатель", "Коробка"])
            for res in self.results:
                ws.append([res["title"], res["mileage"], res["price"], res["engine"], res["transmission"]])
            wb.save(save_path)
            QMessageBox.information(self, "Успех", "✅ Результаты успешно сохранены в Excel.")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"❌ Ошибка при сохранении в Excel: {e}")

    def find_bulls(self, data):
        if isinstance(data, dict):
            for key, value in data.items():
                if key == "bulls":
                    return value
                found = self.find_bulls(value)
                if found:
                    return found
        elif isinstance(data, list):
            for item in data:
                found = self.find_bulls(item)
                if found:
                    return found
        return None


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ModernParserApp()
    window.show()
    sys.exit(app.exec_())