import sys
import sqlite3
import pandas as pd
import numpy as np
import seaborn as sns
import matplotlib
# ВАЖНО: Устанавливаем бэкенд ДО импорта pyplot
matplotlib.use('QtAgg')
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QTabWidget, QPushButton, QFileDialog, 
                             QComboBox, QLabel, QTextEdit, QMessageBox, QTableWidget,
                             QTableWidgetItem, QProgressBar, QSplitter)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont, QAction
import logging
from datetime import datetime
import os

# Настройка логирования
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class DataLoaderThread(QThread):
    """Поток для загрузки данных"""
    progress = pyqtSignal(int)
    finished = pyqtSignal(bool, str)
    log_message = pyqtSignal(str)
    
    def __init__(self, file_path):
        super().__init__()
        self.file_path = file_path
        self.df = None
        
    def run(self):
        try:
            self.log_message.emit(f"Начало загрузки файла: {self.file_path}")
            
            # Чтение Excel файла
            self.progress.emit(10)
            xl = pd.ExcelFile(self.file_path)
            
            self.progress.emit(30)
            sheets_data = {}
            
            # Чтение всех листов
            for i, sheet_name in enumerate(xl.sheet_names):
                self.progress.emit(30 + int(40 * i / len(xl.sheet_names)))
                df = pd.read_excel(self.file_path, sheet_name=sheet_name)
                sheets_data[sheet_name] = df
                self.log_message.emit(f"Загружен лист: {sheet_name} ({len(df)} строк)")
            
            self.progress.emit(80)
            self.df = sheets_data
            
            # Сохранение в SQLite
            self.save_to_sqlite()
            
            self.progress.emit(100)
            self.finished.emit(True, "Данные успешно загружены")
            
        except Exception as e:
            self.log_message.emit(f"Ошибка при загрузке: {str(e)}")
            self.finished.emit(False, str(e))
    
    def save_to_sqlite(self):
        """Сохранение данных в SQLite базу"""
        conn = sqlite3.connect('data_visualization.db')
        try:
            for sheet_name, df in self.df.items():
                # Заменяем пробелы в названиях столбцов
                df.columns = [col.replace(' ', '_') for col in df.columns]
                df.to_sql(sheet_name, conn, if_exists='replace', index=False)
            self.log_message.emit("Данные сохранены в SQLite базу")
        finally:
            conn.close()

class StatisticsTab(QWidget):
    """Вкладка со статистикой"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.init_ui()
        
    def init_ui(self):
        layout = QVBoxLayout()
        
        # Выбор таблицы
        table_layout = QHBoxLayout()
        table_layout.addWidget(QLabel("Выберите таблицу:"))
        self.table_combo = QComboBox()
        self.table_combo.currentTextChanged.connect(self.load_table_data)
        table_layout.addWidget(self.table_combo)
        table_layout.addStretch()
        
        # Кнопка обновления
        self.refresh_btn = QPushButton("Обновить")
        self.refresh_btn.clicked.connect(self.refresh_tables)
        table_layout.addWidget(self.refresh_btn)
        
        layout.addLayout(table_layout)
        
        # Таблица с данными
        self.data_table = QTableWidget()
        layout.addWidget(self.data_table)
        
        # Статистика
        self.stats_text = QTextEdit()
        self.stats_text.setMaximumHeight(200)
        layout.addWidget(self.stats_text)
        
        self.setLayout(layout)
    
    def refresh_tables(self):
        """Обновление списка таблиц"""
        self.table_combo.clear()
        try:
            conn = sqlite3.connect('data_visualization.db')
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
            tables = cursor.fetchall()
            self.table_combo.addItems([table[0] for table in tables])
            conn.close()
            self.parent.log_message("Список таблиц обновлен")
        except Exception as e:
            self.parent.log_message(f"Ошибка при загрузке таблиц: {str(e)}")
    
    def load_table_data(self, table_name):
        """Загрузка данных таблицы"""
        if not table_name:
            return
            
        try:
            conn = sqlite3.connect('data_visualization.db')
            df = pd.read_sql(f"SELECT * FROM {table_name}", conn)
            conn.close()
            
            # Отображение данных в таблице
            self.display_data_table(df)
            
            # Отображение статистики
            self.display_statistics(df)
            
            self.parent.log_message(f"Загружена таблица: {table_name}")
            
        except Exception as e:
            self.parent.log_message(f"Ошибка при загрузке данных: {str(e)}")
    
    def display_data_table(self, df):
        """Отображение данных в QTableWidget"""
        self.data_table.setRowCount(len(df))
        self.data_table.setColumnCount(len(df.columns))
        self.data_table.setHorizontalHeaderLabels(df.columns)
        
        # Ограничиваем количество отображаемых строк для производительности
        display_rows = min(1000, len(df))
        
        for i in range(display_rows):
            for j in range(len(df.columns)):
                value = str(df.iloc[i, j])
                item = QTableWidgetItem(value)
                self.data_table.setItem(i, j, item)
        
        if len(df) > display_rows:
            self.parent.log_message(f"Показано первых {display_rows} строк из {len(df)}")
    
    def display_statistics(self, df):
        """Отображение статистики"""
        stats_text = "=== СТАТИСТИКА ===\n\n"
        
        # Основная информация
        stats_text += f"Количество строк: {len(df)}\n"
        stats_text += f"Количество столбцов: {len(df.columns)}\n\n"
        
        # Информация о столбцах
        stats_text += "=== ИНФОРМАЦИЯ О СТОЛБЦАХ ===\n"
        for col in df.columns:
            stats_text += f"\n{col}:\n"
            stats_text += f"  Тип: {df[col].dtype}\n"
            stats_text += f"  Уникальных значений: {df[col].nunique()}\n"
            stats_text += f"  Пустых значений: {df[col].isnull().sum()}\n"
            
            if pd.api.types.is_numeric_dtype(df[col]):
                # Безопасное вычисление статистики для совместимости с numpy 2.3.1
                try:
                    stats_text += f"  Среднее: {df[col].mean():.2f}\n"
                    stats_text += f"  Медиана: {df[col].median():.2f}\n"
                    stats_text += f"  Стандартное отклонение: {df[col].std():.2f}\n"
                    stats_text += f"  Минимум: {df[col].min():.2f}\n"
                    stats_text += f"  Максимум: {df[col].max():.2f}\n"
                except Exception as e:
                    stats_text += f"  Ошибка вычисления статистики: {str(e)}\n"
        
        self.stats_text.setText(stats_text)

class CorrelationTab(QWidget):
    """Вкладка с корреляционными графиками"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.init_ui()
        
    def init_ui(self):
        layout = QVBoxLayout()
        
        # Выбор таблицы и столбцов
        control_layout = QHBoxLayout()
        control_layout.addWidget(QLabel("Таблица:"))
        self.table_combo = QComboBox()
        control_layout.addWidget(self.table_combo)
        
        control_layout.addWidget(QLabel("Столбец X:"))
        self.x_combo = QComboBox()
        control_layout.addWidget(self.x_combo)
        
        control_layout.addWidget(QLabel("Столбец Y:"))
        self.y_combo = QComboBox()
        control_layout.addWidget(self.y_combo)
        
        self.plot_btn = QPushButton("Построить график")
        self.plot_btn.clicked.connect(self.plot_correlation)
        control_layout.addWidget(self.plot_btn)
        
        layout.addLayout(control_layout)
        
        # Область для графика
        self.figure = Figure(figsize=(10, 6))
        self.canvas = FigureCanvas(self.figure)
        layout.addWidget(self.canvas)
        
        self.setLayout(layout)
        
        # Подключение сигналов
        self.table_combo.currentTextChanged.connect(self.update_columns)
    
    def refresh_tables(self):
        """Обновление списка таблиц"""
        self.table_combo.clear()
        try:
            conn = sqlite3.connect('data_visualization.db')
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
            tables = cursor.fetchall()
            self.table_combo.addItems([table[0] for table in tables])
            conn.close()
        except Exception as e:
            self.parent.log_message(f"Ошибка при загрузке таблиц: {str(e)}")
    
    def update_columns(self, table_name):
        """Обновление списка столбцов"""
        self.x_combo.clear()
        self.y_combo.clear()
        
        if not table_name:
            return
            
        try:
            conn = sqlite3.connect('data_visualization.db')
            df = pd.read_sql(f"SELECT * FROM {table_name} LIMIT 1", conn)
            conn.close()
            
            numeric_columns = [col for col in df.columns if pd.api.types.is_numeric_dtype(df[col])]
            self.x_combo.addItems(numeric_columns)
            self.y_combo.addItems(numeric_columns)
            
            if len(numeric_columns) >= 2:
                self.x_combo.setCurrentIndex(0)
                self.y_combo.setCurrentIndex(1)
            
        except Exception as e:
            self.parent.log_message(f"Ошибка при загрузке столбцов: {str(e)}")
    
    def plot_correlation(self):
        """Построение корреляционного графика"""
        table_name = self.table_combo.currentText()
        x_col = self.x_combo.currentText()
        y_col = self.y_combo.currentText()
        
        if not all([table_name, x_col, y_col]):
            QMessageBox.warning(self, "Ошибка", "Выберите таблицу и оба столбца")
            return
            
        try:
            conn = sqlite3.connect('data_visualization.db')
            df = pd.read_sql(f"SELECT {x_col}, {y_col} FROM {table_name}", conn)
            conn.close()
            
            # Удаляем строки с NaN значениями
            df = df.dropna()
            
            if len(df) == 0:
                QMessageBox.warning(self, "Ошибка", "Нет данных для построения графика")
                return
            
            # Очистка фигуры
            self.figure.clear()
            ax = self.figure.add_subplot(111)
            
            # Построение scatter plot с Seaborn
            sns.scatterplot(data=df, x=x_col, y=y_col, ax=ax, alpha=0.6)
            ax.set_title(f'Корреляция между {x_col} и {y_col}')
            ax.set_xlabel(x_col)
            ax.set_ylabel(y_col)
            
            # Добавление линии тренда
            if len(df) > 1:
                try:
                    # Используем np.float64 для совместимости с numpy 2.3.1
                    x_vals = df[x_col].astype(np.float64).values
                    y_vals = df[y_col].astype(np.float64).values
                    z = np.polyfit(x_vals, y_vals, 1)
                    p = np.poly1d(z)
                    ax.plot(x_vals, p(x_vals), "r--", alpha=0.8, linewidth=2)
                except Exception as polyfit_error:
                    self.parent.log_message(f"Ошибка построения линии тренда: {polyfit_error}")
            
            # Расчет корреляции
            try:
                correlation = df[x_col].corr(df[y_col])
                ax.text(0.05, 0.95, f'Корреляция: {correlation:.3f}', 
                       transform=ax.transAxes, 
                       bbox=dict(boxstyle="round", facecolor='wheat', alpha=0.8),
                       fontsize=12)
            except Exception as corr_error:
                self.parent.log_message(f"Ошибка расчета корреляции: {corr_error}")
            
            self.canvas.draw()
            self.parent.log_message(f"Построен график корреляции: {x_col} vs {y_col}")
            
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось построить график: {str(e)}")
            self.parent.log_message(f"Ошибка при построении графика: {str(e)}")

class HeatmapTab(QWidget):
    """Вкладка с тепловой картой"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.init_ui()
        
    def init_ui(self):
        layout = QVBoxLayout()
        
        # Управление
        control_layout = QHBoxLayout()
        control_layout.addWidget(QLabel("Таблица:"))
        self.table_combo = QComboBox()
        control_layout.addWidget(self.table_combo)
        
        self.plot_btn = QPushButton("Построить тепловую карту")
        self.plot_btn.clicked.connect(self.plot_heatmap)
        control_layout.addWidget(self.plot_btn)
        
        layout.addLayout(control_layout)
        
        # Область для графика
        self.figure = Figure(figsize=(10, 8))
        self.canvas = FigureCanvas(self.figure)
        layout.addWidget(self.canvas)
        
        self.setLayout(layout)
    
    def refresh_tables(self):
        """Обновление списка таблиц"""
        self.table_combo.clear()
        try:
            conn = sqlite3.connect('data_visualization.db')
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
            tables = cursor.fetchall()
            self.table_combo.addItems([table[0] for table in tables])
            conn.close()
        except Exception as e:
            self.parent.log_message(f"Ошибка при загрузке таблиц: {str(e)}")
    
    def plot_heatmap(self):
        """Построение тепловой карты"""
        table_name = self.table_combo.currentText()
        
        if not table_name:
            QMessageBox.warning(self, "Ошибка", "Выберите таблицу")
            return
            
        try:
            conn = sqlite3.connect('data_visualization.db')
            df = pd.read_sql(f"SELECT * FROM {table_name}", conn)
            conn.close()
            
            # Выбор только числовых столбцов
            numeric_df = df.select_dtypes(include=[np.number])
            
            if numeric_df.empty:
                QMessageBox.warning(self, "Ошибка", "В таблице нет числовых столбцов")
                return
            
            if len(numeric_df.columns) < 2:
                QMessageBox.warning(self, "Ошибка", "Нужно как минимум 2 числовых столбца для тепловой карты")
                return
            
            # Очистка фигуры
            self.figure.clear()
            ax = self.figure.add_subplot(111)
            
            # Построение тепловой карты
            correlation_matrix = numeric_df.corr()
            
            # Используем более совместимый метод построения heatmap
            im = ax.imshow(correlation_matrix.values, cmap='coolwarm', aspect='auto', vmin=-1, vmax=1)
            
            # Устанавливаем метки
            ax.set_xticks(np.arange(len(correlation_matrix.columns)))
            ax.set_yticks(np.arange(len(correlation_matrix.index)))
            ax.set_xticklabels(correlation_matrix.columns, rotation=45, ha='right')
            ax.set_yticklabels(correlation_matrix.index)
            
            # Добавляем аннотации
            for i in range(len(correlation_matrix.index)):
                for j in range(len(correlation_matrix.columns)):
                    text = ax.text(j, i, f'{correlation_matrix.iloc[i, j]:.2f}',
                                 ha="center", va="center", color="black", fontsize=8)
            
            ax.set_title(f'Тепловая карта корреляций - {table_name}')
            self.figure.colorbar(im, ax=ax)
            
            self.canvas.draw()
            self.parent.log_message(f"Построена тепловая карта для таблицы: {table_name}")
            
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось построить тепловую карту: {str(e)}")
            self.parent.log_message(f"Ошибка при построении тепловой карты: {str(e)}")

class LinePlotTab(QWidget):
    """Вкладка с линейными графиками"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.init_ui()
        
    def init_ui(self):
        layout = QVBoxLayout()
        
        # Управление
        control_layout = QHBoxLayout()
        control_layout.addWidget(QLabel("Таблица:"))
        self.table_combo = QComboBox()
        control_layout.addWidget(self.table_combo)
        
        control_layout.addWidget(QLabel("Столбец:"))
        self.column_combo = QComboBox()
        control_layout.addWidget(self.column_combo)
        
        self.plot_btn = QPushButton("Построить линейный график")
        self.plot_btn.clicked.connect(self.plot_line)
        control_layout.addWidget(self.plot_btn)
        
        layout.addLayout(control_layout)
        
        # Область для графика
        self.figure = Figure(figsize=(10, 6))
        self.canvas = FigureCanvas(self.figure)
        layout.addWidget(self.canvas)
        
        self.setLayout(layout)
        
        # Подключение сигналов
        self.table_combo.currentTextChanged.connect(self.update_columns)
    
    def refresh_tables(self):
        """Обновление списка таблиц"""
        self.table_combo.clear()
        try:
            conn = sqlite3.connect('data_visualization.db')
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
            tables = cursor.fetchall()
            self.table_combo.addItems([table[0] for table in tables])
            conn.close()
        except Exception as e:
            self.parent.log_message(f"Ошибка при загрузке таблиц: {str(e)}")
    
    def update_columns(self, table_name):
        """Обновление списка столбцов"""
        self.column_combo.clear()
        
        if not table_name:
            return
            
        try:
            conn = sqlite3.connect('data_visualization.db')
            df = pd.read_sql(f"SELECT * FROM {table_name} LIMIT 1", conn)
            conn.close()
            
            numeric_columns = [col for col in df.columns if pd.api.types.is_numeric_dtype(df[col])]
            self.column_combo.addItems(numeric_columns)
            
        except Exception as e:
            self.parent.log_message(f"Ошибка при загрузке столбцов: {str(e)}")
    
    def plot_line(self):
        """Построение линейного графика"""
        table_name = self.table_combo.currentText()
        column_name = self.column_combo.currentText()
        
        if not all([table_name, column_name]):
            QMessageBox.warning(self, "Ошибка", "Выберите таблицу и столбец")
            return
            
        try:
            conn = sqlite3.connect('data_visualization.db')
            df = pd.read_sql(f"SELECT {column_name} FROM {table_name}", conn)
            conn.close()
            
            # Удаляем NaN значения
            df = df.dropna()
            
            if len(df) == 0:
                QMessageBox.warning(self, "Ошибка", "Нет данных для построения графика")
                return
            
            # Очистка фигуры
            self.figure.clear()
            ax = self.figure.add_subplot(111)
            
            # Построение линейного графика
            ax.plot(df[column_name], linewidth=2, color='blue', alpha=0.7)
            ax.set_title(f'Линейный график: {column_name}')
            ax.set_xlabel('Индекс')
            ax.set_ylabel(column_name)
            ax.grid(True, alpha=0.3)
            
            # Добавляем статистику
            try:
                stats_text = f"Среднее: {df[column_name].mean():.2f}\nМедиана: {df[column_name].median():.2f}"
                ax.text(0.02, 0.98, stats_text, transform=ax.transAxes, verticalalignment='top',
                       bbox=dict(boxstyle="round", facecolor='wheat', alpha=0.8))
            except Exception as stats_error:
                self.parent.log_message(f"Ошибка расчета статистики: {stats_error}")
            
            self.canvas.draw()
            self.parent.log_message(f"Построен линейный график для столбца: {column_name}")
            
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось построить график: {str(e)}")
            self.parent.log_message(f"Ошибка при построении линейного графика: {str(e)}")

class LogTab(QWidget):
    """Вкладка с логом действий"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.init_ui()
        
    def init_ui(self):
        layout = QVBoxLayout()
        
        # Текстовое поле для лога
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont("Courier", 9))
        
        # Кнопки управления
        button_layout = QHBoxLayout()
        self.clear_btn = QPushButton("Очистить лог")
        self.clear_btn.clicked.connect(self.clear_log)
        
        self.save_btn = QPushButton("Сохранить лог")
        self.save_btn.clicked.connect(self.save_log)
        
        button_layout.addWidget(self.clear_btn)
        button_layout.addWidget(self.save_btn)
        button_layout.addStretch()
        
        layout.addLayout(button_layout)
        layout.addWidget(self.log_text)
        
        self.setLayout(layout)
    
    def add_log_message(self, message):
        """Добавление сообщения в лог"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] {message}"
        self.log_text.append(log_entry)
        
        # Автопрокрутка к последнему сообщению
        cursor = self.log_text.textCursor()
        cursor.movePosition(cursor.MoveOperation.End)
        self.log_text.setTextCursor(cursor)
    
    def clear_log(self):
        """Очистка лога"""
        self.log_text.clear()
        self.parent.log_message("Лог очищен")
    
    def save_log(self):
        """Сохранение лога в файл"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить лог", "application_log.txt", "Text Files (*.txt)"
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(self.log_text.toPlainText())
                self.add_log_message(f"Лог сохранен в: {file_path}")
                QMessageBox.information(self, "Успех", f"Лог сохранен в: {file_path}")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить лог: {str(e)}")

class MainWindow(QMainWindow):
    """Главное окно приложения"""
    def __init__(self):
        super().__init__()
        self.data_loader = None
        self.init_ui()
        
    def init_ui(self):
        self.setWindowTitle("Data Visualization App - PyQt6")
        self.setGeometry(100, 100, 1200, 800)
        
        # Центральный виджет
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Основной layout
        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)
        
        # Панель загрузки файлов
        file_layout = QHBoxLayout()
        self.load_btn = QPushButton("Загрузить Excel файл")
        self.load_btn.clicked.connect(self.load_excel_file)
        
        self.file_label = QLabel("Файл не загружен")
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        
        file_layout.addWidget(self.load_btn)
        file_layout.addWidget(self.file_label)
        file_layout.addWidget(self.progress_bar)
        file_layout.addStretch()
        
        main_layout.addLayout(file_layout)
        
        # Вкладки
        self.tabs = QTabWidget()
        
        # Создание вкладок
        self.stats_tab = StatisticsTab(self)
        self.correlation_tab = CorrelationTab(self)
        self.heatmap_tab = HeatmapTab(self)
        self.lineplot_tab = LinePlotTab(self)
        self.log_tab = LogTab(self)
        
        # Добавление вкладок
        self.tabs.addTab(self.stats_tab, "📊 Статистика")
        self.tabs.addTab(self.correlation_tab, "🔗 Корреляции")
        self.tabs.addTab(self.heatmap_tab, "🎨 Тепловая карта")
        self.tabs.addTab(self.lineplot_tab, "📈 Линейные графики")
        self.tabs.addTab(self.log_tab, "📝 Ход работы")
        
        main_layout.addWidget(self.tabs)
        
        # Создание меню
        self.create_menu()
        
        # Настройка стиля
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f5f5;
            }
            QTabWidget::pane {
                border: 2px solid #C2C7CB;
                background-color: white;
                border-radius: 5px;
            }
            QTabBar::tab {
                background-color: #E1E1E1;
                border: 1px solid #C4C4C3;
                padding: 8px 20px;
                border-top-left-radius: 5px;
                border-top-right-radius: 5px;
                margin-right: 2px;
                font-weight: bold;
            }
            QTabBar::tab:selected {
                background-color: #FFFFFF;
                border-bottom-color: white;
            }
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
            QPushButton:pressed {
                background-color: #0D47A1;
            }
            QPushButton:disabled {
                background-color: #BDBDBD;
                color: #757575;
            }
            QTableWidget {
                gridline-color: #d0d0d0;
                alternate-background-color: #f8f8f8;
            }
            QProgressBar {
                border: 1px solid #CCC;
                border-radius: 3px;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #4CAF50;
            }
        """)
        
        self.log_message("Приложение запущено и готово к работе")
    
    def create_menu(self):
        """Создание меню приложения"""
        menubar = self.menuBar()
        
        # Меню Файл
        file_menu = menubar.addMenu('Файл')
        
        load_action = QAction('Загрузить Excel файл', self)
        load_action.triggered.connect(self.load_excel_file)
        file_menu.addAction(load_action)
        
        exit_action = QAction('Выход', self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        # Меню Помощь
        help_menu = menubar.addMenu('Помощь')
        
        about_action = QAction('О программе', self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)
    
    def show_about(self):
        """Показать информацию о программе"""
        about_text = """
        Data Visualization App
        
        Версия 1.0
        Создано с использованием:
        - PyQt6 6.10.0
        - pandas 2.3.0
        - matplotlib 3.10.3
        - numpy 2.3.1
        - seaborn
        
        Приложение для визуализации и анализа данных из Excel файлов.
        """
        QMessageBox.about(self, "О программе", about_text)
    
    def load_excel_file(self):
        """Загрузка Excel файла"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите Excel файл", "", "Excel Files (*.xlsx *.xls)"
        )
        
        if file_path:
            self.file_label.setText(f"Загрузка: {os.path.basename(file_path)}")
            self.progress_bar.setVisible(True)
            self.load_btn.setEnabled(False)
            
            # Запуск потока загрузки
            self.data_loader = DataLoaderThread(file_path)
            self.data_loader.progress.connect(self.progress_bar.setValue)
            self.data_loader.finished.connect(self.on_data_loaded)
            self.data_loader.log_message.connect(self.log_message)
            self.data_loader.start()
    
    def on_data_loaded(self, success, message):
        """Обработка завершения загрузки данных"""
        self.progress_bar.setVisible(False)
        self.load_btn.setEnabled(True)
        
        if success:
            self.file_label.setText("Файл успешно загружен")
            self.refresh_all_tabs()
            QMessageBox.information(self, "Успех", message)
        else:
            self.file_label.setText("Ошибка загрузки")
            QMessageBox.critical(self, "Ошибка", message)
    
    def refresh_all_tabs(self):
        """Обновление всех вкладок после загрузки данных"""
        # Обновление комбобоксов во всех вкладках
        for tab in [self.stats_tab, self.correlation_tab, self.heatmap_tab, self.lineplot_tab]:
            tab.refresh_tables()
        self.log_message("Все вкладки обновлены")
    
    def log_message(self, message):
        """Добавление сообщения в лог"""
        self.log_tab.add_log_message(message)
        logger.info(message)

def main():
    # Настройка стиля matplotlib для лучшего отображения в PyQt
    try:
        plt.style.use('seaborn-v0_8')
    except:
        # Fallback для новых версий matplotlib
        plt.style.use('default')
    
    app = QApplication(sys.argv)
    app.setApplicationName("Data Visualization App")
    app.setApplicationVersion("1.0")
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())

if __name__ == '__main__':
    main()