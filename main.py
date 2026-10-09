"""AssTI — interface desktop Qt. Execute com python main.py."""
import json
import logging
import os
from pathlib import Path
import sys
from datetime import datetime
from decimal import Decimal
from PySide6.QtCore import Qt, QThread, Signal, QUrl, QStandardPaths
from PySide6.QtGui import QIcon, QDesktopServices, QColor, QFont, QPixmap
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QFrame, QLabel,
    QPushButton, QVBoxLayout, QHBoxLayout, QStackedWidget, QTableWidget,
    QTableWidgetItem, QHeaderView, QAbstractItemView, QFileDialog, QMessageBox,
    QProgressBar, QSpinBox, QPlainTextEdit, QTabWidget)
from core import read_invoice, export_invoices, brl
from pdf_ui import PDFTab

VERSION = '2.0.0'
ROOT = Path(getattr(sys, '_MEIPASS', Path(__file__).resolve().parent))
STYLE = '''
QWidget {font-family: "Segoe UI"; font-size: 14px; color: #182338;}
QMainWindow, QWidget#page {background: #ffffff;}
QFrame#sidebar {background: #f2f6fb; border-right: 1px solid #e3eaf3;}
QLabel#brand {font-size: 27px; font-weight: 700; color: #0067c0;}
QLabel#heading {font-size: 30px; font-weight: 650;}
QLabel#muted {color: #67778f;}
QLabel#subheading {font-size: 20px; font-weight: 600;}
QPushButton {background: white; border: 1px solid #cad5e5; border-radius: 7px; padding: 11px 18px;}
QPushButton:hover {background: #edf5ff; border-color: #89b5ef;}
QPushButton:pressed {background: #dcecff;}
QPushButton:disabled {color: #8d99aa; background: #eef1f5; border-color: #e2e7ee;}
QPushButton#primary {background: #0067c0; color: white; border: 1px solid #0067c0; font-weight: 600;}
QPushButton#primary:hover {background: #005bab;}
QPushButton#primary:disabled {background: #afc8e3; border-color: #afc8e3;}
QPushButton#nav {text-align: left; background: transparent; border: none; padding: 14px;}
QPushButton#nav:checked {background: #dcecff; color: #0067c0; font-weight: 600;}
QFrame#drop {background: #eff6ff; border: 1px dashed #91bffc; border-radius: 10px;}
QFrame#status {background: #f5f8fc; border-radius: 7px;}
QFrame#card {border: 1px solid #dae3ef; border-radius: 9px; background: white;}
QTableWidget {background: white; alternate-background-color: #f8fafd; border: none; gridline-color: #e9eef5; selection-background-color: #dcecff; selection-color: #172033;}
QTableView::item {padding: 8px;}
QHeaderView::section {background: #eff3f8; padding: 12px 8px; border: none; border-bottom: 1px solid #e1e7ef; font-weight: 600;}
QPlainTextEdit {background: #f8fafd; border: 1px solid #e6edf5; border-radius: 6px; padding: 6px; color: #52647e;}
QSpinBox {padding: 8px; border: 1px solid #cad5e5; border-radius: 5px; background: white;}
QProgressBar {border: none; background: #e9f1fc; border-radius: 3px; max-height: 5px;}
QProgressBar::chunk {background: #0067c0;}
'''


def label(text, name=None):
    w = QLabel(text)
    w.setWordWrap(True)
    w.setTextFormat(Qt.TextFormat.PlainText)
    if name:
        w.setObjectName(name)
    return w


def button(text, action, primary=False):
    w = QPushButton(text)
    w.setCursor(Qt.CursorShape.PointingHandCursor)
    if primary:
        w.setObjectName('primary')
    w.clicked.connect(action)
    return w


def table(headers):
    w = QTableWidget(0, len(headers))
    w.setHorizontalHeaderLabels(headers)
    w.verticalHeader().hide()
    w.verticalHeader().setDefaultSectionSize(42)
    w.setAlternatingRowColors(True)
    w.setShowGrid(False)
    w.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
    w.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
    w.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
    w.horizontalHeader().setStretchLastSection(True)
    return w


class DropArea(QFrame):
    files = Signal(list)

    def __init__(self):
        super().__init__()
        self.setObjectName('drop')
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls() and all(u.isLocalFile() for u in event.mimeData().urls()):
            event.acceptProposedAction()

    def dropEvent(self, event):
        self.files.emit([u.toLocalFile() for u in event.mimeData().urls() if u.isLocalFile()])
        event.acceptProposedAction()


class Worker(QThread):
    def __init__(self, function):
        super().__init__()
        self.function = function
        self.result = None
        self.error = None

    def run(self):
        try:
            self.result = self.function()
        except Exception as exc:
            logging.exception('Falha na operação')
            self.error = str(exc)


class MainWindow(QMainWindow):
    def __init__(self, data_dir=None):
        super().__init__()
        self.setWindowTitle('AssTI 2.0 — Extratores PDF e XML')
        self.setWindowIcon(QIcon(str(ROOT / 'assets' / 'assti.ico')))
        self.resize(1280, 850)
        self.setMinimumSize(1000, 710)
        self.entries = []
        self.worker = None
        self.data_dir = Path(data_dir or QStandardPaths.writableLocation(QStandardPaths.StandardLocation.AppLocalDataLocation))
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.settings = self.load_json('configuracoes.json', {'expected': 22})
        if not isinstance(self.settings, dict):
            self.settings = {'expected': 22}
        self.history = self.load_json('historico.json', [])
        if not isinstance(self.history, list):
            self.history = []
        self.history = [h for h in self.history if isinstance(h, dict) and all(k in h for k in ('date','count','path'))]
        shell = QWidget()
        outer = QHBoxLayout(shell)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)
        side = QFrame()
        side.setObjectName('sidebar')
        side.setFixedWidth(210)
        sl = QVBoxLayout(side)
        sl.setContentsMargins(18, 26, 18, 22)
        sl.addWidget(label('AssTI', 'brand'))
        sl.addSpacing(28)
        self.nav = []
        self.pages = QStackedWidget()
        for n, title in enumerate(('Importar XML', 'Histórico', 'Configurações')):
            b = button(title, lambda checked=False, i=n: self.switch_page(i))
            b.setObjectName('nav')
            b.setCheckable(True)
            self.nav.append(b)
            sl.addWidget(b)
        sl.addStretch()
        sl.addWidget(label('Desktop · ' + VERSION, 'muted'))
        outer.addWidget(side)
        outer.addWidget(self.pages, 1)
        self.setCentralWidget(shell)
        self.make_import()
        self.make_history()
        self.make_settings()
        self.switch_page(0)
        self.refresh()
        # Mantém toda a interface XML original dentro da segunda aba.
        xml_shell = self.takeCentralWidget()
        self.tabs = QTabWidget()
        self.pdf_tab = PDFTab()
        self.tabs.addTab(self.pdf_tab, 'Extrator de PDF / Chaves NFCom')
        self.tabs.addTab(xml_shell, 'Extrator de XML')
        self.setCentralWidget(self.tabs)

    def load_json(self, name, default):
        try:
            return json.loads((self.data_dir / name).read_text(encoding='utf-8'))
        except (OSError, ValueError):
            return default

    def save_json(self, name, value):
        p = self.data_dir / name
        temp = p.with_suffix('.tmp')
        temp.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding='utf-8')
        temp.replace(p)

    def new_page(self, title, subtitle):
        page = QWidget()
        page.setObjectName('page')
        layout = QVBoxLayout(page)
        layout.setContentsMargins(28, 24, 28, 24)
        layout.setSpacing(16)
        layout.addWidget(label(title, 'heading'))
        layout.addWidget(label(subtitle, 'muted'))
        self.pages.addWidget(page)
        return layout

    def make_import(self):
        layout = self.new_page('Extrator de faturas XML', 'Adicione seus arquivos, confira os dados e exporte para Excel.')
        self.drop = DropArea()
        dl = QVBoxLayout(self.drop)
        dl.setContentsMargins(20, 18, 20, 18)
        icon = QLabel()
        if (ROOT / 'assets' / 'assti.png').is_file():
            icon.setPixmap(QPixmap(str(ROOT / 'assets' / 'assti.png')).scaled(42, 42, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        dl.addWidget(icon)
        for text, name in [('Arraste seus arquivos XML aqui', 'subheading'), ('ou selecione os arquivos no computador', 'muted')]:
            l = label(text, name)
            l.setAlignment(Qt.AlignmentFlag.AlignCenter)
            dl.addWidget(l)
        self.add_btn = button('+  Adicionar XML', self.choose_files, True)
        dl.addWidget(self.add_btn, 0, Qt.AlignmentFlag.AlignCenter)
        self.drop.files.connect(self.import_files)
        layout.addWidget(self.drop)
        strip = QFrame()
        strip.setObjectName('status')
        st = QHBoxLayout(strip)
        st.setContentsMargins(14, 12, 14, 12)
        self.status = label('Pronto para adicionar arquivos', 'muted')
        self.total = label('')
        st.addWidget(self.status, 1)
        st.addWidget(self.total)
        layout.addWidget(strip)
        self.progress = QProgressBar()
        self.progress.setTextVisible(False)
        self.progress.setRange(0, 0)
        self.progress.hide()
        layout.addWidget(self.progress)
        card = QFrame()
        card.setObjectName('card')
        cl = QVBoxLayout(card)
        cl.setContentsMargins(14, 12, 14, 12)
        header = QHBoxLayout()
        header.addWidget(label('Prévia dos dados', 'subheading'), 1)
        self.remove_btn = button('Remover seleção', self.remove_selected)
        header.addWidget(self.remove_btn)
        cl.addLayout(header)
        self.grid = table(['Conta', 'Vencimento', 'Consumo', 'Serviço de gestão', 'Juros e multa', 'Total', 'Situação'])
        self.grid.itemSelectionChanged.connect(self.show_detail)
        cl.addWidget(self.grid)
        layout.addWidget(card, 1)
        self.details = QPlainTextEdit()
        self.details.setReadOnly(True)
        self.details.setMaximumHeight(82)
        self.details.setPlaceholderText('Selecione uma fatura para conferir os itens ou o motivo do erro.')
        layout.addWidget(self.details)
        footer = QHBoxLayout()
        footer.addWidget(label('Processamento local dos arquivos', 'muted'), 1)
        self.clear_btn = button('Limpar lista', self.clear_files)
        self.export_btn = button('↓  Exportar para Excel', self.export, True)
        footer.addWidget(self.clear_btn)
        footer.addWidget(self.export_btn)
        layout.addLayout(footer)

    def make_history(self):
        l = self.new_page('Histórico de exportações', 'Planilhas geradas neste computador.')
        self.history_grid = table(['Data', 'Faturas', 'Planilha'])
        self.history_grid.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        l.addWidget(self.history_grid, 1)
        l.addWidget(button('Abrir planilha selecionada', self.open_history), 0, Qt.AlignmentFlag.AlignRight)
        self.refresh_history()

    def make_settings(self):
        l = self.new_page('Configurações', 'Preferências do seu AssTI.')
        l.addWidget(label('Quantidade esperada de faturas por lote', 'subheading'))
        self.expected = QSpinBox()
        self.expected.setRange(0, 10000)
        try:
            self.expected.setValue(int(self.settings.get('expected', 22)))
        except (ValueError, TypeError):
            self.expected.setValue(22)
        self.expected.setFixedWidth(130)
        l.addWidget(self.expected)
        l.addWidget(label('Use 0 para desativar o aviso. Uma quantidade diferente não bloqueia a exportação.', 'muted'))
        l.addWidget(button('Salvar configuração', self.save_settings), 0, Qt.AlignmentFlag.AlignLeft)
        l.addSpacing(12)
        l.addWidget(label('Regras de classificação', 'subheading'))
        l.addWidget(label('• Serviço de Gestão → coluna separada.\n• Juros e multa → somados na mesma coluna.\n• Demais itens, inclusive Gestão e Autosserviço → consumo.\n• XML inválido, duplicidade ou total divergente bloqueiam a exportação.\n• Excel sem cores e sem formato de tabela.'))
        l.addWidget(label('Compatível com XML de NFCOM no formato da fatura fornecida. O detalhamento por telefone exige o PDF. A leitura do protocolo não consulta a situação atual na SEFAZ nem valida a assinatura digital.', 'muted'))
        l.addStretch()

    def save_settings(self):
        try:
            self.settings['expected'] = self.expected.value()
            self.save_json('configuracoes.json', self.settings)
            self.refresh()
            QMessageBox.information(self, 'AssTI', 'Configuração salva.')
        except OSError as exc:
            self.error(str(exc))

    def switch_page(self, index):
        self.pages.setCurrentIndex(index)
        for n, b in enumerate(self.nav):
            b.setChecked(n == index)

    def choose_files(self):
        paths, _ = QFileDialog.getOpenFileNames(self, 'Adicionar faturas XML', '', 'Arquivos XML (*.xml)')
        self.import_files(paths)

    def run_job(self, function, completed):
        if self.worker:
            return
        self.worker = Worker(function)
        for b in (self.add_btn, self.remove_btn, self.clear_btn, self.export_btn):
            b.setEnabled(False)
        self.drop.setAcceptDrops(False)
        self.progress.show()
        self.status.setText('Processando…')
        def finish():
            worker = self.worker
            self.worker = None
            self.progress.hide()
            self.drop.setAcceptDrops(True)
            for b in (self.add_btn, self.remove_btn, self.clear_btn):
                b.setEnabled(True)
            try:
                if worker.error:
                    self.error(worker.error)
                else:
                    completed(worker.result)
            finally:
                self.refresh()
                worker.deleteLater()
        self.worker.finished.connect(finish)
        self.worker.start()

    def import_files(self, paths):
        if self.worker:
            return
        known = {os.path.normcase(str(Path(e['path']).resolve())) for e in self.entries}
        new = []
        for p in paths:
            key = os.path.normcase(str(Path(p).resolve()))
            if key not in known:
                known.add(key)
                new.append(p)
        if not new:
            return
        def read():
            entries = []
            for p in new:
                try:
                    entries.append({'path': p, 'invoice': read_invoice(p), 'error': ''})
                except Exception as exc:
                    entries.append({'path': p, 'invoice': None, 'error': str(exc)})
            return entries
        self.run_job(read, lambda result: self.entries.extend(result))

    def refresh(self):
        seen = set()
        valid = []
        self.grid.setRowCount(len(self.entries))
        for n, e in enumerate(self.entries):
            i = e['invoice']
            error = e['error']
            if i:
                if i.key in seen:
                    error = 'Fatura duplicada: a chave de acesso já está neste lote. Remova uma das cópias.'
                else:
                    seen.add(i.key)
                    valid.append(i)
            e['display_error'] = error
            values = [i.account, i.due.strftime('%d/%m/%Y'), brl(i.consumption), brl(i.management), brl(i.interest), brl(i.total)] if i else ['—'] * 6
            values.append('Revisar arquivo' if error else 'Conferido')
            for c, value in enumerate(values):
                cell = QTableWidgetItem(value)
                cell.setToolTip(error or Path(e['path']).name)
                if error:
                    cell.setForeground(QColor('#b42318'))
                self.grid.setItem(n, c, cell)
        errors = sum(bool(e['display_error']) for e in self.entries)
        self.export_btn.setEnabled(bool(self.entries) and not errors and not self.worker)
        self.total.setText('Total: ' + brl(sum((i.total for i in valid), Decimal(0))))
        count = len(self.entries)
        status = f'{count} arquivos · {count-errors} conferidos'
        expected = self.expected.value()
        if errors:
            status += f' · {errors} para revisar'
        elif count and expected and count != expected:
            status += f' · Esperados: {expected}'
        self.status.setText('Processando…' if self.worker else status)
        self.show_detail()

    def show_detail(self):
        row = self.grid.currentRow()
        if not 0 <= row < len(self.entries):
            self.details.clear()
            return
        e = self.entries[row]
        info = Path(e['path']).name
        if e.get('display_error'):
            info += '\n' + e['display_error']
        elif e['invoice']:
            i = e['invoice']
            info += f' · Competência {i.period} · NFCOM {i.number}\n'
            info += '  |  '.join(f'{p.description}: {brl(p.value)} [{p.category}]' for p in i.items)
        self.details.setPlainText(info)

    def remove_selected(self):
        for row in sorted({i.row() for i in self.grid.selectedIndexes()}, reverse=True):
            self.entries.pop(row)
        self.refresh()

    def clear_files(self):
        self.entries.clear()
        self.refresh()

    def export(self):
        if self.worker or not self.entries or any(e.get('display_error') for e in self.entries):
            return
        path, _ = QFileDialog.getSaveFileName(self, 'Salvar planilha', f'AssTI_Faturas_{datetime.now():%Y-%m-%d_%H%M}.xlsx', 'Excel (*.xlsx)')
        if not path:
            return
        if not path.lower().endswith('.xlsx'):
            path += '.xlsx'
            if Path(path).exists() and QMessageBox.question(self, 'Substituir planilha?', 'Este arquivo já existe. Deseja substituí-lo?') != QMessageBox.StandardButton.Yes:
                return
        invoices = [e['invoice'] for e in self.entries]
        def completed(_):
            self.history.insert(0, {'date': datetime.now().strftime('%d/%m/%Y %H:%M'), 'count': len(invoices), 'path': path})
            note = ''
            try:
                self.save_json('historico.json', self.history)
            except OSError:
                note = '\nA planilha foi salva, mas não foi possível registrar o histórico.'
            self.refresh_history()
            QMessageBox.information(self, 'Exportação concluída', f'{len(invoices)} faturas exportadas.\n{path}{note}')
        self.run_job(lambda: export_invoices(invoices, path), completed)

    def refresh_history(self):
        self.history_grid.setRowCount(len(self.history))
        for row, h in enumerate(self.history):
            for col, value in enumerate((h['date'], h['count'], h['path'])):
                item = QTableWidgetItem(str(value))
                item.setToolTip(str(value))
                self.history_grid.setItem(row, col, item)

    def open_history(self):
        row = self.history_grid.currentRow()
        if row < 0:
            self.error('Selecione uma planilha no histórico.')
            return
        path = Path(self.history[row]['path'])
        if path.suffix.lower() != '.xlsx' or not path.is_file():
            self.error('A planilha foi movida ou não está disponível neste computador.')
            return
        if not QDesktopServices.openUrl(QUrl.fromLocalFile(str(path))):
            self.error('Não foi possível abrir a planilha no aplicativo padrão.')

    def error(self, message):
        QMessageBox.warning(self, 'AssTI', message)

    def closeEvent(self, event):
        if self.worker or self.pdf_tab.worker:
            event.ignore()
            self.error('Aguarde o processamento terminar antes de fechar.')
        else:
            event.accept()


def main():
    app = QApplication(sys.argv)
    app.setApplicationName('AssTI')
    app.setOrganizationName('AssTI')
    app.setStyle('Fusion')
    app.setStyleSheet(STYLE)
    try:
        window = MainWindow()
        logging.basicConfig(filename=window.data_dir / 'assti.log', level=logging.ERROR, encoding='utf-8')
        window.show()
        return app.exec()
    except Exception as exc:
        QMessageBox.critical(None, 'AssTI', 'Não foi possível iniciar o aplicativo.\n' + str(exc))
        return 1

if __name__ == '__main__':
    sys.exit(main())
