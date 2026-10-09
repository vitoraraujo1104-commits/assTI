"""Aba PDF independente da interface e da lógica financeira XML."""
from pathlib import Path
from PySide6.QtCore import QThread, Signal
from PySide6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView,
    QFileDialog, QMessageBox, QProgressBar, QApplication, QCheckBox)
from pdf_core import sources_from_paths, process_batch, export_results


class PDFWorker(QThread):
    progress = Signal(int, list)
    failed = Signal(str)

    def __init__(self, sources, ocr):
        super().__init__()
        self.sources = list(sources)
        self.ocr = ocr
        self.error = ''

    def run(self):
        try:
            process_batch(self.sources, self.progress.emit, self.isInterruptionRequested, self.ocr)
        except Exception:
            self.error = 'Não foi possível concluir o lote. Os resultados já processados foram preservados.'
            self.failed.emit(self.error)


class PDFTab(QWidget):
    def __init__(self):
        super().__init__()
        self.sources = []
        self.results = []
        self.worker = None
        self.processed = 0
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        heading = QLabel('Extrator de PDF / Chaves NFCom')
        heading.setObjectName('heading')
        layout.addWidget(heading)
        intro = QLabel('Adicione faturas Claro. A extração ocorre localmente; documentos sem chave ficam disponíveis para revisão.')
        intro.setWordWrap(True)
        layout.addWidget(intro)
        row = QHBoxLayout()
        self.input_buttons = []
        for title, action in (('Adicionar PDFs', self.choose_pdfs), ('Selecionar pasta', self.choose_folder), ('Importar ZIP', self.choose_zip)):
            btn = QPushButton(title)
            btn.clicked.connect(action)
            self.input_buttons.append(btn)
            row.addWidget(btn)
        layout.addLayout(row)
        self.ocr = QCheckBox('Usar OCR em páginas sem texto (requer Tesseract com idiomas por e eng)')
        layout.addWidget(self.ocr)
        row = QHBoxLayout()
        self.start_btn = QPushButton('Iniciar extração')
        self.start_btn.setObjectName('primary')
        self.start_btn.clicked.connect(self.start)
        row.addWidget(self.start_btn)
        self.cancel_btn = QPushButton('Cancelar processamento')
        self.cancel_btn.clicked.connect(self.cancel)
        row.addWidget(self.cancel_btn)
        row.addStretch()
        layout.addLayout(row)
        self.counts = QLabel()
        self.counts.setWordWrap(True)
        layout.addWidget(self.counts)
        self.progress = QProgressBar()
        self.progress.setRange(0, 1)
        layout.addWidget(self.progress)
        self.grid = QTableWidget(0, 5)
        self.grid.setHorizontalHeaderLabels(['Nome do arquivo PDF', 'Número da conta', 'Chave de acesso NFCom', 'Status da validação', 'Observações'])
        self.grid.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.grid.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.grid.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.grid.setAlternatingRowColors(True)
        self.grid.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.grid.horizontalHeader().setStretchLastSection(True)
        for i, width in enumerate((200, 140, 380, 180, 300)):
            self.grid.setColumnWidth(i, width)
        layout.addWidget(self.grid, 1)
        row = QHBoxLayout()
        for title, action in (('Copiar selecionadas', self.copy_selected), ('Copiar todas as chaves válidas', self.copy_valid)):
            btn = QPushButton(title)
            btn.clicked.connect(action)
            row.addWidget(btn)
        layout.addLayout(row)
        row = QHBoxLayout()
        self.export_btn = QPushButton('Exportar Excel')
        self.export_btn.clicked.connect(self.export)
        row.addWidget(self.export_btn)
        self.clear_btn = QPushButton('Limpar resultados')
        self.clear_btn.clicked.connect(self.clear)
        row.addWidget(self.clear_btn)
        self.message = QLabel('Pronto para adicionar arquivos.')
        self.message.setWordWrap(True)
        layout.addLayout(row)
        layout.addWidget(self.message)
        self.refresh_controls()

    def add_paths(self, paths):
        if self.worker:
            return
        errors = 0
        for path in paths:
            try:
                for source in sources_from_paths([path]):
                    if source not in self.sources:
                        self.sources.append(source)
            except Exception:
                errors += 1
        self.message.setText(f'{len(self.sources)} PDFs recebidos.' + (f' {errors} entradas não puderam ser importadas; revise os arquivos ZIP ou permissões.' if errors else ''))
        self.refresh_controls()

    def choose_pdfs(self):
        paths, _ = QFileDialog.getOpenFileNames(self, 'Adicionar PDFs', '', 'PDF (*.pdf)')
        self.add_paths(paths)

    def choose_folder(self):
        path = QFileDialog.getExistingDirectory(self, 'Selecionar pasta de PDFs')
        if path:
            self.add_paths([path])

    def choose_zip(self):
        paths, _ = QFileDialog.getOpenFileNames(self, 'Importar ZIP', '', 'ZIP (*.zip)')
        self.add_paths(paths)

    def start(self):
        if self.worker or not self.sources:
            return
        self.results.clear()
        self.grid.setRowCount(0)
        self.processed = 0
        self.progress.setRange(0, len(self.sources))
        self.progress.setValue(0)
        self.worker = PDFWorker(self.sources, self.ocr.isChecked())
        self.worker.progress.connect(self.receive)
        self.worker.failed.connect(self.message.setText)
        self.worker.finished.connect(self.finished)
        self.message.setText('Processando…')
        self.refresh_controls()
        self.worker.start()

    def receive(self, count, results):
        self.processed = count
        self.results.extend(results)
        for result in results:
            row = self.grid.rowCount()
            self.grid.insertRow(row)
            for column, text in enumerate((result.filename, result.account, result.key, result.status, result.notes)):
                item = QTableWidgetItem(text)
                item.setToolTip(text)
                self.grid.setItem(row, column, item)
        self.progress.setValue(count)
        self.refresh_controls()

    def finished(self):
        worker = self.worker
        cancelled = worker.isInterruptionRequested()
        self.worker = None
        worker.deleteLater()
        self.message.setText(worker.error or ('Processamento cancelado; resultados concluídos preservados.' if cancelled else 'Extração concluída. Confira as observações e os documentos para revisão.'))
        self.refresh_controls()

    def cancel(self):
        if self.worker:
            self.worker.requestInterruption()
            self.cancel_btn.setEnabled(False)
            self.message.setText('Cancelando após a leitura em andamento…')

    def refresh_controls(self):
        busy = self.worker is not None
        for button in self.input_buttons + [self.clear_btn]:
            button.setEnabled(not busy)
        self.ocr.setEnabled(not busy)
        self.start_btn.setEnabled(bool(self.sources) and not busy)
        self.cancel_btn.setEnabled(busy)
        self.export_btn.setEnabled(bool(self.results) and not busy)
        valid = sum(r.valid for r in self.results)
        invalid = sum(bool(r.key) and not r.valid for r in self.results)
        errors = sum(r.status == 'Erro de leitura' for r in self.results)
        self.counts.setText(f'Recebidos: {len(self.sources)} · Processados: {self.processed} · Chaves válidas: {valid} · Inválidas: {invalid} · Erros de leitura: {errors}')

    def copy_rows(self, rows):
        keys = [self.results[i].key for i in rows if self.results[i].key]
        if not keys:
            self.message.setText('Selecione linhas que contenham chaves.')
            return
        QApplication.clipboard().setText('\n'.join(keys))
        self.message.setText(f'{len(keys)} chaves copiadas, uma por linha. Confira o status das linhas selecionadas.')

    def copy_selected(self):
        self.copy_rows(sorted({i.row() for i in self.grid.selectedIndexes()}))

    def copy_valid(self):
        self.copy_rows([i for i, r in enumerate(self.results) if r.valid])

    def export(self):
        path, _ = QFileDialog.getSaveFileName(self, 'Exportar chaves NFCom', 'AssTI_Chaves_NFCom.xlsx', 'Excel (*.xlsx)')
        if not path:
            return
        if not path.lower().endswith('.xlsx'):
            path += '.xlsx'
            if Path(path).exists() and QMessageBox.question(self, 'Substituir planilha?', 'Este arquivo já existe. Deseja substituí-lo?') != QMessageBox.StandardButton.Yes:
                return
        try:
            export_results(self.results, path)
            self.message.setText('Planilha exportada com resultados e resumo.')
        except Exception:
            QMessageBox.warning(self, 'AssTI', 'Não foi possível salvar a planilha. Confira permissões e se o arquivo está aberto.')

    def clear(self):
        if self.worker:
            return
        self.sources.clear()
        self.results.clear()
        self.grid.setRowCount(0)
        self.processed = 0
        self.progress.setValue(0)
        self.message.setText('Lista limpa. Adicione um novo lote.')
        self.refresh_controls()
