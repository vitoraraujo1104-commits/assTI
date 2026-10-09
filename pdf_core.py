"""Extração local de NFCom. Nenhum documento é enviado a serviços externos."""
from dataclasses import dataclass
from pathlib import Path
import os
import re
import tempfile
import zipfile
from typing import Callable

import pymupdf
from openpyxl import Workbook

MAX_PDF_BYTES = 100 * 1024 * 1024
MAX_ZIP_BYTES = 2 * 1024 * 1024 * 1024
UF_CODES = set('11 12 13 14 15 16 17 21 22 23 24 25 26 27 28 29 31 32 33 35 41 42 43 50 51 52 53'.split())
DIGIT_RUN = re.compile(r'[0-9](?:[0-9 \t\r\n.\-]*[0-9])?')
KEY_LABEL = re.compile(r'chave\s*(?:de\s*)?acesso\s*[:：\-]?', re.I)
NFCOM_LABEL = re.compile(r'NF\s*COM|nota\s+fiscal[^\n]{0,60}comunica', re.I)


@dataclass(frozen=True)
class PDFSource:
    path: str
    member: str = ''

    @property
    def name(self):
        return f'{Path(self.path).name} / {self.member}' if self.member else Path(self.path).name

    def read(self):
        if self.member:
            with zipfile.ZipFile(self.path) as archive:
                info = archive.getinfo(self.member)
                if info.file_size > MAX_PDF_BYTES:
                    raise ValueError('PDF excede o limite de 100 MB.')
                with archive.open(info) as stream:
                    data = stream.read(MAX_PDF_BYTES + 1)
        else:
            with open(self.path, 'rb') as stream:
                data = stream.read(MAX_PDF_BYTES + 1)
        if len(data) > MAX_PDF_BYTES:
            raise ValueError('PDF excede o limite de 100 MB.')
        return data


@dataclass
class PDFResult:
    filename: str
    account: str = ''
    key: str = ''
    status: str = 'Sem chave'
    notes: str = ''
    valid: bool = False
    duplicate: bool = False
    document_id: str = ''


def sources_from_paths(paths):
    """Não descompacta ZIPs no disco; rejeita lotes compactados excessivos."""
    result = []
    for raw in paths:
        path = Path(raw).resolve()
        if path.is_dir():
            result.extend(PDFSource(str(p)) for p in sorted(path.rglob('*'))
                          if p.is_file() and p.suffix.lower() == '.pdf')
        elif path.suffix.lower() == '.pdf':
            result.append(PDFSource(str(path)))
        elif path.suffix.lower() == '.zip':
            with zipfile.ZipFile(path) as archive:
                infos = [i for i in archive.infolist() if not i.is_dir() and i.filename.lower().endswith('.pdf')]
                if sum(i.file_size for i in infos) > MAX_ZIP_BYTES:
                    raise ValueError('PDFs do ZIP excedem o limite descompactado de 2 GB.')
                if any(i.file_size > MAX_PDF_BYTES for i in infos):
                    raise ValueError('O ZIP contém PDF maior que 100 MB.')
                if len({i.filename for i in infos}) != len(infos):
                    raise ValueError('ZIP contém nomes de entradas duplicados.')
                result.extend(PDFSource(str(path), i.filename) for i in infos)
        else:
            raise ValueError('Selecione PDFs, pastas ou arquivos ZIP.')
    return list(dict.fromkeys(result))


def check_digit(body):
    total = sum(int(d) * (2 + i % 8) for i, d in enumerate(reversed(body)))
    remainder = total % 11
    return str(0 if remainder < 2 else 11 - remainder)


def valid_cnpj(value):
    if len(set(value)) == 1:
        return False
    for size, weights in ((12, (5,4,3,2,9,8,7,6,5,4,3,2)),
                          (13, (6,5,4,3,2,9,8,7,6,5,4,3,2))):
        remainder = sum(int(d)*w for d, w in zip(value[:size], weights)) % 11
        if int(value[size]) != (0 if remainder < 2 else 11-remainder):
            return False
    return True


def validate_key(key):
    if not re.fullmatch(r'[0-9]{44}', key):
        return False, f'Esperados 44 dígitos; encontrados {len(key)}.'
    if key[:2] not in UF_CODES or not 1 <= int(key[4:6]) <= 12:
        return False, 'UF ou mês de emissão inválido.'
    if key[20:22] != '62':
        return False, 'Modelo fiscal diferente de 62 (NFCom).'
    if not valid_cnpj(key[6:20]):
        return False, 'CNPJ do emitente inválido.'
    if int(key[25:34]) == 0 or key[34] not in '12':
        return False, 'Número da NFCom ou tipo de emissão inválido.'
    if key[-1] != check_digit(key[:-1]):
        return False, 'Dígito verificador inválido.'
    return True, ''


def extract_text(text):
    account_match = re.search(r'(?:n[úu]mero\s*(?:da\s*)?conta|n[º°o.]?\s*(?:da\s*)?conta|c[óo]digo\s*(?:do\s*)?cliente|conta\s*(?:Claro)?)\s*[:：\-]?\s*([0-9][0-9 .\-]{2,30})', text, re.I)
    account = re.sub(r'[^0-9]', '', account_match.group(1)) if account_match else ''
    candidates = []
    for run in DIGIT_RUN.finditer(text):
        raw = run.group()
        key = re.sub(r'[^0-9]', '', raw)
        preceding = text[max(0, run.start()-250):run.start()]
        following = text[run.end():run.end()+100]
        explicit = bool(KEY_LABEL.search(preceding[-80:]))
        # A rótulo deve terminar imediatamente antes do número, evitando
        # atribuir códigos de pagamento posteriores à chave anterior.
        if explicit:
            label = list(KEY_LABEL.finditer(preceding))[-1]
            explicit = not preceding[label.end():].strip(' \t\r\n:.-')
        nearby = bool(NFCOM_LABEL.search(preceding + following))
        if explicit or (nearby and len(key) == 44 and key[20:22] == '62'):
            # Várias chaves completas na mesma linha ou em linhas adjacentes.
            # Só separa números já delimitados: não corta sequências contínuas.
            separate = re.findall(r'(?<![0-9])[0-9]{44}(?![0-9])', raw)
            if len(separate) > 1 and ''.join(separate) == key:
                candidates.extend(separate)
            else:
                candidates.append(key)
    return account, candidates


def read_pdf(source, ocr=False, cancelled=lambda: False):
    notes = []
    parts = []
    with pymupdf.open(stream=source.read(), filetype='pdf') as document:
        if document.needs_pass:
            raise ValueError('PDF protegido por senha.')
        for page in document:
            if cancelled():
                break
            text = page.get_text('text', sort=True)
            if not text.strip():
                if ocr:
                    try:
                        text = page.get_text(textpage=page.get_textpage_ocr(language='por+eng', dpi=200, full=True))
                        notes.append('OCR utilizado; confira com o documento original.')
                    except Exception:
                        notes.append('OCR indisponível ou falhou; revisar página digitalizada.')
                else:
                    notes.append('Página sem texto; OCR opcional não ativado.')
            parts.append(text)
    # Analisa páginas separadamente: nunca junta números de páginas distintas.
    account = ''
    keys = []
    for part in parts:
        found_account, found_keys = extract_text(part)
        account = account or found_account
        keys.extend(found_keys)
    if not account:
        notes.append('Número da conta não identificado; revisar.')
    common = ' '.join(dict.fromkeys(notes))
    if not keys:
        return [PDFResult(source.name, account, notes=common or 'Nenhuma chave identificável; revisar.')]
    results = []
    for key in keys:
        valid, reason = validate_key(key)
        results.append(PDFResult(source.name, account, key, 'Válida' if valid else 'Inválida',
                                 ' '.join(filter(None, (reason, common))), valid))
    return results


def process_batch(sources, progress: Callable = lambda *_: None, cancelled=lambda: False, ocr=False):
    results = []
    seen = set()
    for index, source in enumerate(sources, 1):
        if cancelled():
            break
        try:
            rows = read_pdf(source, ocr, cancelled)
        except Exception as exc:
            # Registra apenas a categoria da falha, sem conteúdo financeiro.
            import logging
            logging.error('Falha de leitura PDF: %s', type(exc).__name__)
            rows = [PDFResult(source.name, status='Erro de leitura', notes='Arquivo ilegível, protegido ou fora dos limites. Revise o PDF.')]
        if cancelled():
            break
        for row in rows:
            row.document_id = str(index)
            if row.key:
                row.duplicate = row.key in seen
                seen.add(row.key)
                if row.duplicate:
                    row.status += ' · Duplicada'
                    row.notes = (row.notes + ' Chave já encontrada neste lote.').strip()
        results.extend(rows)
        progress(index, rows)
    return results


def summarize(results):
    ids = {r.document_id or r.filename for r in results}
    return {
        'Quantidade de PDFs processados': len(ids),
        'Quantidade de chaves encontradas': sum(bool(r.key) for r in results),
        'Quantidade de chaves válidas': sum(r.valid for r in results),
        'Quantidade de chaves inválidas': sum(bool(r.key) and not r.valid for r in results),
        'Quantidade de documentos sem chave': len({r.document_id or r.filename for r in results if r.status == 'Sem chave'}),
        'Quantidade de chaves duplicadas': sum(r.duplicate for r in results),
        'Quantidade de arquivos com erro de leitura': len({r.document_id or r.filename for r in results if r.status == 'Erro de leitura'}),
    }


def export_results(results, destination):
    if not results:
        raise ValueError('Não há resultados para exportar.')
    destination = Path(destination)
    wb = Workbook()
    sheet = wb.active
    sheet.title = 'Chaves NFCom'
    sheet.append(['Nome do PDF', 'Número da conta Claro', 'Chave de acesso NFCom', 'Quantidade de dígitos', 'Status', 'Observações'])
    for row in results:
        sheet.append([row.filename, row.account, row.key, len(row.key), row.status, row.notes])
    for cells in sheet:
        for cell in cells:
            if isinstance(cell.value, str):
                cell.data_type = 's'
                cell.number_format = '@'
    sheet.freeze_panes = 'A2'
    sheet.auto_filter.ref = sheet.dimensions
    for col, width in zip('ABCDEF', (40, 25, 52, 22, 26, 80)):
        sheet.column_dimensions[col].width = width
    summary = wb.create_sheet('Resumo')
    summary.append(['Indicador', 'Quantidade'])
    for pair in summarize(results).items():
        summary.append(pair)
    summary.column_dimensions['A'].width = 45
    summary.column_dimensions['B'].width = 20
    descriptor, temporary = tempfile.mkstemp(suffix='.xlsx', dir=destination.parent)
    os.close(descriptor)
    try:
        wb.save(temporary)
        os.replace(temporary, destination)
    finally:
        wb.close()
        if os.path.exists(temporary):
            os.unlink(temporary)
