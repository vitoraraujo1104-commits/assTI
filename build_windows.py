"""Gera a distribuição Windows usando o Python do ambiente virtual."""
from pathlib import Path
from importlib import metadata
import json
import shutil
import subprocess
import sys

ROOT=Path(__file__).resolve().parent

def main():
    if sys.platform != 'win32':
        raise SystemExit('Gere o executável no Windows. Este script não faz compilação cruzada.')
    subprocess.run([sys.executable,'-m','unittest','discover','-s','tests','-v'],cwd=ROOT,check=True)
    command = [sys.executable, '-m', 'PyInstaller', '--noconfirm', '--clean',
               '--windowed', '--onedir', '--name', 'AssTI',
               '--collect-all', 'pymupdf']
    # O checkout original não inclui esses recursos; use-os quando fornecidos.
    if (ROOT / 'assets/assti.ico').is_file():
        command += ['--icon', str(ROOT / 'assets/assti.ico')]
    if (ROOT / 'assets').is_dir():
        command += ['--add-data', str(ROOT / 'assets') + ';assets']
    if (ROOT / 'version_info.txt').is_file():
        command += ['--version-file', str(ROOT / 'version_info.txt')]
    command.append(str(ROOT / 'main.py'))
    subprocess.run(command, cwd=ROOT, check=True)
    folder=ROOT/'dist'/'AssTI'
    if not (folder/'AssTI.exe').is_file():
        raise RuntimeError('O executável não foi criado.')
    for name in ('Criar_Atalho.cmd','Criar_Atalho.ps1','LEIA_PRIMEIRO.txt','TERCEIROS.txt'):
        if (ROOT/name).is_file():
            shutil.copy2(ROOT/name,folder/name)
    licenses=folder/'Licencas';licenses.mkdir(exist_ok=True)
    versions={}
    for name in ('PySide6','PySide6-Essentials','PySide6-Addons','shiboken6','openpyxl','et_xmlfile','defusedxml','pyinstaller','PyMuPDF'):
        try:dist=metadata.distribution(name)
        except metadata.PackageNotFoundError:continue
        versions[name]=dist.version
        for entry in dist.files or []:
            if any(word in str(entry).lower() for word in ('license','licence','copying','copyright')):
                source=Path(dist.locate_file(entry))
                if source.is_file():
                    dest=licenses/name/str(entry).replace('..','_')
                    dest.parent.mkdir(parents=True,exist_ok=True);shutil.copy2(source,dest)
    (folder/'versoes.json').write_text(json.dumps(versions,indent=2),encoding='utf-8')
    # Inclui o código desta versão para facilitar manutenção e auditoria.
    source_dir=folder/'Codigo_fonte';source_dir.mkdir(exist_ok=True)
    for name in ('main.py','core.py','pdf_core.py','pdf_ui.py','requirements.txt','build_windows.py','README.md'):
        shutil.copy2(ROOT/name,source_dir/name)
    shutil.copytree(ROOT/'tests', source_dir/'tests', ignore=shutil.ignore_patterns('__pycache__'))
    archive=shutil.make_archive(str(ROOT/'AssTI_Compartilhar'),'zip',root_dir=ROOT/'dist',base_dir='AssTI')
    print('\nPRONTO: '+str(folder/'AssTI.exe'))
    print('Para compartilhar: '+archive)
    print('Mantenha a pasta _internal junto do AssTI.exe.')

if __name__=='__main__':main()
