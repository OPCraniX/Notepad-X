from pathlib import Path


project_dir = Path(SPECPATH)


def collect_files(relative_dir, pattern, destination):
    return [
        (str(source_path), destination)
        for source_path in sorted((project_dir / relative_dir).glob(pattern))
        if source_path.is_file()
    ]


datas = [
    (str(project_dir / 'Notepad-X-help.txt'), '.'),
    (str(project_dir / 'VERSION'), '.'),
    (str(project_dir / 'gfx' / 'Notepad-X.ico'), 'gfx'),
    (str(project_dir / 'gfx' / 'splash.png'), 'gfx'),
]
datas += collect_files('audio', '*.mp3', 'audio')
datas += collect_files('cfg/language', '*.yml', 'cfg/language')
datas += collect_files('cfg/themes', '*.json', 'cfg/themes')
datas.append((str(project_dir / 'cfg' / 'spellcheck' / 'en.json.gz'), 'cfg/spellcheck'))

analysis = Analysis(
    [str(project_dir / 'Notepad-X.py')],
    pathex=[str(project_dir)],
    binaries=[],
    datas=datas,
    hiddenimports=[
        'cryptography.hazmat.primitives.ciphers.aead',
        'idlelib.colorizer',
        'idlelib.percolator',
        'PIL.Image',
        'PIL.ImageTk',
        'spellchecker',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['numpy', 'psutil', 'requests'],
    noarchive=False,
    optimize=1,
)
pyz = PYZ(analysis.pure)

exe = EXE(
    pyz,
    analysis.scripts,
    analysis.binaries,
    analysis.datas,
    [],
    name='Notepad-X',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(project_dir / 'gfx' / 'Notepad-X.ico'),
    version=str(project_dir / 'Notepad-X.Version.File.txt'),
)
