#!/usr/bin/env python3
import os
import glob
import time
import win32print
from datetime import datetime

BACKUP_DIR = os.path.join(os.environ['APPDATA'], 'Notepad++', 'backup')
PRINTER_NAME = None
printed = 0

APERTURE_LOGO = [
    "       ,-:;//;:=,",
    "         . :H@@@MM@M#H/.,+%;,",
    "      ,/X+ +M@@M@MM%=,-%HMMM@X/,",
    "     -+@MM; $M@@MH+-,;XMMMM@MMMM@+-",
    "    ;@M@@M- XM@X;. -+XXXXXHHH@M@M#@/.",
    "  ,%MM@@MH ,@%=            .---=-=:=,.",
    "  -@#@@@MX .,              -%HX$$%%%+;",
    " =-./@M@M$                  .;@MMMM@MM:",
    " X@/ -$MM/                    .+MM@@@M$",
    ",@M@H: :@:                    . -X#@@@@-",
    ",@@@MMX, .                    /H- ;@M@M=",
    ".H@@@@M@+,                    %MM+..%#$.",
    " /MMMM@MMH/.                  XM@MH; -;",
    "  /%+%$XHH@$=              , .H@@@@MX,",
    "   .=--------.           -%H.,@@@@@MX,",
    "   .%MM@@@HHHXX$$$%+- .:$MMX -M@@MM%.",
    "     =XMMM@MM@MM#H;,-+HMM@M+ /MMMX=",
    "       =%@M@M#@$-.=$@MM@@@M; %M%=",
    "         ,:+$+-,/H#MMMMMMM@- -,",
    "               =++%%%%+/:-.  ",
]


def find_backup_file():
    matches = glob.glob(os.path.join(BACKUP_DIR, 'spike.txt@*'))
    return matches[0] if matches else None


def find_printer():
    printers = win32print.EnumPrinters(
        win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    )
    keywords = ['pos', 'thermal', 'receipt', 'netum', '58', '80', 'xprinter', 'epson tm']
    for p in printers:
        name = p[2].lower()
        if any(k in name for k in keywords):
            print(f"✓  Found receipt printer: {p[2]}")
            return p[2]
    default = win32print.GetDefaultPrinter()
    print(f"✓  Using default printer: {default}")
    return default


def print_receipt(text, number):
    ESC = b'\x1b'
    GS  = b'\x1d'
    now = datetime.now().strftime('%b %d  %I:%M %p')

    data = ESC + b'@'  # full reset

    # Logo, compressed font (Font B), centered
    data += ESC + b'M' + b'\x01'    # Font B (compressed)
    data += ESC + b'a' + b'\x01'    # center
    for line in APERTURE_LOGO:
        data += (line + '\n').encode('ascii', errors='replace')

    # Full reset back to normal font before text
    data += ESC + b'@'
    data += ESC + b'a' + b'\x01'    # center
    data += ESC + b'E' + b'\x01'    # bold
    data += b'APERTURE SCIENCE\n'
    data += ESC + b'E' + b'\x00'
    data += b'"We do what we must,\n'
    data += b' because we can"\n'
    data += b'--------------------------------\n'

    # Note text, centered, bold, normal font
    data += ESC + b'a' + b'\x01'    # center
    data += ESC + b'E' + b'\x01'    # bold
    data += (text.strip() + '\n').encode('utf-8', errors='replace')
    data += ESC + b'E' + b'\x00'

    # Footer
    data += b'--------------------------------\n'
    data += ESC + b'a' + b'\x01'
    data += (f'{now}  #{str(number).zfill(3)}\n').encode('ascii', errors='replace')

    # Feed and cut
    data += ESC + b'd' + bytes([4])
    data += GS + b'V' + b'\x41' + b'\x03'

    hPrinter = win32print.OpenPrinter(PRINTER_NAME)
    try:
        win32print.StartDocPrinter(hPrinter, 1, ('SPIKE', None, 'RAW'))
        win32print.StartPagePrinter(hPrinter)
        win32print.WritePrinter(hPrinter, data)
        win32print.EndPagePrinter(hPrinter)
        win32print.EndDocPrinter(hPrinter)
    finally:
        win32print.ClosePrinter(hPrinter)


def read_lines(filepath):
    with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
        raw = f.read()
    raw = raw.replace('\r\n', '\n').replace('\r', '\n')
    return raw.split('\n')


def watch():
    global printed

    watch_file = find_backup_file()
    if not watch_file:
        print("⚠  Could not find spike.txt backup file.")
        print("   Make sure spike.txt is open in Notepad++ with auto-save on.\n")
        while not watch_file:
            time.sleep(1)
            watch_file = find_backup_file()

    print(f"  Watching: {watch_file}\n")

    prev_lines = read_lines(watch_file)
    prev_count = len(prev_lines)
    last_mtime = os.path.getmtime(watch_file)
    last_watch_file = watch_file

    while True:
        time.sleep(0.4)
        try:
            current_file = find_backup_file()
            if current_file and current_file != last_watch_file:
                # Backup file was recreated, reset state to avoid stale counter
                print(f"  ↻ Backup file changed, re-syncing...")
                watch_file = current_file
                last_watch_file = current_file
                prev_lines = read_lines(watch_file)
                prev_count = len(prev_lines)
                last_mtime = os.path.getmtime(watch_file)
                continue

            mtime = os.path.getmtime(watch_file)
            if mtime == last_mtime:
                continue
            last_mtime = mtime

            curr_lines = read_lines(watch_file)
            curr_count = len(curr_lines)

            # If file shrank (user deleted lines), just re-sync without printing
            if curr_count <= prev_count:
                prev_lines = curr_lines
                prev_count = curr_count
                continue

            # New lines were added — print each completed one
            for i in range(prev_count - 1, curr_count - 1):
                text = curr_lines[i].strip()
                if not text:
                    continue
                printed += 1
                print(f"  Printing #{printed}: {text[:50]}")
                try:
                    print_receipt(text, printed)
                    print(f"  Printed.\n")
                except Exception as e:
                    print(f"  ✗ Print error: {e}\n")

            prev_lines = curr_lines
            prev_count = curr_count

        except Exception as e:
            print(f"  Error: {e}")


if __name__ == '__main__':
    print("=" * 40)
    print("  Press Ctrl+C to stop")
    print("=" * 40 + "\n")

    PRINTER_NAME = find_printer()

    try:
        watch()
    except KeyboardInterrupt:
        print("\n  Stopped. Bye!")
