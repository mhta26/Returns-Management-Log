# Third-party components

This application uses dynamically linked PySide6 / Qt 6 under their applicable LGPL/GPL/commercial licenses, Qt PDF/PDFium and their bundled third-party notices. Qt shared libraries remain separate in the portable and installed packages so they can be replaced under the relevant license. Source and license information: https://doc.qt.io/qtforpython-6/licenses.html and https://www.qt.io/licensing/open-source-lgpl-obligations .

Other components: FastAPI (MIT), Starlette (BSD), Uvicorn (BSD), HTTPX (BSD), ReportLab (BSD), arabic-reshaper (MIT), python-bidi (LGPL), openpyxl (MIT), Pydantic (MIT), tzdata (Apache-2.0/IANA), and Python (PSF). Refer to upstream package distributions for exact terms and notices.

The user-supplied IBM Plex Sans Arabic Medium font (Copyright 2019 IBM Corp.) is distributed under SIL Open Font License 1.1 and includes its copyright and license in `returns_app/assets/FONT-LICENSE.txt`. The company logo and master data were supplied in the user's existing workbook and remain their materials. These are not licensed for unrelated redistribution by this project.

Build-time tools: PyInstaller (GPL with bootloader exception), Inno Setup (its own license), pytest (MIT), pypdf (BSD). Windows installers are not code-signed in this initial version.
