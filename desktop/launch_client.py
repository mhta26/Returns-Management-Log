import sys
if __name__ == '__main__':
    if '--self-test' in sys.argv:
        import os,tempfile
        with tempfile.TemporaryDirectory() as folder:
            os.environ['QT_QPA_PLATFORM']='offscreen'
            os.environ['XDG_DATA_HOME']=folder
            os.environ['XDG_CACHE_HOME']=folder
            from PySide6.QtWidgets import QApplication
            from returns_app.ui import LoginDialog,STYLE
            app=QApplication([]);app.setApplicationName('ReturnsManagerSelfTest');app.setStyleSheet(STYLE)
            window=LoginDialog();window.show();app.processEvents()
            assert not window.grab().isNull()
        raise SystemExit(0)
    from returns_app.ui import main
    raise SystemExit(main())
