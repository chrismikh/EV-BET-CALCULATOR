from __future__ import annotations

import sys, os

from PyQt6.QtCore import QThread, QTimer
from PyQt6.QtWidgets import QApplication, QMessageBox

# Ensure project root is in sys.path so we can import from src
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

try:
    from src.database import DatabaseManager
except ModuleNotFoundError:
    from database import DatabaseManager

try:
    from src.main_window import MainWindow
except ModuleNotFoundError:
    from main_window import MainWindow

try:
    from src.dialogs import PreloadDialog
except ModuleNotFoundError:
    from dialogs import PreloadDialog

try:
    from src.workers import PreloadWorker
except ModuleNotFoundError:
    from workers import PreloadWorker

def main():
    app = QApplication(sys.argv)

    # Ensure the database is created
    db = DatabaseManager()
    db.create_tables()

    win = MainWindow()
    dlg = PreloadDialog()
    preload_thread = QThread()
    worker = PreloadWorker(win.db)
    worker.moveToThread(preload_thread)
    preload_thread.started.connect(worker.run)
    worker.progress.connect(dlg.update_progress)
    worker.status.connect(dlg.update_status)

    def done(ok: bool, msg: str):
        preload_thread.quit(); preload_thread.wait(); worker.deleteLater(); dlg.accept()
        if not ok:
            QMessageBox.critical(win, "Preload Failed", f"Failed to preload data:\n{msg}")
            win.close()
            return
        win.set_initial_cache(worker.cache)
        win.set_matchbet_data(worker.matchbet_data)
        win.show()
        QTimer.singleShot(0, lambda: win.refresh_data(force=True))

    worker.finished.connect(done)
    QTimer.singleShot(60000, lambda: done(False, "Preloading timed out") if preload_thread.isRunning() else None)
    preload_thread.start()
    dlg.exec()
    return app.exec()


if __name__ == '__main__':
    sys.exit(main())