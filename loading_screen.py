from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLabel, QProgressBar
from PyQt5.QtCore import Qt, QTimer, pyqtSignal


class LoadingScreen(QWidget):
    """A loading screen widget that displays initialization progress."""

    finished = pyqtSignal()

    def __init__(self, version="v1.0.0"):
        super().__init__()
        self.version_string = version
        self.init_ui()
        self.progress = 0
        self.steps = []
        self.current_step = 0

    def init_ui(self):
        """Initialize the loading screen UI."""
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground, False)
        self.setStyleSheet("""
            QWidget {
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:1,
                    stop:0 #1a1a2e, stop:1 #16213e
                );
            }
        """)

        layout = QVBoxLayout()
        layout.setContentsMargins(50, 50, 50, 50)
        layout.setSpacing(20)

        # Title
        title = QLabel("SignSynth")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("""
            QLabel {
                color: #00d4ff;
                font-size: 48px;
                font-weight: bold;
                font-family: 'Segoe UI', Arial;
                margin-bottom: 10px;
            }
        """)
        layout.addWidget(title)

        # Subtitle
        subtitle = QLabel("YouTube Sign Language Integrator")
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet("""
            QLabel {
                color: #a0a0a0;
                font-size: 18px;
                font-family: 'Segoe UI', Arial;
                margin-bottom: 30px;
            }
        """)
        layout.addWidget(subtitle)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 2px solid #00d4ff;
                border-radius: 8px;
                background-color: #0f3460;
                text-align: center;
                color: white;
                font-size: 14px;
                font-weight: bold;
                height: 30px;
            }
            QProgressBar::chunk {
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:0,
                    stop:0 #00d4ff, stop:1 #0099cc
                );
                border-radius: 6px;
            }
        """)
        layout.addWidget(self.progress_bar)

        # Status label
        self.status_label = QLabel("Initializing...")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet("""
            QLabel {
                color: #ffffff;
                font-size: 14px;
                font-family: 'Segoe UI', Arial;
                margin-top: 10px;
            }
        """)
        layout.addWidget(self.status_label)

        # Detailed status
        self.detail_label = QLabel("")
        self.detail_label.setAlignment(Qt.AlignCenter)
        self.detail_label.setStyleSheet("""
            QLabel {
                color: #808080;
                font-size: 12px;
                font-family: 'Segoe UI', Arial;
                margin-top: 5px;
                min-height: 20px;
            }
        """)
        layout.addWidget(self.detail_label)

        # Add stretch to center content
        layout.addStretch()

        # Version info
        version = QLabel(self.version_string)
        version.setAlignment(Qt.AlignCenter)
        version.setStyleSheet("""
                    QLabel {
                        color: #606060;
                        font-size: 10px;
                        font-family: 'Segoe UI', Arial;
                    }
                """)
        layout.addWidget(version)

        self.setLayout(layout)

        # Set size
        self.setFixedSize(600, 400)

    def set_steps(self, steps):
        """Set the initialization steps to display."""
        self.steps = steps
        self.current_step = 0

    def update_progress(self, step_name, detail=""):
        """Update the progress bar and status."""
        if self.steps:
            self.current_step += 1
            progress_value = int((self.current_step / len(self.steps)) * 100)
            self.progress_bar.setValue(progress_value)

        self.status_label.setText(step_name)
        self.detail_label.setText(detail)

        # Force UI update
        QApplication.processEvents()

    def complete(self):
        """Mark loading as complete and close after a brief delay."""
        self.progress_bar.setValue(100)
        self.status_label.setText("Ready!")
        self.detail_label.setText("Starting application...")
        QApplication.processEvents()

        # Close after a short delay
        QTimer.singleShot(500, self.close_and_finish)

    def close_and_finish(self):
        """Close the loading screen and emit finished signal."""
        self.finished.emit()
        self.close()

    def center(self):
        """Center the window on the screen."""
        screen = QApplication.desktop().screenGeometry()
        size = self.geometry()
        self.move(
            (screen.width() - size.width()) // 2,
            (screen.height() - size.height()) // 2
        )


from PyQt5.QtWidgets import QApplication
